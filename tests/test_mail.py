"""Tests für MailService — Graph-API-Aufrufe."""

import pytest
from unittest.mock import AsyncMock, MagicMock, patch

import aiohttp

from m365_connector.auth import M365Auth
from m365_connector.mail import MailService


@pytest.fixture
def auth():
    auth = MagicMock(spec=M365Auth)
    auth.get_token = AsyncMock(return_value="fake-token")
    return auth


@pytest.fixture
def mock_session():
    return MagicMock(spec=aiohttp.ClientSession)


@pytest.fixture
def service(auth, mock_session):
    return MailService(auth, lambda: mock_session)


def _make_mock_response(status: int, json_data: dict | None = None, text: str = ""):
    mock_resp = AsyncMock()
    mock_resp.status = status
    mock_resp.json = AsyncMock(return_value=json_data or {})
    mock_resp.text = AsyncMock(return_value=text)
    mock_resp.__aenter__ = AsyncMock(return_value=mock_resp)
    mock_resp.__aexit__ = AsyncMock(return_value=False)
    return mock_resp


async def test_send_mail_success(service, mock_session):
    mock_session.post = MagicMock(return_value=_make_mock_response(202))
    await service.send(mailbox="bot@firma.de", to="user@firma.de", subject="Test", body="<p>Hi</p>")
    mock_session.post.assert_called_once()


async def test_send_mail_raises_on_error(service, mock_session):
    mock_session.post = MagicMock(return_value=_make_mock_response(403, text='{"error":"Forbidden"}'))
    with pytest.raises(RuntimeError, match="mail.send failed"):
        await service.send(mailbox="bot@firma.de", to="user@firma.de", subject="Test", body="body")


async def test_list_inbox_returns_messages(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(200, {
        "value": [
            {"id": "1", "subject": "Test 1", "isRead": False},
            {"id": "2", "subject": "Test 2", "isRead": True},
        ]
    }))
    messages = await service.list_inbox(mailbox="bot@firma.de", limit=10)
    assert len(messages) == 2
    assert messages[0]["subject"] == "Test 1"


async def test_list_inbox_raises_on_error(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(401, text="Unauthorized"))
    with pytest.raises(RuntimeError, match="mail.list_inbox failed"):
        await service.list_inbox(mailbox="bot@firma.de")


async def test_session_reused(auth, mock_session):
    """Session-Callable wird pro Aufruf einmal aufgerufen — kein neues Session-Objekt."""
    call_count = 0

    def get_session():
        nonlocal call_count
        call_count += 1
        return mock_session

    mock_session.get = MagicMock(return_value=_make_mock_response(200, {"value": []}))
    svc = MailService(auth, get_session)
    await svc.list_inbox("box@firma.de")
    await svc.list_inbox("box@firma.de")
    assert call_count == 2  # get_session aufgerufen, aber dasselbe Objekt zurück


# ===== list_messages mit Pagination =====


async def test_list_messages_first_page(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(200, {
        "value": [{"id": "1", "subject": "A"}, {"id": "2", "subject": "B"}],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/bot/messages?$skiptoken=xyz",
    }))
    messages, next_token = await service.list_messages(mailbox="bot@firma.de", folder="inbox", limit=2)
    assert len(messages) == 2
    assert next_token == "https://graph.microsoft.com/v1.0/users/bot/messages?$skiptoken=xyz"


async def test_list_messages_last_page_returns_none_token(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(200, {
        "value": [{"id": "9"}],
    }))
    messages, next_token = await service.list_messages(mailbox="bot@firma.de", folder="archive")
    assert len(messages) == 1
    assert next_token is None


async def test_list_messages_uses_page_token_as_url(service, mock_session):
    """page_token muss als nextLink-URL direkt verwendet werden, keine eigenen Query-Params."""
    mock_session.get = MagicMock(return_value=_make_mock_response(200, {"value": []}))
    await service.list_messages(
        mailbox="bot@firma.de",
        page_token="https://graph.microsoft.com/v1.0/users/bot/messages?$skiptoken=abc",
    )
    call_args = mock_session.get.call_args
    assert call_args.args[0] == "https://graph.microsoft.com/v1.0/users/bot/messages?$skiptoken=abc"
    assert call_args.kwargs["params"] == {}


async def test_list_messages_without_folder_targets_all_messages(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(200, {"value": []}))
    await service.list_messages(mailbox="bot@firma.de")
    url = mock_session.get.call_args.args[0]
    assert url == "https://graph.microsoft.com/v1.0/users/bot@firma.de/messages"


async def test_list_messages_unread_filter(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(200, {"value": []}))
    await service.list_messages(mailbox="bot@firma.de", folder="inbox", unread_only=True)
    params = mock_session.get.call_args.kwargs["params"]
    assert params["$filter"] == "isRead eq false"


async def test_list_messages_raises_on_error(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(500, text="Server Error"))
    with pytest.raises(RuntimeError, match="mail.list_messages failed"):
        await service.list_messages(mailbox="bot@firma.de")


# ===== fetch_attachments =====


async def test_fetch_attachments_returns_list(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(200, {
        "value": [
            {"id": "att1", "name": "invoice.pdf", "contentBytes": "..."},
            {"id": "att2", "name": "image.png", "contentBytes": "..."},
        ]
    }))
    atts = await service.fetch_attachments(mailbox="bot@firma.de", message_id="msg-1")
    assert len(atts) == 2
    assert atts[0]["name"] == "invoice.pdf"


async def test_fetch_attachments_empty(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(200, {"value": []}))
    atts = await service.fetch_attachments(mailbox="bot@firma.de", message_id="msg-1")
    assert atts == []


async def test_fetch_attachments_raises_on_error(service, mock_session):
    mock_session.get = MagicMock(return_value=_make_mock_response(404, text="Not Found"))
    with pytest.raises(RuntimeError, match="mail.fetch_attachments failed"):
        await service.fetch_attachments(mailbox="bot@firma.de", message_id="missing")


# ===== send_forward =====


async def test_send_forward_single_recipient(service, mock_session):
    mock_session.post = MagicMock(return_value=_make_mock_response(202))
    await service.send_forward(
        mailbox="bot@firma.de",
        message_id="msg-1",
        to="empfaenger@firma.de",
        comment="FYI",
    )
    call_args = mock_session.post.call_args
    assert call_args.args[0].endswith("/messages/msg-1/forward")
    body = call_args.kwargs["json"]
    assert body["comment"] == "FYI"
    assert body["toRecipients"][0]["emailAddress"]["address"] == "empfaenger@firma.de"


async def test_send_forward_multiple_recipients(service, mock_session):
    mock_session.post = MagicMock(return_value=_make_mock_response(202))
    await service.send_forward(
        mailbox="bot@firma.de",
        message_id="msg-1",
        to=["a@firma.de", "b@firma.de"],
    )
    body = mock_session.post.call_args.kwargs["json"]
    assert len(body["toRecipients"]) == 2
    assert "comment" not in body


async def test_send_forward_raises_on_error(service, mock_session):
    mock_session.post = MagicMock(return_value=_make_mock_response(403, text="Forbidden"))
    with pytest.raises(RuntimeError, match="mail.send_forward failed"):
        await service.send_forward(mailbox="bot@firma.de", message_id="msg-1", to="x@y.de")


# ===== move_batch =====


async def test_move_batch_builds_batch_request(service, mock_session):
    mock_session.post = MagicMock(return_value=_make_mock_response(200, {
        "responses": [
            {"id": "0", "status": 201, "body": {"id": "new-1"}},
            {"id": "1", "status": 201, "body": {"id": "new-2"}},
        ]
    }))
    responses = await service.move_batch(
        mailbox="bot@firma.de",
        message_ids=["msg-1", "msg-2"],
        destination_folder="archive",
    )
    assert len(responses) == 2
    body = mock_session.post.call_args.kwargs["json"]
    assert len(body["requests"]) == 2
    assert body["requests"][0]["url"] == "/users/bot@firma.de/messages/msg-1/move"
    assert body["requests"][0]["body"]["destinationId"] == "archive"


async def test_move_batch_empty_list_short_circuits(service, mock_session):
    mock_session.post = MagicMock()
    responses = await service.move_batch(mailbox="bot@firma.de", message_ids=[], destination_folder="archive")
    assert responses == []
    mock_session.post.assert_not_called()


async def test_move_batch_rejects_more_than_20(service, mock_session):
    with pytest.raises(ValueError, match="max 20"):
        await service.move_batch(
            mailbox="bot@firma.de",
            message_ids=[f"id-{i}" for i in range(21)],
            destination_folder="archive",
        )


async def test_move_batch_sorts_responses_by_id(service, mock_session):
    """Responses kommen evtl. in falscher Reihenfolge — Modul muss sortieren."""
    mock_session.post = MagicMock(return_value=_make_mock_response(200, {
        "responses": [
            {"id": "2", "status": 201},
            {"id": "0", "status": 201},
            {"id": "1", "status": 201},
        ]
    }))
    responses = await service.move_batch(
        mailbox="bot@firma.de",
        message_ids=["a", "b", "c"],
        destination_folder="archive",
    )
    assert [r["id"] for r in responses] == ["0", "1", "2"]
