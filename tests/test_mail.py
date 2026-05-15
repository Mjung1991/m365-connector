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
