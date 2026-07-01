"""Tests für MailDeltaService — Delta-Polling."""

import pytest
from unittest.mock import AsyncMock, MagicMock

import aiohttp

from m365_connector.auth import M365Auth
from m365_connector.delta import MailDeltaService
from m365_connector.exceptions import M365ServiceError


@pytest.fixture
def auth():
    a = MagicMock(spec=M365Auth)
    a.get_token = AsyncMock(return_value="fake-token")
    return a


@pytest.fixture
def mock_session():
    return MagicMock(spec=aiohttp.ClientSession)


@pytest.fixture
def service(auth, mock_session):
    return MailDeltaService(auth, lambda: mock_session)


def _resp(status: int, json_data: dict | None = None, text: str = ""):
    r = AsyncMock()
    r.status = status
    r.json = AsyncMock(return_value=json_data or {})
    r.text = AsyncMock(return_value=text)
    r.__aenter__ = AsyncMock(return_value=r)
    r.__aexit__ = AsyncMock(return_value=False)
    return r


# ===== initial =====


async def test_initial_returns_messages_and_next_link(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [{"id": "1"}, {"id": "2"}],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/...?$skiptoken=abc",
    }))
    messages, link, done = await service.initial(mailbox="bot@firma.de", folder="inbox")
    assert len(messages) == 2
    assert link.endswith("$skiptoken=abc")
    assert done is False


async def test_initial_returns_delta_link_when_complete(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [{"id": "1"}],
        "@odata.deltaLink": "https://graph.microsoft.com/v1.0/...?$deltatoken=xyz",
    }))
    messages, link, done = await service.initial(mailbox="bot@firma.de")
    assert len(messages) == 1
    assert "$deltatoken=xyz" in link
    assert done is True


async def test_initial_uses_default_inbox_folder(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [],
        "@odata.deltaLink": "https://x",
    }))
    await service.initial(mailbox="bot@firma.de")
    url = mock_session.get.call_args.args[0]
    assert "/mailFolders/inbox/messages/delta" in url


async def test_initial_respects_custom_folder(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [], "@odata.deltaLink": "https://x",
    }))
    await service.initial(mailbox="bot@firma.de", folder="archive")
    url = mock_session.get.call_args.args[0]
    assert "/mailFolders/archive/messages/delta" in url


async def test_initial_with_latest_skips_pagination(service, mock_session):
    """latest=True fügt $deltaToken=latest an — Graph liefert sofort finalen Link."""
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [],
        "@odata.deltaLink": "https://graph.microsoft.com/...?$deltatoken=fresh",
    }))
    messages, link, done = await service.initial(
        mailbox="bot@firma.de", folder="inbox", latest=True
    )
    assert messages == []
    assert done is True
    url = mock_session.get.call_args.args[0]
    assert url.endswith("/mailFolders/inbox/messages/delta?$deltaToken=latest")


async def test_initial_without_latest_no_query_string(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [], "@odata.deltaLink": "https://x",
    }))
    await service.initial(mailbox="bot@firma.de", folder="inbox")
    url = mock_session.get.call_args.args[0]
    assert "?" not in url


async def test_initial_raises_on_error(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(403, text="Forbidden"))
    with pytest.raises(M365ServiceError, match=r"mail\.delta"):
        await service.initial(mailbox="bot@firma.de")


# ===== next =====


async def test_next_uses_link_directly(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [{"id": "3"}],
        "@odata.deltaLink": "https://final",
    }))
    messages, link, done = await service.next("https://graph.microsoft.com/v1.0/...?$skiptoken=abc")
    assert done is True
    assert link == "https://final"
    url = mock_session.get.call_args.args[0]
    assert url == "https://graph.microsoft.com/v1.0/...?$skiptoken=abc"


async def test_next_can_resume_from_delta_link(service, mock_session):
    """Bei nächstem Sync-Lauf: next(stored_delta_link) liefert nur Änderungen seitdem."""
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [{"id": "new-since-last-sync"}],
        "@odata.deltaLink": "https://updated-delta-link",
    }))
    messages, link, done = await service.next("https://stored-delta-from-last-run")
    assert messages[0]["id"] == "new-since-last-sync"
    assert link == "https://updated-delta-link"
    assert done is True


# ===== Error: weder nextLink noch deltaLink =====


async def test_response_without_either_link_raises(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {"value": []}))
    with pytest.raises(RuntimeError, match="missing both"):
        await service.initial(mailbox="bot@firma.de")
