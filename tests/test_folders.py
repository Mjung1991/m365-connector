"""Tests für MailFolderService — Folder-Operationen."""

import pytest
from unittest.mock import AsyncMock, MagicMock

import aiohttp

from m365_connector.auth import M365Auth
from m365_connector.folders import MailFolderService
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
    return MailFolderService(auth, lambda: mock_session)


def _resp(status: int, json_data: dict | None = None, text: str = ""):
    r = AsyncMock()
    r.status = status
    r.json = AsyncMock(return_value=json_data or {})
    r.text = AsyncMock(return_value=text)
    r.__aenter__ = AsyncMock(return_value=r)
    r.__aexit__ = AsyncMock(return_value=False)
    return r


# ===== list =====


async def test_list_top_level_folders(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [
            {"id": "f1", "displayName": "Inbox"},
            {"id": "f2", "displayName": "Archive"},
        ]
    }))
    folders = await service.list(mailbox="bot@firma.de")
    assert len(folders) == 2
    url = mock_session.get.call_args.args[0]
    assert url == "https://graph.microsoft.com/v1.0/users/bot@firma.de/mailFolders"


async def test_list_child_folders(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {"value": []}))
    await service.list(mailbox="bot@firma.de", parent_id="parent-1")
    url = mock_session.get.call_args.args[0]
    assert url.endswith("/mailFolders/parent-1/childFolders")


async def test_list_raises_on_error(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(403, text="Forbidden"))
    with pytest.raises(M365ServiceError, match=r"folders\.list"):
        await service.list(mailbox="bot@firma.de")


# ===== create =====


async def test_create_top_level_folder(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(201, {
        "id": "new-id", "displayName": "Custom", "parentFolderId": "root",
    }))
    folder = await service.create(mailbox="bot@firma.de", name="Custom")
    assert folder["id"] == "new-id"
    body = mock_session.post.call_args.kwargs["json"]
    assert body == {"displayName": "Custom"}


async def test_create_child_folder(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(201, {"id": "child-id"}))
    await service.create(mailbox="bot@firma.de", name="Sub", parent_id="parent-1")
    url = mock_session.post.call_args.args[0]
    assert url.endswith("/mailFolders/parent-1/childFolders")


async def test_create_raises_on_error(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(409, text="Conflict"))
    with pytest.raises(M365ServiceError, match=r"folders\.create"):
        await service.create(mailbox="bot@firma.de", name="Dupe")


# ===== ensure =====


async def test_ensure_returns_existing_id_when_found(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [
            {"id": "existing-id", "displayName": "MyFolder"},
            {"id": "other-id", "displayName": "Other"},
        ]
    }))
    mock_session.post = MagicMock()  # darf nicht aufgerufen werden
    folder_id = await service.ensure(mailbox="bot@firma.de", name="MyFolder")
    assert folder_id == "existing-id"
    mock_session.post.assert_not_called()


async def test_ensure_creates_when_not_found(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [{"id": "other-id", "displayName": "Other"}]
    }))
    mock_session.post = MagicMock(return_value=_resp(201, {
        "id": "fresh-id", "displayName": "NewFolder",
    }))
    folder_id = await service.ensure(mailbox="bot@firma.de", name="NewFolder")
    assert folder_id == "fresh-id"
    mock_session.post.assert_called_once()


async def test_ensure_case_sensitive_match(service, mock_session):
    """displayName-Vergleich ist case-sensitive — 'inbox' != 'Inbox' => create."""
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [{"id": "inbox-id", "displayName": "Inbox"}]
    }))
    mock_session.post = MagicMock(return_value=_resp(201, {"id": "new-id", "displayName": "inbox"}))
    folder_id = await service.ensure(mailbox="bot@firma.de", name="inbox")
    assert folder_id == "new-id"


async def test_ensure_respects_parent_id(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {"value": []}))
    mock_session.post = MagicMock(return_value=_resp(201, {"id": "x"}))
    await service.ensure(mailbox="bot@firma.de", name="Nested", parent_id="parent-1")

    list_url = mock_session.get.call_args.args[0]
    create_url = mock_session.post.call_args.args[0]
    assert list_url.endswith("/mailFolders/parent-1/childFolders")
    assert create_url.endswith("/mailFolders/parent-1/childFolders")
