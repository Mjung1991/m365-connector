"""Tests für SubscriptionService + validate_subscription_token."""

import pytest
from unittest.mock import AsyncMock, MagicMock

import aiohttp

from m365_connector.auth import M365Auth
from m365_connector.subscriptions import SubscriptionService, validate_subscription_token
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
    return SubscriptionService(auth, lambda: mock_session)


def _resp(status: int, json_data: dict | None = None, text: str = ""):
    r = AsyncMock()
    r.status = status
    r.json = AsyncMock(return_value=json_data or {})
    r.text = AsyncMock(return_value=text)
    r.__aenter__ = AsyncMock(return_value=r)
    r.__aexit__ = AsyncMock(return_value=False)
    return r


# ===== validate_subscription_token =====


def test_validate_token_returns_token_when_present():
    token = validate_subscription_token({"validationToken": "abc123"})
    assert token == "abc123"


def test_validate_token_returns_none_when_absent():
    assert validate_subscription_token({}) is None
    assert validate_subscription_token({"foo": "bar"}) is None


def test_validate_token_accepts_mapping_likes():
    from collections.abc import Mapping

    class StarletteLikeQueryParams(Mapping):
        def __init__(self, d):
            self._d = d
        def __getitem__(self, k): return self._d[k]
        def __iter__(self): return iter(self._d)
        def __len__(self): return len(self._d)
        def get(self, k, default=None): return self._d.get(k, default)

    qp = StarletteLikeQueryParams({"validationToken": "hello"})
    assert validate_subscription_token(qp) == "hello"


# ===== create =====


async def test_create_basic_subscription(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(201, {
        "id": "sub-1",
        "expirationDateTime": "2026-05-18T12:00:00Z",
    }))
    sub = await service.create(
        resource="users/abc/mailFolders('Inbox')/messages",
        notification_url="https://example.com/webhook",
        expires="2026-05-18T12:00:00Z",
        client_state="secret",
    )
    assert sub["id"] == "sub-1"
    body = mock_session.post.call_args.kwargs["json"]
    assert body["changeType"] == "created,updated"
    assert body["notificationUrl"] == "https://example.com/webhook"
    assert body["clientState"] == "secret"
    assert "lifecycleNotificationUrl" not in body


async def test_create_with_lifecycle_url(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(201, {"id": "sub-2"}))
    await service.create(
        resource="users/x/messages",
        notification_url="https://e.com/wh",
        expires="2026-05-18T00:00:00Z",
        client_state="s",
        change_type="created",
        lifecycle_notification_url="https://e.com/lifecycle",
    )
    body = mock_session.post.call_args.kwargs["json"]
    assert body["lifecycleNotificationUrl"] == "https://e.com/lifecycle"
    assert body["changeType"] == "created"


async def test_create_raises_on_error(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(400, text="ValidationError"))
    with pytest.raises(M365ServiceError, match=r"subscriptions\.create"):
        await service.create(
            resource="r", notification_url="u", expires="e", client_state="s"
        )


# ===== renew =====


async def test_renew_updates_expiration(service, mock_session):
    mock_session.patch = MagicMock(return_value=_resp(200, {
        "id": "sub-1",
        "expirationDateTime": "2026-05-20T00:00:00Z",
    }))
    sub = await service.renew("sub-1", "2026-05-20T00:00:00Z")
    assert sub["expirationDateTime"] == "2026-05-20T00:00:00Z"
    body = mock_session.patch.call_args.kwargs["json"]
    assert body == {"expirationDateTime": "2026-05-20T00:00:00Z"}


async def test_renew_raises_on_error(service, mock_session):
    mock_session.patch = MagicMock(return_value=_resp(404))
    with pytest.raises(M365ServiceError, match=r"subscriptions\.renew"):
        await service.renew("missing", "2026-05-20T00:00:00Z")


# ===== delete =====


async def test_delete_accepts_204(service, mock_session):
    mock_session.delete = MagicMock(return_value=_resp(204))
    await service.delete("sub-1")
    mock_session.delete.assert_called_once()


async def test_delete_accepts_200(service, mock_session):
    mock_session.delete = MagicMock(return_value=_resp(200))
    await service.delete("sub-1")


async def test_delete_raises_on_error(service, mock_session):
    mock_session.delete = MagicMock(return_value=_resp(403))
    with pytest.raises(M365ServiceError, match=r"subscriptions\.delete"):
        await service.delete("sub-1")


# ===== list / get =====


async def test_list_subscriptions(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {
        "value": [{"id": "a"}, {"id": "b"}]
    }))
    subs = await service.list()
    assert len(subs) == 2


async def test_get_subscription(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {"id": "sub-1"}))
    sub = await service.get("sub-1")
    assert sub["id"] == "sub-1"


async def test_get_raises_on_404(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(404))
    with pytest.raises(M365ServiceError, match=r"subscriptions\.get"):
        await service.get("missing")
