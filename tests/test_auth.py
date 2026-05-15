"""Tests für M365Auth — Token-Handling."""

import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from m365_connector.auth import M365Auth


@pytest.fixture
def auth():
    return M365Auth(
        tenant_id="test-tenant-id",
        client_id="test-client-id",
        client_secret="test-secret",
    )


async def test_get_token_returns_token(auth):
    mock_token = MagicMock()
    mock_token.token = "fake-access-token"

    with patch.object(auth._credential, "get_token", new_callable=AsyncMock, return_value=mock_token):
        token = await auth.get_token()

    assert token == "fake-access-token"


async def test_get_token_called_with_correct_scope(auth):
    mock_token = MagicMock()
    mock_token.token = "token"

    with patch.object(auth._credential, "get_token", new_callable=AsyncMock, return_value=mock_token) as mock_get:
        await auth.get_token()

    mock_get.assert_called_once_with("https://graph.microsoft.com/.default")


async def test_context_manager_closes_credential(auth):
    with patch.object(auth._credential, "close", new_callable=AsyncMock) as mock_close:
        async with auth:
            pass
    mock_close.assert_called_once()


async def test_close_calls_credential_close(auth):
    with patch.object(auth._credential, "close", new_callable=AsyncMock) as mock_close:
        await auth.close()
    mock_close.assert_called_once()
