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


# ===== Zertifikats-Authentifizierung (Q4) =====


def test_secret_auth_mode(auth):
    """Der Secret-Pfad (rueckwaertskompatibel) meldet auth_mode 'secret'."""
    assert auth.auth_mode == "secret"


def test_certificate_from_path_uses_certificate_credential():
    with patch("m365_connector.auth.CertificateCredential") as mock_cert:
        a = M365Auth(
            tenant_id="t", client_id="c", certificate_path="/secrets/q.pem"
        )
    assert a.auth_mode == "certificate"
    mock_cert.assert_called_once_with(
        tenant_id="t", client_id="c", certificate_path="/secrets/q.pem"
    )


def test_certificate_from_data_with_password_and_chain():
    with patch("m365_connector.auth.CertificateCredential") as mock_cert:
        a = M365Auth(
            tenant_id="t",
            client_id="c",
            certificate_data=b"PEMBYTES",
            certificate_password="pw",
            send_certificate_chain=True,
        )
    assert a.auth_mode == "certificate"
    mock_cert.assert_called_once_with(
        tenant_id="t",
        client_id="c",
        certificate_path=None,
        certificate_data=b"PEMBYTES",
        password="pw",
        send_certificate_chain=True,
    )


def test_from_certificate_classmethod():
    with patch("m365_connector.auth.CertificateCredential") as mock_cert:
        a = M365Auth.from_certificate("t", "c", certificate_path="/x.pfx", password="pw")
    assert a.auth_mode == "certificate"
    mock_cert.assert_called_once_with(
        tenant_id="t", client_id="c", certificate_path="/x.pfx", password="pw"
    )


def test_both_secret_and_certificate_is_error():
    with pytest.raises(ValueError, match="not both"):
        M365Auth("t", "c", client_secret="s", certificate_path="/x.pem")


def test_no_credentials_is_error():
    with pytest.raises(ValueError, match="Missing credentials"):
        M365Auth("t", "c")
