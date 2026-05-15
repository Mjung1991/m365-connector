"""Tests für oauth-Helpers und Storage-Adapter."""

import pytest
from pathlib import Path
from unittest.mock import MagicMock

from m365_connector.oauth import build_admin_consent_url, parse_consent_callback
from m365_connector.storage import CallbackStorage, EnvFileStorage, CredentialStorage


# ===== build_admin_consent_url =====


def test_build_consent_url_basic():
    url = build_admin_consent_url(
        client_id="abc-123",
        redirect_uri="http://localhost:8888/callback",
    )
    assert url.startswith("https://login.microsoftonline.com/organizations/v2.0/adminconsent?")
    assert "client_id=abc-123" in url
    assert "redirect_uri=http%3A%2F%2Flocalhost%3A8888%2Fcallback" in url
    assert "scope=https%3A%2F%2Fgraph.microsoft.com%2F.default" in url
    assert "state=" not in url


def test_build_consent_url_with_state():
    url = build_admin_consent_url(
        client_id="abc",
        redirect_uri="https://kw.example.com/m365-callback",
        state="tenant-42",
    )
    assert "state=tenant-42" in url


def test_build_consent_url_custom_scope():
    url = build_admin_consent_url(
        client_id="abc",
        redirect_uri="https://e.com/cb",
        scope="custom://scope",
    )
    assert "scope=custom%3A%2F%2Fscope" in url


# ===== parse_consent_callback =====


def test_parse_callback_success():
    result = parse_consent_callback({
        "admin_consent": "True",
        "tenant": "abc-tenant-id",
        "state": "kw-tenant-42",
    })
    assert result["success"] is True
    assert result["tenant_id"] == "abc-tenant-id"
    assert result["state"] == "kw-tenant-42"
    assert result["error"] is None


def test_parse_callback_error():
    result = parse_consent_callback({
        "error": "access_denied",
        "error_description": "User declined consent",
    })
    assert result["success"] is False
    assert result["tenant_id"] is None
    assert result["error"] == "access_denied"
    assert result["error_description"] == "User declined consent"


def test_parse_callback_admin_consent_false():
    """Wenn admin_consent != 'True' (z.B. false), gilt es als nicht erfolgreich."""
    result = parse_consent_callback({"admin_consent": "False", "tenant": "abc"})
    assert result["success"] is False
    assert result["tenant_id"] is None


def test_parse_callback_empty():
    result = parse_consent_callback({})
    assert result["success"] is False
    assert result["tenant_id"] is None


# ===== CallbackStorage =====


def test_callback_storage_passes_credentials():
    received = {}

    def capture(tenant_id, client_id, client_secret):
        received["tenant_id"] = tenant_id
        received["client_id"] = client_id
        received["client_secret"] = client_secret

    storage = CallbackStorage(capture)
    storage.save(tenant_id="t1", client_id="c1", client_secret="s1")

    assert received == {"tenant_id": "t1", "client_id": "c1", "client_secret": "s1"}


def test_callback_storage_satisfies_protocol():
    """CallbackStorage muss vom Protocol CredentialStorage erkannt werden."""
    storage = CallbackStorage(lambda t, c, s: None)
    assert isinstance(storage, CredentialStorage)


def test_callback_storage_propagates_exceptions():
    def failing(tenant_id, client_id, client_secret):
        raise ValueError("DB unreachable")

    storage = CallbackStorage(failing)
    with pytest.raises(ValueError, match="DB unreachable"):
        storage.save(tenant_id="t", client_id="c", client_secret="s")


# ===== EnvFileStorage =====


def test_env_file_storage_writes_file(tmp_path: Path):
    env_path = tmp_path / "test.env"
    storage = EnvFileStorage(env_path)
    storage.save(
        tenant_id="my-tenant",
        client_id="my-client",
        client_secret="my-secret",
    )

    content = env_path.read_text()
    # save_credentials wraps values in quotes — match key prefix and value substring
    assert "M365_TENANT_ID=" in content and "my-tenant" in content
    assert "M365_CLIENT_ID=" in content and "my-client" in content
    assert "M365_CLIENT_SECRET=" in content and "my-secret" in content


def test_env_file_storage_default_path():
    storage = EnvFileStorage()
    assert storage.path == Path(".env")


def test_env_file_storage_accepts_string_path():
    storage = EnvFileStorage("/tmp/custom.env")
    assert storage.path == Path("/tmp/custom.env")


def test_env_file_storage_satisfies_protocol():
    storage = EnvFileStorage()
    assert isinstance(storage, CredentialStorage)
