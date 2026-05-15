"""Liest und schreibt M365-Credentials in .env-Dateien."""

import os
from pathlib import Path

from dotenv import dotenv_values, load_dotenv, set_key


_REQUIRED_KEYS = ("M365_TENANT_ID", "M365_CLIENT_ID", "M365_CLIENT_SECRET")
_DEV_DIR = Path.home() / ".m365_connector"


def load_credentials(env_path: Path = Path(".env")) -> dict[str, str]:
    """Liest Tenant-Credentials aus einer .env-Datei."""
    load_dotenv(env_path, override=True)
    missing = [k for k in _REQUIRED_KEYS if not os.environ.get(k)]
    if missing:
        raise RuntimeError(
            f"Missing M365 credentials in {env_path}: {', '.join(missing)}\n"
            "Run 'm365-setup' to set up the connection."
        )
    return {
        "tenant_id": os.environ["M365_TENANT_ID"],
        "client_id": os.environ["M365_CLIENT_ID"],
        "client_secret": os.environ["M365_CLIENT_SECRET"],
    }


def list_developer_apps() -> list[str]:
    """Gibt alle verfügbaren App-Namen aus ~/.m365_connector/ zurück."""
    if not _DEV_DIR.exists():
        return []
    return sorted(p.stem for p in _DEV_DIR.glob("*.env"))


def load_developer_credentials(app_name: str) -> dict[str, str]:
    """Liest Credentials für eine benannte App aus ~/.m365_connector/<app_name>.env."""
    path = _DEV_DIR / f"{app_name}.env"
    if not path.exists():
        available = list_developer_apps()
        hint = f"\n  Available: {', '.join(available)}" if available else ""
        raise FileNotFoundError(
            f"App credentials not found: {path}{hint}\n"
            f"Run: m365-setup --add-app"
        )
    values = dotenv_values(path)
    missing = [k for k in ("M365_CLIENT_ID", "M365_CLIENT_SECRET") if not values.get(k)]
    if missing:
        raise RuntimeError(f"Missing keys in {path}: {', '.join(missing)}")
    return {
        "app_name": app_name,
        "client_id": values["M365_CLIENT_ID"],
        "client_secret": values["M365_CLIENT_SECRET"],
    }


def save_developer_credentials(app_name: str, client_id: str, client_secret: str) -> None:
    """Speichert Developer-App-Credentials in ~/.m365_connector/<app_name>.env."""
    _DEV_DIR.mkdir(parents=True, exist_ok=True)
    path = _DEV_DIR / f"{app_name}.env"
    path.touch(exist_ok=True)
    path.chmod(0o600)
    set_key(str(path), "M365_CLIENT_ID", client_id)
    set_key(str(path), "M365_CLIENT_SECRET", client_secret)


def save_credentials(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    env_path: Path = Path(".env"),
) -> None:
    """Speichert Tenant-Credentials in eine .env-Datei."""
    env_path.touch(exist_ok=True)
    env_path.chmod(0o600)
    set_key(str(env_path), "M365_TENANT_ID", tenant_id)
    set_key(str(env_path), "M365_CLIENT_ID", client_id)
    set_key(str(env_path), "M365_CLIENT_SECRET", client_secret)
