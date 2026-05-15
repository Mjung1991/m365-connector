"""
Verfügbare Microsoft Graph API Permissions für den Connector.
Alle als Application Permissions (App-only / Daemon-Modus).
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Permission:
    key: str
    scope: str
    description: str


PERMISSIONS: dict[str, Permission] = {
    "mail_send": Permission(
        key="mail_send",
        scope="Mail.Send",
        description="E-Mails senden (als beliebiges Postfach im Tenant)",
    ),
    "mail_read": Permission(
        key="mail_read",
        scope="Mail.Read.All",
        description="E-Mails lesen (alle Postfächer im Tenant)",
    ),
    "mail_readwrite": Permission(
        key="mail_readwrite",
        scope="Mail.ReadWrite.All",
        description="E-Mails lesen und schreiben (alle Postfächer)",
    ),
    "calendar_readwrite": Permission(
        key="calendar_readwrite",
        scope="Calendars.ReadWrite",
        description="Kalender lesen und schreiben",
    ),
    "files_read": Permission(
        key="files_read",
        scope="Files.Read.All",
        description="OneDrive-Dateien lesen",
    ),
    "users_read": Permission(
        key="users_read",
        scope="User.Read.All",
        description="Benutzer im Tenant auflisten",
    ),
}

# Standard-Permissions für neue Projekte
DEFAULT_PERMISSIONS = {"mail_send", "mail_read", "users_read"}
