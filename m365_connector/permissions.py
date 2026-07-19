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
    # --- Schreib-Rechte fuer die Datei-Ablage (ab v0.8.0, siehe files.py) ---
    # BEVORZUGT: sites_selected. Es gibt der Anwendung Zugriff auf GENAU die Sites, die ein
    # Administrator einzeln freigibt — nicht auf den ganzen Tenant. Achtung, dreistufig:
    #   1. Zustimmung fuer `Sites.Selected` in Entra ID,
    #   2. je Site ein `POST /sites/{siteId}/permissions` mit Rolle `write`,
    #   3. Token holen.
    # Fehlt Schritt 2, hat die Anwendung trotz erteilter Zustimmung KEINEN Zugriff — das ist
    # der haeufigste Stolperstein (Beleg: concepts/permissions-selected-overview.md).
    "sites_selected": Permission(
        key="sites_selected",
        scope="Sites.Selected",
        description="Nur ausdruecklich freigegebene SharePoint-Sites (Rechte je Site separat vergeben)",
    ),
    # NOTNAGEL: gibt Schreibzugriff auf ALLE Dateien im gesamten Tenant. Nur nutzen, wenn
    # Sites.Selected nicht in Frage kommt — und dann bewusst dokumentieren.
    "files_readwrite": Permission(
        key="files_readwrite",
        scope="Files.ReadWrite.All",
        description="Dateien im ganzen Tenant lesen und schreiben (weitreichend — Sites.Selected bevorzugen)",
    ),
    "users_read": Permission(
        key="users_read",
        scope="User.Read.All",
        description="Benutzer im Tenant auflisten",
    ),
}

# Standard-Permissions für neue Projekte
DEFAULT_PERMISSIONS = {"mail_send", "mail_read", "users_read"}
