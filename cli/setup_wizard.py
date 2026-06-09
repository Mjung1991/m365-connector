"""
M365 Setup-Wizard — einmalig pro Kundenprojekt ausführen.

Ablauf:
  1. App aus ~/.m365_connector/ auswählen (oder --app <name> übergeben)
  2. Lokaler HTTP-Server startet (Port 8888) für den OAuth-Callback
  3. Browser öffnet Admin-Consent-URL
  4. Kundentenant-Admin authentifiziert sich und erteilt Consent
  5. Wizard erfasst TENANT_ID aus dem Callback
  6. Credentials werden via Storage gespeichert (Default: .env in cwd)
  7. Verbindungstest mit Microsoft Graph

CLI-Verwendung:
  m365-setup                    # App interaktiv auswählen, .env in cwd
  m365-setup --app mail-read    # App direkt angeben
  m365-setup --add-app          # Neue App-Credentials hinterlegen

Programmatische Verwendung (z.B. aus einem Web-Backend):
  from m365_connector.cli.setup_wizard import run_wizard
  from m365_connector import CallbackStorage

  def write_to_db(tenant_id, client_id, client_secret):
      db.tenants.update(...).set(...)

  run_wizard(app_name="mail-read", storage=CallbackStorage(write_to_db))
"""

import argparse
import asyncio
import sys
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from threading import Event, Thread

# Repo-Wurzel (enthält die Pakete m365_connector + cli) in den Suchpfad, falls der
# Wizard als Script aus einem nicht-installierten Checkout läuft. Bewusst NICHT der
# darüberliegende module/-Ordner — der würde das installierte preflight-Paket durch
# das gleichnamige Projektverzeichnis module/preflight verdecken.
sys.path.insert(0, str(Path(__file__).parent.parent))

from m365_connector.credentials import (
    _DEV_DIR,
    list_developer_apps,
    load_developer_credentials,
    save_developer_credentials,
)
from m365_connector.oauth import build_admin_consent_url, parse_consent_callback
from m365_connector.storage import CredentialStorage, EnvFileStorage

_CALLBACK_PORT = 8888
_CALLBACK_PATH = "/callback"
_REDIRECT_URI = f"http://localhost:{_CALLBACK_PORT}{_CALLBACK_PATH}"


def _make_callback_handler(result_holder: list, done_event: Event):
    """Erzeugt einen Handler der Ergebnisse in die übergebenen Container schreibt."""

    class _Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            if not self.path.startswith(_CALLBACK_PATH):
                self.send_response(404)
                self.end_headers()
                return

            params = dict(urllib.parse.parse_qsl(urllib.parse.urlparse(self.path).query))
            result_holder.append(params)

            if params.get("admin_consent") == "True":
                body = "<html><body><h2>✅ Consent erteilt!</h2><p>Du kannst dieses Fenster schließen.</p></body></html>".encode("utf-8")
            else:
                error = params.get("error", "Unbekannt")
                body = f"<html><body><h2>❌ Fehler: {error}</h2><p>Du kannst dieses Fenster schließen.</p></body></html>".encode("utf-8")

            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(body)
            done_event.set()

        def log_message(self, *_):
            pass

    return _Handler


def _run_callback_server(server: HTTPServer, timeout: int = 180) -> None:
    server.timeout = timeout
    server.handle_request()


def _ensure_callback_port_free() -> None:
    """Preflight: bricht mit klarer Meldung ab, wenn der Callback-Port belegt ist.

    Verhindert den kryptischen OSError beim HTTPServer-Bind, falls Port 8888 schon
    von einem anderen Prozess (oder einem hängenden früheren Lauf) belegt ist.
    """
    from preflight import Status, check_port_free

    result = check_port_free(_CALLBACK_PORT, "OAuth-Callback-Server")
    if result.status is Status.FAIL:
        print(f"\n❌ {result.detail}")
        print(f"   Der Setup-Wizard braucht Port {_CALLBACK_PORT} für den Microsoft-Login.")
        sys.exit(1)


def _select_app(app_arg: str | None) -> str:
    """Wählt eine App aus — direkt per Argument oder interaktiv."""
    apps = list_developer_apps()

    if app_arg:
        if app_arg not in apps:
            print(f"\n❌ App '{app_arg}' nicht gefunden in {_DEV_DIR}")
            if apps:
                print(f"   Verfügbar: {', '.join(apps)}")
            sys.exit(1)
        return app_arg

    if not apps:
        print(f"\n❌ Keine Apps in {_DEV_DIR} gefunden.")
        print("   Führe 'm365-setup --add-app' aus um eine App hinzuzufügen.")
        sys.exit(1)

    if len(apps) == 1:
        print(f"\n   App: {apps[0]}")
        return apps[0]

    print("\nVerfügbare Apps:\n")
    for i, name in enumerate(apps, 1):
        creds = load_developer_credentials(name)
        print(f"  {i}. {name}  (CLIENT_ID: {creds['client_id'][:8]}...)")
    print()
    raw = input("App auswählen (Nummer oder Name): ").strip()
    if raw.isdigit():
        idx = int(raw) - 1
        if 0 <= idx < len(apps):
            return apps[idx]
    if raw in apps:
        return raw
    print(f"❌ Ungültige Auswahl: {raw}")
    sys.exit(1)


def _add_app_flow() -> None:
    """Interaktiver Flow zum Hinterlegen neuer App-Credentials."""
    print("\nNeue App-Credentials hinterlegen\n")
    app_name = input("App-Name (z.B. 'mail-read', 'mail-send'): ").strip()
    if not app_name:
        print("❌ App-Name darf nicht leer sein.")
        sys.exit(1)
    client_id = input("Application (Client) ID: ").strip()
    client_secret = input("Client Secret Value: ").strip()
    if not client_id or not client_secret:
        print("❌ CLIENT_ID und CLIENT_SECRET dürfen nicht leer sein.")
        sys.exit(1)
    save_developer_credentials(app_name, client_id, client_secret)
    print(f"\n✅ App '{app_name}' gespeichert in {_DEV_DIR / f'{app_name}.env'}")


def run_wizard(
    app_name: str | None = None,
    storage: CredentialStorage | None = None,
    test_connection: bool = True,
) -> dict:
    """Treibt den Setup-Flow programmatisch.

    Args:
        app_name: Name der Developer-App aus ~/.m365_connector/. None = interaktive Auswahl.
        storage: Persistenz für die Credentials. None = EnvFileStorage(".env") in cwd.
        test_connection: Wenn True, validiert die gespeicherten Credentials gegen Graph.

    Returns:
        Dict mit tenant_id, client_id, client_secret.

    Raises:
        SystemExit: bei Abbruch oder Fehler im Consent-Flow.
    """
    if storage is None:
        storage = EnvFileStorage(".env")

    # App auswählen
    selected_app = _select_app(app_name)
    dev_creds = load_developer_credentials(selected_app)
    print(f"   ✅ {selected_app}  (CLIENT_ID: {dev_creds['client_id'][:8]}...)")

    # Consent-URL bauen
    consent_url = build_admin_consent_url(
        client_id=dev_creds["client_id"],
        redirect_uri=_REDIRECT_URI,
    )
    print(f"\n2. Admin-Consent")
    print(f"\n   Öffne den folgenden Link im Browser und melde dich")
    print(f"   mit einem M365 Global Administrator des Kunden an:\n")
    print(f"   {consent_url}\n")

    # Callback-Server starten — vorher prüfen, ob der Port wirklich frei ist
    # (klare Meldung statt rohem OSError beim Bind).
    _ensure_callback_port_free()
    result_holder: list[dict] = []
    done_event = Event()
    handler_class = _make_callback_handler(result_holder, done_event)
    server = HTTPServer(("localhost", _CALLBACK_PORT), handler_class)
    Thread(target=_run_callback_server, args=(server, 180), daemon=True).start()

    opened = webbrowser.open(consent_url)
    if not opened:
        print("   (Browser konnte nicht automatisch geöffnet werden — URL oben manuell aufrufen)")

    print("   Warte auf Bestätigung im Browser... (Timeout: 3 Minuten)")
    received = done_event.wait(timeout=180)
    if not received:
        print("\n❌ Timeout — kein Callback empfangen. Bitte erneut ausführen.")
        sys.exit(1)

    raw_params = result_holder[0] if result_holder else {}
    parsed = parse_consent_callback(raw_params)
    if not parsed["success"]:
        print(f"\n❌ Consent abgelehnt: {parsed['error']}\n   {parsed['error_description'] or ''}")
        sys.exit(1)

    tenant_id = parsed["tenant_id"]
    if not tenant_id:
        print("\n❌ TENANT_ID konnte nicht aus dem Callback gelesen werden.")
        print("   Parameter empfangen:", list(raw_params.keys()))
        sys.exit(1)

    print(f"\n   ✅ Consent erteilt! Tenant-ID: {tenant_id}")

    # Storage speichern
    print(f"\n3. Credentials speichern ({type(storage).__name__})")
    storage.save(
        tenant_id=tenant_id,
        client_id=dev_creds["client_id"],
        client_secret=dev_creds["client_secret"],
    )
    print(f"   ✅ Gespeichert")

    credentials = {
        "tenant_id": tenant_id,
        "client_id": dev_creds["client_id"],
        "client_secret": dev_creds["client_secret"],
    }

    if test_connection:
        print("\n4. Verbindungstest...")
        asyncio.run(_test_connection_with_credentials(credentials))

    return credentials


async def _test_connection_with_credentials(credentials: dict) -> None:
    from m365_connector import M365Client
    try:
        async with M365Client.from_credentials(**credentials) as client:
            info = await client.verify_connection()
            print(f"   ✅ Verbindung erfolgreich!")
            print(f"   Tenant: {info['display_name']} ({info['tenant_id']})")
            domains = ", ".join(info["domains"][:3])
            print(f"   Domains: {domains}")
    except Exception as e:
        print(f"   ⚠️  Verbindungstest fehlgeschlagen: {e}")
        print("   Die Credentials wurden gespeichert — prüfe Permissions und Tenant.")


def main() -> None:
    parser = argparse.ArgumentParser(description="M365 Connector Setup-Wizard")
    parser.add_argument("--app", help="App-Name aus ~/.m365_connector/ (z.B. mail-read)")
    parser.add_argument("--add-app", action="store_true", help="Neue App-Credentials hinterlegen")
    parser.add_argument(
        "--env-file",
        default=".env",
        help="Pfad zur Ziel-.env-Datei (Default: .env in cwd)",
    )
    args = parser.parse_args()

    print("=" * 60)
    print("  M365 Connector — Setup-Wizard")
    print("=" * 60)

    if args.add_app:
        _add_app_flow()
        return

    print("\n1. App auswählen")
    storage = EnvFileStorage(args.env_file)
    run_wizard(app_name=args.app, storage=storage)

    print("\n" + "=" * 60)
    print("  Setup abgeschlossen! Verwende in deinem Projekt:")
    print()
    print("    from m365_connector import M365Client")
    print("    async with M365Client.from_env() as client:")
    print("        messages = await client.mail.list_inbox(...)")
    print("=" * 60)


if __name__ == "__main__":
    main()
