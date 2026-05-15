"""
M365 Setup-Wizard — einmalig pro Kundenprojekt ausführen.

Ablauf:
  1. App aus ~/.m365_connector/ auswählen (oder --app <name> übergeben)
  2. Lokaler HTTP-Server startet (Port 8888) für den OAuth-Callback
  3. Browser öffnet Admin-Consent-URL
  4. Kundentenant-Admin authentifiziert sich und erteilt Consent
  5. Wizard erfasst TENANT_ID aus dem Callback
  6. .env im aktuellen Ordner wird generiert
  7. Verbindungstest mit Microsoft Graph

Verwendung:
  m365-setup                    # App interaktiv auswählen
  m365-setup --app mail-read    # App direkt angeben
  m365-setup --add-app          # Neue App-Credentials hinterlegen
"""

import argparse
import asyncio
import sys
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from threading import Event, Thread

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from m365_connector.credentials import (
    _DEV_DIR,
    list_developer_apps,
    load_developer_credentials,
    save_credentials,
    save_developer_credentials,
)

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


def _build_consent_url(client_id: str) -> str:
    params = urllib.parse.urlencode({
        "client_id": client_id,
        "redirect_uri": _REDIRECT_URI,
        "scope": "https://graph.microsoft.com/.default",
    })
    return f"https://login.microsoftonline.com/organizations/v2.0/adminconsent?{params}"


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


def main() -> None:
    parser = argparse.ArgumentParser(description="M365 Connector Setup-Wizard")
    parser.add_argument("--app", help="App-Name aus ~/.m365_connector/ (z.B. mail-read)")
    parser.add_argument("--add-app", action="store_true", help="Neue App-Credentials hinterlegen")
    args = parser.parse_args()

    print("=" * 60)
    print("  M365 Connector — Setup-Wizard")
    print("=" * 60)

    if args.add_app:
        _add_app_flow()
        return

    # App auswählen
    print("\n1. App auswählen")
    app_name = _select_app(args.app)
    dev_creds = load_developer_credentials(app_name)
    print(f"   ✅ {app_name}  (CLIENT_ID: {dev_creds['client_id'][:8]}...)")

    # Consent-URL bauen
    consent_url = _build_consent_url(dev_creds["client_id"])
    print(f"\n2. Admin-Consent")
    print(f"\n   Öffne den folgenden Link im Browser und melde dich")
    print(f"   mit einem M365 Global Administrator des Kunden an:\n")
    print(f"   {consent_url}\n")

    # Callback-Server starten (state via Closure, kein Class-State)
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

    params = result_holder[0] if result_holder else {}
    if params.get("admin_consent") != "True":
        error = params.get("error", "Unbekannt")
        desc = params.get("error_description", "")
        print(f"\n❌ Consent abgelehnt: {error}\n   {desc}")
        sys.exit(1)

    tenant_id = params.get("tenant")
    if not tenant_id:
        print("\n❌ TENANT_ID konnte nicht aus dem Callback gelesen werden.")
        print("   Parameter empfangen:", list(params.keys()))
        sys.exit(1)

    print(f"\n   ✅ Consent erteilt! Tenant-ID: {tenant_id}")

    # .env speichern
    env_path = Path(".env")
    print(f"\n3. Credentials speichern → {env_path.absolute()}")
    save_credentials(
        tenant_id=tenant_id,
        client_id=dev_creds["client_id"],
        client_secret=dev_creds["client_secret"],
        env_path=env_path,
    )
    print("   ✅ .env gespeichert (chmod 600)")

    # Verbindungstest
    print("\n4. Verbindungstest...")
    asyncio.run(_test_connection(env_path))


async def _test_connection(env_path: Path) -> None:
    from m365_connector import M365Client
    try:
        async with M365Client.from_env(env_path) as client:
            info = await client.verify_connection()
            print(f"   ✅ Verbindung erfolgreich!")
            print(f"   Tenant: {info['display_name']} ({info['tenant_id']})")
            domains = ", ".join(info["domains"][:3])
            print(f"   Domains: {domains}")
    except Exception as e:
        print(f"   ⚠️  Verbindungstest fehlgeschlagen: {e}")
        print("   Die .env wurde gespeichert — prüfe Credentials und Permissions.")

    print("\n" + "=" * 60)
    print("  Setup abgeschlossen! Verwende in deinem Projekt:")
    print()
    print("    from m365_connector import M365Client")
    print("    async with M365Client.from_env() as client:")
    print("        messages = await client.mail.list_inbox(...)")
    print("=" * 60)


if __name__ == "__main__":
    main()
