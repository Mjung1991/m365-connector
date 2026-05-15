# m365_connector

Wiederverwendbares Microsoft 365 Authentifizierungs-Modul für alle Projekte.

**App-only / Daemon-Modus** — kein User muss angemeldet sein. Läuft vollständig im Hintergrund.

---

## Einmalige Developer-Einrichtung (nur einmal für alle Projekte)

### 1. Azure App Registration anlegen

Im [Azure Portal](https://portal.azure.com) → **App Registrations** → **New Registration**:

- Name: `MJung M365 Connector`
- Supported account types: **Accounts in any organizational directory (Any Azure AD tenant - Multitenant)**
- Redirect URI: `Web` → `http://localhost:8888/callback`
- Klick auf **Register**

### 2. App Permissions konfigurieren

**API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**:

| Permission | Zweck |
|---|---|
| `Mail.Send` | E-Mails senden |
| `Mail.Read.All` | E-Mails lesen |
| `Mail.ReadWrite.All` | E-Mails lesen + schreiben |
| `Calendars.ReadWrite` | Kalender |
| `Files.Read.All` | OneDrive lesen |
| `User.Read.All` | Benutzer auflisten |

→ **Grant admin consent for [dein Tenant]** klicken (einmalig für deinen eigenen Tenant)

### 3. Client Secret generieren

**Certificates & secrets** → **New client secret** → 24 Monate → **Add**

Secret-Wert sofort kopieren (nur einmal sichtbar).

### 4. Developer-Credentials speichern

```bash
mkdir -p ~/.m365_connector
cat > ~/.m365_connector/app.env << EOF
M365_CLIENT_ID=<Application (client) ID aus dem Azure Portal>
M365_CLIENT_SECRET=<Secret-Wert aus Schritt 3>
EOF
chmod 600 ~/.m365_connector/app.env
```

---

## Verwendung in einem neuen Projekt

### Installation

```bash
pip install -e /pfad/zu/Module/m365_connector
```

Oder in `requirements.txt`:
```
-e /mnt/d/MJung Unternehmen/claude-setup/Module/m365_connector
```

### Kundenprojekt einrichten (Setup-Wizard)

Im Projektordner ausführen:

```bash
m365-setup
# oder: python -m m365_connector.cli.setup_wizard
```

Der Wizard:
1. Öffnet den Browser mit der Microsoft-Login-Seite
2. Kundentenant-Admin meldet sich an und erteilt Consent
3. Generiert `.env` im aktuellen Ordner

**Voraussetzung beim Kunden:** Der Admin muss ein **Global Administrator** im M365-Tenant sein.

### Code-Verwendung

```python
import asyncio
from m365_connector import M365Client

async def main():
    async with M365Client.from_env() as client:

        # Verbindung testen
        info = await client.verify_connection()
        print(f"Verbunden mit: {info['display_name']}")

        # Mail senden
        await client.mail.send(
            mailbox="bot@kundenfirma.de",
            to="empfaenger@kundenfirma.de",
            subject="Betreff",
            body="<p>Hallo!</p>",
        )

        # Inbox lesen
        messages = await client.mail.list_inbox(
            mailbox="postfach@kundenfirma.de",
            limit=10,
            unread_only=True,
        )
        for msg in messages:
            print(f"  [{msg['receivedDateTime']}] {msg['subject']}")

        # Kalender-Events
        from datetime import datetime
        events = await client.calendar.list_events(
            user="user@kundenfirma.de",
            start=datetime(2026, 6, 1),
            end=datetime(2026, 6, 30),
        )

asyncio.run(main())
```

### Ohne Context Manager (für langlebige Services)

```python
client = M365Client.from_env()
# ... verwenden ...
await client.close()  # am Ende aufrufen
```

### Credentials direkt übergeben (z.B. aus Secrets Manager)

```python
client = M365Client.from_credentials(
    tenant_id=os.environ["TENANT_ID"],
    client_id=os.environ["CLIENT_ID"],
    client_secret=os.environ["CLIENT_SECRET"],
)
```

---

## .env Format

```env
M365_TENANT_ID=<Tenant-ID des Kunden>
M365_CLIENT_ID=<App Client-ID (Developer-App)>
M365_CLIENT_SECRET=<App Secret (Developer-App)>
```

---

## Tests ausführen

```bash
cd Module/m365_connector
pip install -e ".[dev]"
pytest tests/ -v
```

---

## Unterstützte Operationen

| Service | Methode | Beschreibung |
|---|---|---|
| `client.mail` | `send(mailbox, to, subject, body)` | E-Mail senden |
| `client.mail` | `list_inbox(mailbox, limit, unread_only)` | Inbox lesen (Convenience) |
| `client.mail` | `list_messages(mailbox, folder, limit, unread_only, page_token)` | Generisches Listen mit Pagination |
| `client.mail` | `get_message(mailbox, message_id)` | Einzelne Mail mit Body |
| `client.mail` | `mark_as_read(mailbox, message_id)` | Als gelesen markieren |
| `client.mail` | `move_to_folder(mailbox, message_id, folder)` | In Ordner verschieben |
| `client.mail` | `move_batch(mailbox, message_ids, destination_folder)` | Batch-Move via Graph `$batch` (max 20) |
| `client.mail` | `fetch_attachments(mailbox, message_id)` | Anhänge laden (base64 in `contentBytes`) |
| `client.mail` | `send_forward(mailbox, message_id, to, comment)` | Mail weiterleiten |
| `client.mail.folders` | `list(mailbox, parent_id)` | Mail-Folder auflisten |
| `client.mail.folders` | `create(mailbox, name, parent_id)` | Folder anlegen |
| `client.mail.folders` | `ensure(mailbox, name, parent_id)` | Idempotent: gibt existierende ID oder erstellt neu |
| `client.calendar` | `list_events(user, start, end)` | Kalendereinträge lesen |
| `client.calendar` | `create_event(user, subject, start, end)` | Eintrag erstellen |
| `client.calendar` | `delete_event(user, event_id)` | Eintrag löschen |
| `client` | `verify_connection()` | Verbindungstest |
