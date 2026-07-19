"""Microsoft Graph — Datei-Ablage in SharePoint/OneDrive (Drive-Operationen).

Ergaenzt den Connector um das, was bisher fehlte: **Dateien schreiben**. Bis v0.7.3 kannte das
Modul nur `Files.Read.All` als Rechte-Bezeichner und hatte keinerlei Drive-Code.

Belegte Grundlage (offizielle Microsoft-Graph-Doku, ueber Context7 geprueft am 19.07.2026):
  * `api-reference/v1.0/api/driveitem-put-content.md`
      PUT /sites/{site-id}/drive/items/{parent-id}:/{filename}:/content   — bis 250 MB in einem Rutsch
  * `api-reference/v1.0/api/driveitem-createuploadsession.md`
      POST …:/{fileName}:/createUploadSession  → `uploadUrl`, danach PUT-Fragmente
  * `api-reference/v1.0/api/site-getbypath.md`
      GET /sites/{hostname}:/{server-relativer-pfad}  → Site-`id` (Form `hostname,guid,guid`)
  * `concepts/permissions-selected-overview.md`
      `Sites.Selected` braucht DREI Schritte (Zustimmung, Rechte je Site, Token)

DREI FALLEN, die hier bewusst behandelt sind (jede hat schon Projekte gekostet):

  1. **Der Fragment-PUT darf KEINEN `Authorization`-Header tragen.** Die `uploadUrl` ist bereits
     vorautorisiert; schickt man den Token trotzdem mit, antwortet Graph mit **401** — und zwar
     mitten im Upload, nicht am Anfang. (Doku: driveitem-createuploadsession.md, Fehlerliste.)
  2. **Fragmentgroesse MUSS ein Vielfaches von 320 KiB (327.680 Bytes) sein.** Sonst schlaegt der
     Upload erst **nach dem letzten Fragment** beim Zusammensetzen fehl — der Fehler erscheint also
     weit weg von seiner Ursache. Max. 60 MiB je Anfrage, ~10 MiB ist der empfohlene Wert.
  3. **`_http.request_with_retry` ist hier NICHT nutzbar.** Der Helfer sendet Inhalte ausschliesslich
     als `json=…` (siehe `_http.py`) — eine Datei wuerde damit JSON-kodiert und damit **still
     beschaedigt** im Ziel landen. Datei-Inhalte gehen darum ueber einen eigenen Byte-Pfad
     (`data=…` + eigener `Content-Type`).

Grundsatz dieses Moduls: **es transportiert nur.** Was abgelegt werden *darf* (Datenklasse,
erlaubte Zielordner), entscheidet die aufrufende Anwendung — bei Q ist das die Regel-Schicht
`m365_files.py`, analog zur Allowlist beim Mail-Versand.
"""

from __future__ import annotations

import logging
import re
from typing import Callable
from urllib.parse import quote

import aiohttp

from .auth import M365Auth
from ._http import to_typed as _to_typed

logger = logging.getLogger(__name__)

_GRAPH = "https://graph.microsoft.com/v1.0"

# --- Groessen (aus der offiziellen Doku, nicht geraten) ---------------------------------------
# Vielfaches von 320 KiB ist PFLICHT fuer Fragment-Uploads (driveitem-createuploadsession.md).
FRAGMENT_EINHEIT = 320 * 1024                 # 327.680 Bytes
# Empfohlen: ~10 MiB je Fragment bei stabiler Leitung. 32 x 320 KiB = 10 MiB, exakt teilbar.
STANDARD_FRAGMENT = 32 * FRAGMENT_EINHEIT     # 10.485.760 Bytes
# Ab dieser Groesse wird die Upload-Sitzung genutzt statt eines einzelnen PUT.
# Die Doku empfiehlt Sitzungen ab 10 MiB; wir setzen bewusst frueher an (4 MiB), weil der
# einzelne PUT bei langsamer Leitung sonst ins 60s-Timeout der Session laeuft.
SITZUNG_AB_BYTES = 4 * 1024 * 1024

# SharePoint verbietet diese Zeichen in Datei-/Ordnernamen. Wer sie durchlaesst, bekommt eine
# nichtssagende 400 von Graph statt einer klaren Fehlermeldung.
_VERBOTENE_ZEICHEN = re.compile(r'[<>:"|?*\\]')


class M365PfadFehler(ValueError):
    """Der gewuenschte Ziel-Pfad ist fuer SharePoint/OneDrive nicht zulaessig."""


def pfad_pruefen(pfad: str) -> str:
    """Prueft und normalisiert einen Ziel-Pfad wie ``Q-Ablage/Vorlagen/angebot.pdf``.

    Rein (kein Netz) → direkt testbar. Wirft `M365PfadFehler` statt eine kaputte Anfrage zu senden.

    Abgewiesen werden: leere Pfade, absolute Pfade, `..`-Sprünge (Ausbruch aus dem Zielordner),
    verbotene SharePoint-Zeichen, sowie Namen mit fuehrenden/abschliessenden Leerzeichen oder
    Punkten (die legt SharePoint stillschweigend um).
    """
    if not pfad or not pfad.strip():
        raise M365PfadFehler("leerer Pfad")
    p = pfad.strip().replace("\\", "/").strip("/")
    if not p:
        raise M365PfadFehler("leerer Pfad")
    if len(p) > 400:
        raise M365PfadFehler(f"Pfad zu lang ({len(p)} Zeichen, max 400)")

    teile = [t for t in p.split("/")]
    for t in teile:
        if not t:
            raise M365PfadFehler(f"leerer Pfad-Abschnitt in {pfad!r} (doppelter Schraegstrich?)")
        if t in (".", ".."):
            # Ausbruch aus dem freigegebenen Zielordner — genau das soll die Regel-Schicht verhindern.
            raise M365PfadFehler(f"unzulaessiger Pfad-Abschnitt {t!r} in {pfad!r}")
        if _VERBOTENE_ZEICHEN.search(t):
            raise M365PfadFehler(f"verbotenes Zeichen in {t!r} (nicht erlaubt: < > : \" | ? * \\)")
        if t != t.strip() or t.endswith("."):
            raise M365PfadFehler(f"Name {t!r} darf nicht mit Leerzeichen/Punkt beginnen oder enden")
    return "/".join(teile)


def _pfad_kodieren(pfad: str) -> str:
    """URL-kodiert einen Pfad, laesst die Schraegstriche als Trenner stehen."""
    return quote(pfad, safe="/")


def fragmente(gesamt: int, fragment_groesse: int = STANDARD_FRAGMENT) -> list[tuple[int, int]]:
    """Zerlegt eine Dateigroesse in (start, ende_einschliesslich)-Bereiche.

    Rein → testbar ohne Netz. Jedes Fragment ausser dem letzten ist ein Vielfaches von 320 KiB
    (siehe Falle 2 im Modul-Kopf).
    """
    if gesamt <= 0:
        raise ValueError("Dateigroesse muss > 0 sein")
    if fragment_groesse % FRAGMENT_EINHEIT != 0:
        raise ValueError(
            f"Fragmentgroesse {fragment_groesse} ist kein Vielfaches von 320 KiB "
            f"({FRAGMENT_EINHEIT}) — Graph scheitert sonst erst beim Zusammensetzen"
        )
    out: list[tuple[int, int]] = []
    start = 0
    while start < gesamt:
        ende = min(start + fragment_groesse, gesamt) - 1
        out.append((start, ende))
        start = ende + 1
    return out


class FilesService:
    """Datei-Ablage in einem SharePoint-Site-Drive bzw. OneDrive.

    Beispiel::

        async with M365Client.from_credentials(...) as client:
            site_id = await client.files.site_id("firma.sharepoint.com", "/sites/Q-Ablage")
            drive_id = await client.files.drive_id(site_id)
            await client.files.upload(drive_id, "Q-Ablage/PDF/anleitung.pdf", daten,
                                      content_type="application/pdf")
    """

    def __init__(self, auth: M365Auth, get_session: Callable[[], aiohttp.ClientSession]) -> None:
        self._auth = auth
        self._get_session = get_session

    # --- Ziel finden ---------------------------------------------------------------------------

    async def site_id(self, hostname: str, site_pfad: str) -> str:
        """Loest eine SharePoint-Site ueber ihren Web-Pfad auf → Site-`id`.

        `hostname` z.B. `firma.sharepoint.com`, `site_pfad` z.B. `/sites/Q-Ablage`.
        Beleg: site-getbypath.md.
        """
        pfad = "/" + site_pfad.strip("/")
        token = await self._auth.get_token()
        url = f"{_GRAPH}/sites/{hostname}:{_pfad_kodieren(pfad)}"
        async with self._get_session().get(
            url, headers={"Authorization": f"Bearer {token}"}, params={"$select": "id,webUrl,displayName"}
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "files.site_id", await resp.text())
            return (await resp.json())["id"]

    async def drive_id(self, site_id: str) -> str:
        """Standard-Dokumentbibliothek („Dokumente") einer Site → Drive-`id`."""
        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/sites/{site_id}/drive",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id,name,webUrl"},
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "files.drive_id", await resp.text())
            return (await resp.json())["id"]

    async def user_drive_id(self, benutzer: str) -> str:
        """OneDrive eines Benutzers (Postfach-Adresse oder Objekt-ID) → Drive-`id`."""
        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/users/{benutzer}/drive",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id,name,webUrl"},
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "files.user_drive_id", await resp.text())
            return (await resp.json())["id"]

    # --- Ordner --------------------------------------------------------------------------------

    async def ensure_folder(self, drive_id: str, ordner_pfad: str) -> dict:
        """Legt einen Ordner-Pfad an, Ebene fuer Ebene — idempotent.

        Graph legt Zwischen-Ebenen **nicht** automatisch an; ein Upload nach
        `Q-Ablage/Vorlagen/x.pdf` scheitert, wenn `Q-Ablage/Vorlagen` fehlt. Darum hier
        Schritt fuer Schritt mit `conflictBehavior=fail`, wobei ein 409 (existiert schon)
        der Normalfall und **kein** Fehler ist.
        """
        pfad = pfad_pruefen(ordner_pfad)
        token = await self._auth.get_token()
        letzte: dict = {}
        bisher = ""
        for teil in pfad.split("/"):
            eltern = bisher
            ziel = f"{_GRAPH}/drives/{drive_id}/root/children" if not eltern else \
                   f"{_GRAPH}/drives/{drive_id}/root:/{_pfad_kodieren(eltern)}:/children"
            async with self._get_session().post(
                ziel,
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json={
                    "name": teil,
                    "folder": {},
                    # 'fail' statt 'replace': ein bestehender Ordner darf NIE ueberschrieben werden
                    # (das wuerde seinen Inhalt verlieren). 409 = existiert bereits = gut.
                    "@microsoft.graph.conflictBehavior": "fail",
                },
            ) as resp:
                if resp.status in (200, 201):
                    letzte = await resp.json()
                elif resp.status == 409:
                    letzte = {"name": teil, "existierte_schon": True}
                else:
                    raise _to_typed(resp.status, f"files.ensure_folder({teil})", await resp.text())
            bisher = f"{eltern}/{teil}" if eltern else teil
        return letzte

    # --- Hochladen -----------------------------------------------------------------------------

    async def upload(
        self,
        drive_id: str,
        ziel_pfad: str,
        inhalt: bytes,
        *,
        content_type: str = "application/octet-stream",
        conflict_behavior: str = "replace",
        ordner_anlegen: bool = True,
        fragment_groesse: int = STANDARD_FRAGMENT,
    ) -> dict:
        """Legt eine Datei ab. Waehlt selbst zwischen einfachem PUT und Upload-Sitzung.

        `conflict_behavior`: `replace` (Standard), `rename` oder `fail`.
        Rueckgabe: das `driveItem` aus Graph (u.a. `id`, `name`, `size`, `webUrl`).
        """
        pfad = pfad_pruefen(ziel_pfad)
        if not isinstance(inhalt, (bytes, bytearray)):
            raise TypeError("inhalt muss bytes sein (kein str) — sonst landet Text falsch kodiert im Ziel")
        if not inhalt:
            raise ValueError("leerer Inhalt — es wird keine leere Datei abgelegt")
        if conflict_behavior not in ("replace", "rename", "fail"):
            raise ValueError(f"unbekanntes conflict_behavior {conflict_behavior!r}")

        if ordner_anlegen and "/" in pfad:
            await self.ensure_folder(drive_id, pfad.rsplit("/", 1)[0])

        if len(inhalt) < SITZUNG_AB_BYTES:
            return await self._upload_klein(drive_id, pfad, bytes(inhalt), content_type, conflict_behavior)
        return await self._upload_gross(drive_id, pfad, bytes(inhalt), conflict_behavior, fragment_groesse)

    async def _upload_klein(
        self, drive_id: str, pfad: str, inhalt: bytes, content_type: str, conflict_behavior: str
    ) -> dict:
        """Einzelner PUT (Doku: bis 250 MB moeglich; wir nutzen ihn nur fuer kleine Dateien).

        WICHTIG: `data=` (rohe Bytes), NICHT `json=` — sonst wird die Datei JSON-kodiert und ist
        im Ziel beschaedigt (Falle 3 im Modul-Kopf).
        """
        token = await self._auth.get_token()
        url = f"{_GRAPH}/drives/{drive_id}/root:/{_pfad_kodieren(pfad)}:/content"
        async with self._get_session().put(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": content_type},
            params={"@microsoft.graph.conflictBehavior": conflict_behavior},
            data=inhalt,
        ) as resp:
            if resp.status not in (200, 201):
                raise _to_typed(resp.status, "files.upload(klein)", await resp.text())
            return await resp.json()

    async def _upload_gross(
        self, drive_id: str, pfad: str, inhalt: bytes, conflict_behavior: str, fragment_groesse: int
    ) -> dict:
        """Upload-Sitzung: einmal anlegen, dann Fragmente der Reihe nach schicken."""
        token = await self._auth.get_token()
        name = pfad.rsplit("/", 1)[-1]
        url = f"{_GRAPH}/drives/{drive_id}/root:/{_pfad_kodieren(pfad)}:/createUploadSession"
        async with self._get_session().post(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"item": {"@microsoft.graph.conflictBehavior": conflict_behavior, "name": name}},
        ) as resp:
            if resp.status not in (200, 201):
                raise _to_typed(resp.status, "files.createUploadSession", await resp.text())
            upload_url = (await resp.json())["uploadUrl"]

        gesamt = len(inhalt)
        letzte_antwort: dict = {}
        try:
            for start, ende in fragmente(gesamt, fragment_groesse):
                stueck = inhalt[start:ende + 1]
                # KEIN Authorization-Header! Die uploadUrl ist vorautorisiert; ein mitgesendeter
                # Token fuehrt laut Doku zu 401 (Falle 1 im Modul-Kopf).
                async with self._get_session().put(
                    upload_url,
                    headers={
                        "Content-Length": str(len(stueck)),
                        "Content-Range": f"bytes {start}-{ende}/{gesamt}",
                    },
                    data=stueck,
                ) as resp:
                    if resp.status in (200, 201):
                        letzte_antwort = await resp.json()   # fertig — das ist das driveItem
                    elif resp.status == 202:
                        letzte_antwort = {}                  # weiter mit dem naechsten Fragment
                    else:
                        raise _to_typed(
                            resp.status, f"files.upload(fragment {start}-{ende})", await resp.text()
                        )
        except Exception:
            # Abgebrochene Sitzung aufraeumen, damit im Ziel keine halbe Datei haengen bleibt.
            await self._sitzung_abbrechen(upload_url)
            raise

        if not letzte_antwort:
            raise _to_typed(500, "files.upload(gross): Sitzung endete ohne driveItem")
        return letzte_antwort

    async def _sitzung_abbrechen(self, upload_url: str) -> None:
        """Bricht eine Upload-Sitzung ab (best-effort — ein Fehler hier darf nichts ueberdecken)."""
        try:
            async with self._get_session().delete(upload_url) as resp:
                logger.info("Upload-Sitzung abgebrochen (HTTP %s)", resp.status)
        except Exception as exc:      # noqa: BLE001 — bewusst geschluckt
            logger.warning("Upload-Sitzung konnte nicht abgebrochen werden: %s", exc)

    # --- Lesen / Aufraeumen --------------------------------------------------------------------

    async def get_item(self, drive_id: str, pfad: str) -> dict:
        """Metadaten einer Datei/eines Ordners. Wirft `M365NotFoundError`, wenn es sie nicht gibt."""
        p = pfad_pruefen(pfad)
        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/drives/{drive_id}/root:/{_pfad_kodieren(p)}",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id,name,size,webUrl,lastModifiedDateTime,folder,file"},
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "files.get_item", await resp.text())
            return await resp.json()

    async def list_children(self, drive_id: str, ordner_pfad: str = "", limit: int = 200) -> list[dict]:
        """Inhalt eines Ordners (leerer Pfad = Wurzel des Drives)."""
        token = await self._auth.get_token()
        if ordner_pfad.strip("/"):
            p = pfad_pruefen(ordner_pfad)
            url = f"{_GRAPH}/drives/{drive_id}/root:/{_pfad_kodieren(p)}:/children"
        else:
            url = f"{_GRAPH}/drives/{drive_id}/root/children"
        async with self._get_session().get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params={"$top": limit, "$select": "id,name,size,webUrl,lastModifiedDateTime,folder,file"},
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "files.list_children", await resp.text())
            return (await resp.json()).get("value", [])

    async def delete_item(self, drive_id: str, pfad: str) -> None:
        """Loescht eine Datei/einen Ordner (landet im Papierkorb der Site).

        Gebraucht fuer das **Aufraeumen nach Tests** — es bleibt kein Datenmuell im echten
        OneDrive zurueck (Nacht-Regel: niemals Demo-Dateien liegen lassen).
        """
        p = pfad_pruefen(pfad)
        token = await self._auth.get_token()
        async with self._get_session().delete(
            f"{_GRAPH}/drives/{drive_id}/root:/{_pfad_kodieren(p)}",
            headers={"Authorization": f"Bearer {token}"},
        ) as resp:
            if resp.status not in (200, 204, 404):
                raise _to_typed(resp.status, "files.delete_item", await resp.text())
