"""Tests fuer FilesService — Datei-Ablage in SharePoint/OneDrive.

Schwerpunkt liegt auf den drei Fallen aus dem Modul-Kopf von `files.py`, weil genau die
in der Praxis Daten beschaedigen bzw. mitten im Upload zu 401 fuehren:
  1. Fragment-PUT darf KEINEN Authorization-Header tragen
  2. Fragmentgroesse muss ein Vielfaches von 320 KiB sein
  3. Datei-Inhalt geht als rohe Bytes raus (`data=`), niemals als `json=`
"""

import pytest
from unittest.mock import AsyncMock, MagicMock

import aiohttp

from m365_connector.auth import M365Auth
from m365_connector.exceptions import M365ServiceError
from m365_connector.files import (
    FRAGMENT_EINHEIT,
    STANDARD_FRAGMENT,
    FilesService,
    M365PfadFehler,
    fragmente,
    pfad_pruefen,
)


@pytest.fixture
def auth():
    a = MagicMock(spec=M365Auth)
    a.get_token = AsyncMock(return_value="fake-token")
    return a


@pytest.fixture
def mock_session():
    return MagicMock(spec=aiohttp.ClientSession)


@pytest.fixture
def service(auth, mock_session):
    return FilesService(auth, lambda: mock_session)


def _resp(status: int, json_data: dict | None = None, text: str = ""):
    r = AsyncMock()
    r.status = status
    r.json = AsyncMock(return_value=json_data or {})
    r.text = AsyncMock(return_value=text)
    r.__aenter__ = AsyncMock(return_value=r)
    r.__aexit__ = AsyncMock(return_value=False)
    return r


# --- Pfad-Pruefung (rein, kein Netz) -----------------------------------------------------------

def test_pfad_normalisiert_schraegstriche():
    assert pfad_pruefen("/Q-Ablage/Vorlagen/x.pdf/") == "Q-Ablage/Vorlagen/x.pdf"
    assert pfad_pruefen("Q-Ablage\\Vorlagen\\x.pdf") == "Q-Ablage/Vorlagen/x.pdf"


@pytest.mark.parametrize("schlecht", [
    "",
    "   ",
    "Q-Ablage//x.pdf",           # leerer Abschnitt
    "Q-Ablage/../../geheim.pdf", # Ausbruch aus dem Zielordner
    "Q-Ablage/./x.pdf",
    'Q-Ablage/bericht:2026.pdf', # verbotenes Zeichen
    "Q-Ablage/was<>.pdf",
    "Q-Ablage/ fuehrendes-leerzeichen.pdf",
    "Q-Ablage/endet-mit-punkt.",
])
def test_pfad_weist_unzulaessiges_ab(schlecht):
    with pytest.raises(M365PfadFehler):
        pfad_pruefen(schlecht)


def test_pfad_zu_lang():
    with pytest.raises(M365PfadFehler, match="zu lang"):
        pfad_pruefen("a/" * 250 + "x.pdf")


# --- Fragmentierung (Falle 2) ------------------------------------------------------------------

def test_fragmente_sind_vielfache_von_320kib():
    gesamt = 25 * 1024 * 1024
    teile = fragmente(gesamt)
    # Alle ausser dem letzten muessen exakt durch 320 KiB teilbar sein
    for start, ende in teile[:-1]:
        assert (ende - start + 1) % FRAGMENT_EINHEIT == 0
    # Luecken- und ueberlappungsfrei, deckt die ganze Datei ab
    assert teile[0][0] == 0
    assert teile[-1][1] == gesamt - 1
    for vorher, nachher in zip(teile, teile[1:]):
        assert nachher[0] == vorher[1] + 1


def test_fragmentgroesse_muss_teilbar_sein():
    # 1 MB ist KEIN Vielfaches von 320 KiB — das muss sofort auffallen, nicht erst beim
    # Zusammensetzen der Datei durch Graph.
    with pytest.raises(ValueError, match="Vielfaches von 320 KiB"):
        fragmente(10_000_000, fragment_groesse=1_000_000)


def test_standard_fragment_ist_teilbar():
    assert STANDARD_FRAGMENT % FRAGMENT_EINHEIT == 0
    assert STANDARD_FRAGMENT == 10 * 1024 * 1024


# --- Ziel finden -------------------------------------------------------------------------------

async def test_site_id_aufloesen(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {"id": "host,guid1,guid2"}))
    sid = await service.site_id("firma.sharepoint.com", "/sites/Q-Ablage")
    assert sid == "host,guid1,guid2"
    url = mock_session.get.call_args[0][0]
    assert url.endswith("/sites/firma.sharepoint.com:/sites/Q-Ablage")


async def test_site_id_fehler_wird_typisiert(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(401, text='{"error":{"code":"generalException"}}'))
    with pytest.raises(M365ServiceError) as exc:
        await service.site_id("firma.sharepoint.com", "/sites/Q-Ablage")
    assert exc.value.status_code == 401


# --- Ordner anlegen ----------------------------------------------------------------------------

async def test_ensure_folder_legt_jede_ebene_an(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(201, {"id": "f1"}))
    await service.ensure_folder("drive1", "Q-Ablage/Vorlagen/Angebote")
    assert mock_session.post.call_count == 3


async def test_ensure_folder_ist_idempotent(service, mock_session):
    # 409 = Ordner existiert schon. Das ist der Normalfall und darf NICHT scheitern.
    mock_session.post = MagicMock(return_value=_resp(409, text="already exists"))
    ergebnis = await service.ensure_folder("drive1", "Q-Ablage/Vorlagen")
    assert ergebnis["existierte_schon"] is True


async def test_ensure_folder_ueberschreibt_nie(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(201, {"id": "f1"}))
    await service.ensure_folder("drive1", "Q-Ablage")
    body = mock_session.post.call_args.kwargs["json"]
    # 'replace' wuerde den Inhalt eines bestehenden Ordners verlieren
    assert body["@microsoft.graph.conflictBehavior"] == "fail"


# --- Kleiner Upload (Falle 3) ------------------------------------------------------------------

async def test_upload_klein_sendet_rohe_bytes_nicht_json(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(201, {"id": "f1"}))
    mock_session.put = MagicMock(return_value=_resp(201, {"id": "datei1", "name": "x.pdf"}))
    inhalt = b"%PDF-1.7 echte bytes \x00\x01\x02"
    ergebnis = await service.upload("drive1", "Q-Ablage/PDF/x.pdf", inhalt,
                                    content_type="application/pdf")
    assert ergebnis["id"] == "datei1"
    kwargs = mock_session.put.call_args.kwargs
    # DAS ist der Kern: die Datei geht als data= raus, nicht als json=
    assert kwargs["data"] == inhalt
    assert "json" not in kwargs
    assert kwargs["headers"]["Content-Type"] == "application/pdf"


async def test_upload_setzt_conflict_behavior(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(201, {"id": "f1"}))
    mock_session.put = MagicMock(return_value=_resp(200, {"id": "d"}))
    await service.upload("drive1", "Q-Ablage/PDF/x.pdf", b"abc", conflict_behavior="rename")
    params = mock_session.put.call_args.kwargs["params"]
    assert params["@microsoft.graph.conflictBehavior"] == "rename"


async def test_upload_weist_text_ab(service):
    with pytest.raises(TypeError, match="bytes"):
        await service.upload("drive1", "Q-Ablage/x.txt", "kein bytes")


async def test_upload_weist_leeren_inhalt_ab(service):
    with pytest.raises(ValueError, match="leerer Inhalt"):
        await service.upload("drive1", "Q-Ablage/x.txt", b"")


async def test_upload_weist_unbekanntes_conflict_behavior_ab(service):
    with pytest.raises(ValueError, match="conflict_behavior"):
        await service.upload("drive1", "Q-Ablage/x.txt", b"abc", conflict_behavior="ueberschreiben")


# --- Grosser Upload (Falle 1) ------------------------------------------------------------------

async def test_upload_gross_nutzt_sitzung_und_sendet_keinen_token(service, mock_session):
    mock_session.post = MagicMock(side_effect=[
        _resp(201, {"id": "o1"}),                                      # ensure_folder: Q-Ablage
        _resp(201, {"id": "o2"}),                                      # ensure_folder: Q-Ablage/PDF
        _resp(200, {"uploadUrl": "https://up.example/sitzung"}),       # createUploadSession
    ])
    # 12 MB -> zwei Fragmente (10 MiB + Rest)
    inhalt = b"x" * (12 * 1024 * 1024)
    mock_session.put = MagicMock(side_effect=[
        _resp(202, {"nextExpectedRanges": ["10485760-"]}),
        _resp(201, {"id": "grossdatei", "name": "gross.pdf"}),
    ])

    ergebnis = await service.upload("drive1", "Q-Ablage/PDF/gross.pdf", inhalt)
    assert ergebnis["id"] == "grossdatei"
    assert mock_session.put.call_count == 2

    for aufruf in mock_session.put.call_args_list:
        headers = aufruf.kwargs["headers"]
        # Falle 1: ein mitgesendeter Token laesst Graph mitten im Upload mit 401 antworten
        assert "Authorization" not in headers
        assert headers["Content-Range"].startswith("bytes ")
        assert headers["Content-Range"].endswith(f"/{len(inhalt)}")


async def test_upload_gross_bricht_sitzung_bei_fehler_ab(service, mock_session):
    mock_session.post = MagicMock(side_effect=[
        _resp(201, {"id": "o1"}),                                      # ensure_folder: Q-Ablage
        _resp(201, {"id": "o2"}),                                      # ensure_folder: Q-Ablage/PDF
        _resp(200, {"uploadUrl": "https://up.example/sitzung"}),
    ])
    mock_session.put = MagicMock(return_value=_resp(500, text="server kaputt"))
    mock_session.delete = MagicMock(return_value=_resp(204))

    with pytest.raises(M365ServiceError):
        await service.upload("drive1", "Q-Ablage/PDF/gross.pdf", b"x" * (5 * 1024 * 1024))

    # Sonst bleibt im Ziel eine halbe Datei haengen
    mock_session.delete.assert_called_once_with("https://up.example/sitzung")


async def test_kleine_datei_nutzt_keine_sitzung(service, mock_session):
    mock_session.post = MagicMock(return_value=_resp(201, {"id": "ordner"}))
    mock_session.put = MagicMock(return_value=_resp(201, {"id": "d"}))
    await service.upload("drive1", "Q-Ablage/PDF/klein.pdf", b"x" * 1024)
    # Nur ensure_folder darf POSTen — keine createUploadSession
    for aufruf in mock_session.post.call_args_list:
        assert "createUploadSession" not in aufruf[0][0]


# --- Aufraeumen --------------------------------------------------------------------------------

async def test_delete_item_schluckt_404(service, mock_session):
    # Schon weg = Ziel erreicht. Aufraeumen darf nie scheitern, nur weil es nichts zu tun gab.
    mock_session.delete = MagicMock(return_value=_resp(404))
    await service.delete_item("drive1", "Q-Ablage/Test/weg.pdf")


async def test_delete_item_meldet_echten_fehler(service, mock_session):
    mock_session.delete = MagicMock(return_value=_resp(403, text="verboten"))
    with pytest.raises(M365ServiceError):
        await service.delete_item("drive1", "Q-Ablage/Test/x.pdf")


async def test_list_children_wurzel(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {"value": [{"id": "1", "name": "a"}]}))
    eintraege = await service.list_children("drive1")
    assert len(eintraege) == 1
    assert mock_session.get.call_args[0][0].endswith("/root/children")


async def test_list_children_leerer_ordner(service, mock_session):
    mock_session.get = MagicMock(return_value=_resp(200, {}))
    assert await service.list_children("drive1", "Q-Ablage/Leer") == []
