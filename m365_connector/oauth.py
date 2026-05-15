"""OAuth/Consent helpers — wiederverwendbare Bausteine für Setup-Flows.

Diese Funktionen sind unabhängig von Browser, CLI oder einer bestimmten Web-Framework-
Integration. Sie können vom CLI-Wizard genauso wie von einem Web-Portal-Backend
verwendet werden, um den Microsoft-Admin-Consent-Flow zu treiben.
"""

from __future__ import annotations

import urllib.parse
from collections.abc import Mapping

ADMIN_CONSENT_BASE = "https://login.microsoftonline.com/organizations/v2.0/adminconsent"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"


def build_admin_consent_url(
    client_id: str,
    redirect_uri: str,
    state: str | None = None,
    scope: str = GRAPH_SCOPE,
) -> str:
    """Baut die Microsoft Admin-Consent-URL.

    Args:
        client_id: Application (Client) ID aus der Azure App Registration.
        redirect_uri: Muss exakt mit einer in Azure konfigurierten Redirect-URI übereinstimmen.
        state: Beliebiger Wert, der im Callback zurückkommt (CSRF-Schutz / Tenant-Kontext).
        scope: Permission-Scope. Default = Graph .default (alle erteilten App-Permissions).

    Returns:
        Vollständige URL zum Aufruf im Browser.
    """
    params: dict[str, str] = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "scope": scope,
    }
    if state:
        params["state"] = state
    return f"{ADMIN_CONSENT_BASE}?{urllib.parse.urlencode(params)}"


def parse_consent_callback(query_params: Mapping[str, str]) -> dict:
    """Parst die Query-Params eines Admin-Consent-Callbacks.

    Microsoft schickt nach erfolgreichem Consent:
        admin_consent=True, tenant=<tenant-id>, state=<echoed-state>
    Bei Fehler:
        error=<code>, error_description=<text>

    Returns:
        Dict mit:
            - success (bool): admin_consent==True
            - tenant_id (str|None): Tenant-ID falls Success
            - error (str|None): Fehlercode falls vorhanden
            - error_description (str|None)
            - state (str|None): Zurückgegebener state-Parameter
    """
    success = query_params.get("admin_consent") == "True"
    return {
        "success": success,
        "tenant_id": query_params.get("tenant") if success else None,
        "error": query_params.get("error"),
        "error_description": query_params.get("error_description"),
        "state": query_params.get("state"),
    }
