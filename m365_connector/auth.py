"""
Microsoft 365 Authentifizierung via Client Credentials (App-only / Daemon).
Nutzt azure-identity — Token-Caching und Auto-Refresh sind eingebaut.

Zwei App-only-Wege (Microsoft-empfohlen fuer unbeaufsichtigte Dienste):
  * **Client-Secret** (Standard, rueckwaertskompatibel): ``M365Auth(tenant, client, secret)``.
  * **Zertifikat** (fuer Dauerbetrieb empfohlen, MS): ``M365Auth.from_certificate(...)`` bzw.
    ``M365Auth(tenant, client, certificate_path=...)``. Der private Schluessel gehoert in einen
    Secrets-Tresor (z.B. Infisical) und wird regelmaessig rotiert.
"""

from azure.identity.aio import CertificateCredential, ClientSecretCredential

GRAPH_SCOPE = "https://graph.microsoft.com/.default"


class M365Auth:
    """App-only-Authentifizierung gegen Microsoft Graph — per Secret ODER Zertifikat.

    Genau EINE Quelle muss gesetzt sein: entweder ``client_secret`` oder ein Zertifikat
    (``certificate_path`` oder ``certificate_data``). Beides gleichzeitig oder gar nichts ist
    ein Fehler (fail-closed) — verhindert unklare Konfiguration im Produktivbetrieb.
    """

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str | None = None,
        *,
        certificate_path: str | None = None,
        certificate_data: bytes | None = None,
        certificate_password: str | bytes | None = None,
        send_certificate_chain: bool = False,
    ) -> None:
        has_cert = bool(certificate_path) or certificate_data is not None
        has_secret = bool(client_secret)
        if has_cert and has_secret:
            raise ValueError(
                "Provide either client_secret OR a certificate (certificate_path/"
                "certificate_data), not both."
            )
        if not has_cert and not has_secret:
            raise ValueError(
                "Missing credentials: provide client_secret or a certificate "
                "(certificate_path or certificate_data)."
            )

        if has_cert:
            cert_kwargs: dict = {}
            if certificate_data is not None:
                cert_kwargs["certificate_data"] = certificate_data
            if certificate_password is not None:
                cert_kwargs["password"] = certificate_password
            if send_certificate_chain:
                cert_kwargs["send_certificate_chain"] = True
            self._credential = CertificateCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                certificate_path=certificate_path,
                **cert_kwargs,
            )
            self.auth_mode = "certificate"
        else:
            self._credential = ClientSecretCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
            )
            self.auth_mode = "secret"

    @classmethod
    def from_certificate(
        cls,
        tenant_id: str,
        client_id: str,
        *,
        certificate_path: str | None = None,
        certificate_data: bytes | None = None,
        password: str | bytes | None = None,
        send_certificate_chain: bool = False,
    ) -> "M365Auth":
        """Erzeugt eine zertifikatsbasierte App-only-Authentifizierung.

        Genau eine Schluesselquelle: ``certificate_path`` (PEM/PFX auf Platte) ODER
        ``certificate_data`` (Bytes, z.B. aus dem Secrets-Tresor). ``password`` nur, wenn der
        private Schluessel verschluesselt ist. ``send_certificate_chain=True`` fuer Subject-Name-
        und-Issuer-Auth (SNI), falls der Tenant das verlangt.
        """
        return cls(
            tenant_id,
            client_id,
            certificate_path=certificate_path,
            certificate_data=certificate_data,
            certificate_password=password,
            send_certificate_chain=send_certificate_chain,
        )

    async def get_token(self) -> str:
        """Gibt einen gültigen Bearer-Token zurück. Wird automatisch erneuert."""
        token = await self._credential.get_token(GRAPH_SCOPE)
        return token.token

    async def close(self) -> None:
        await self._credential.close()

    async def __aenter__(self) -> "M365Auth":
        return self

    async def __aexit__(self, *_) -> None:
        await self.close()
