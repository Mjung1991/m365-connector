"""
Microsoft 365 Authentifizierung via Client Credentials (App-only / Daemon).
Nutzt azure-identity — Token-Caching und Auto-Refresh sind eingebaut.
"""

from azure.identity.aio import ClientSecretCredential

GRAPH_SCOPE = "https://graph.microsoft.com/.default"


class M365Auth:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str) -> None:
        self._credential = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
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
