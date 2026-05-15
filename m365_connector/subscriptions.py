"""Microsoft Graph — Subscriptions / Webhook-Operationen.

Subscriptions trigger HTTP notifications to a notificationUrl when a resource changes
(e.g. new mail arrives). On creation, Microsoft sends a validation handshake to the URL —
use `validate_subscription_token()` in the receiving endpoint to handle it correctly.
"""

from __future__ import annotations

from typing import Callable, Mapping

import aiohttp

from .auth import M365Auth

_GRAPH = "https://graph.microsoft.com/v1.0"


class SubscriptionService:
    """Manage Graph subscriptions (webhooks).

    Access via `client.subscriptions.X` — create, renew, delete, list, get.
    """

    def __init__(self, auth: M365Auth, get_session: Callable[[], aiohttp.ClientSession]) -> None:
        self._auth = auth
        self._get_session = get_session

    async def create(
        self,
        resource: str,
        notification_url: str,
        expires: str,
        client_state: str,
        change_type: str = "created,updated",
        lifecycle_notification_url: str | None = None,
    ) -> dict:
        """Creates a Graph subscription.

        Args:
            resource: Graph resource path (e.g. `users/{id}/mailFolders('Inbox')/messages`).
            notification_url: HTTPS endpoint that will receive change notifications.
            expires: ISO8601 expiration timestamp (max ~3 days for mail resources).
            client_state: Secret string echoed back in every notification — verify on receive.
            change_type: Comma-separated list: "created", "updated", "deleted".
            lifecycle_notification_url: Optional separate URL for lifecycle events (subscription
                expiring soon, reauth needed).

        Returns:
            Subscription object with id, expirationDateTime, applicationId, etc.
        """
        payload: dict = {
            "changeType": change_type,
            "notificationUrl": notification_url,
            "resource": resource,
            "expirationDateTime": expires,
            "clientState": client_state,
        }
        if lifecycle_notification_url:
            payload["lifecycleNotificationUrl"] = lifecycle_notification_url

        token = await self._auth.get_token()
        async with self._get_session().post(
            f"{_GRAPH}/subscriptions",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
        ) as resp:
            if resp.status != 201:
                raise RuntimeError(f"subscriptions.create failed ({resp.status})")
            return await resp.json()

    async def renew(self, subscription_id: str, expires: str) -> dict:
        """Extends the expirationDateTime of an existing subscription.

        Mail subscriptions max ~3 days — renew before expires or you lose the stream.
        """
        token = await self._auth.get_token()
        async with self._get_session().patch(
            f"{_GRAPH}/subscriptions/{subscription_id}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"expirationDateTime": expires},
        ) as resp:
            if resp.status != 200:
                raise RuntimeError(f"subscriptions.renew failed ({resp.status})")
            return await resp.json()

    async def delete(self, subscription_id: str) -> None:
        """Deletes an existing subscription (stops notifications)."""
        token = await self._auth.get_token()
        async with self._get_session().delete(
            f"{_GRAPH}/subscriptions/{subscription_id}",
            headers={"Authorization": f"Bearer {token}"},
        ) as resp:
            if resp.status not in (200, 204):
                raise RuntimeError(f"subscriptions.delete failed ({resp.status})")

    async def list(self) -> list[dict]:
        """Lists all subscriptions for the current application."""
        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/subscriptions",
            headers={"Authorization": f"Bearer {token}"},
        ) as resp:
            if resp.status != 200:
                raise RuntimeError(f"subscriptions.list failed ({resp.status})")
            data = await resp.json()
            return data.get("value", [])

    async def get(self, subscription_id: str) -> dict:
        """Returns details of a single subscription."""
        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/subscriptions/{subscription_id}",
            headers={"Authorization": f"Bearer {token}"},
        ) as resp:
            if resp.status != 200:
                raise RuntimeError(f"subscriptions.get failed ({resp.status})")
            return await resp.json()


def validate_subscription_token(query_params: Mapping[str, str]) -> str | None:
    """Returns the validationToken if Microsoft sent a handshake request.

    On subscription creation Microsoft GETs the notificationUrl with a `validationToken`
    query parameter and expects the same value echoed back as text/plain within 10 seconds.

    Usage in a FastAPI/Starlette webhook handler::

        @app.post("/webhook/m365")
        async def webhook(request: Request):
            token = validate_subscription_token(dict(request.query_params))
            if token is not None:
                return PlainTextResponse(token, status_code=200)
            # ... process actual notification body ...

    Returns:
        Token string if handshake request, None if normal notification call.
    """
    return query_params.get("validationToken")
