"""Microsoft Graph — Mail-Operationen."""

from __future__ import annotations

from typing import Callable

import aiohttp

from .auth import M365Auth

_GRAPH = "https://graph.microsoft.com/v1.0"


class MailService:
    def __init__(self, auth: M365Auth, get_session: Callable[[], aiohttp.ClientSession]) -> None:
        self._auth = auth
        self._get_session = get_session

    async def send(
        self,
        mailbox: str,
        to: str | list[str],
        subject: str,
        body: str,
        content_type: str = "HTML",
        cc: str | list[str] | None = None,
    ) -> None:
        """Sends an email from the given mailbox."""
        recipients = [to] if isinstance(to, str) else to
        payload: dict = {
            "message": {
                "subject": subject,
                "body": {"contentType": content_type, "content": body},
                "toRecipients": [
                    {"emailAddress": {"address": addr}} for addr in recipients
                ],
            }
        }
        if cc:
            cc_list = [cc] if isinstance(cc, str) else cc
            payload["message"]["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc_list
            ]

        token = await self._auth.get_token()
        async with self._get_session().post(
            f"{_GRAPH}/users/{mailbox}/sendMail",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
        ) as resp:
            if resp.status not in (200, 202):
                raise RuntimeError(f"mail.send failed ({resp.status})")

    async def list_inbox(
        self,
        mailbox: str,
        limit: int = 10,
        unread_only: bool = False,
    ) -> list[dict]:
        """Lists messages in the inbox of the given mailbox."""
        params: dict = {
            "$top": limit,
            "$select": "id,subject,from,receivedDateTime,bodyPreview,isRead",
            "$orderby": "receivedDateTime DESC",
        }
        if unread_only:
            params["$filter"] = "isRead eq false"

        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/users/{mailbox}/mailFolders/inbox/messages",
            headers={"Authorization": f"Bearer {token}"},
            params=params,
        ) as resp:
            if resp.status != 200:
                raise RuntimeError(f"mail.list_inbox failed ({resp.status})")
            data = await resp.json()
            return data.get("value", [])

    async def get_message(self, mailbox: str, message_id: str) -> dict:
        """Fetches a single message including full body."""
        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/users/{mailbox}/messages/{message_id}",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id,subject,from,receivedDateTime,body,isRead"},
        ) as resp:
            if resp.status != 200:
                raise RuntimeError(f"mail.get_message failed ({resp.status})")
            return await resp.json()

    async def mark_as_read(self, mailbox: str, message_id: str) -> None:
        """Marks a message as read."""
        token = await self._auth.get_token()
        async with self._get_session().patch(
            f"{_GRAPH}/users/{mailbox}/messages/{message_id}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"isRead": True},
        ) as resp:
            if resp.status != 200:
                raise RuntimeError(f"mail.mark_as_read failed ({resp.status})")

    async def move_to_folder(self, mailbox: str, message_id: str, folder: str) -> None:
        """Moves a message to a folder (e.g. 'deleteditems', 'archive')."""
        token = await self._auth.get_token()
        async with self._get_session().post(
            f"{_GRAPH}/users/{mailbox}/messages/{message_id}/move",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"destinationId": folder},
        ) as resp:
            if resp.status != 201:
                raise RuntimeError(f"mail.move_to_folder failed ({resp.status})")
