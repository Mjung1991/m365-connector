"""Microsoft Graph — Mail-Operationen."""

from __future__ import annotations

from typing import Callable

import aiohttp

from .auth import M365Auth
from .delta import MailDeltaService
from .folders import MailFolderService
from ._http import to_typed as _to_typed

_GRAPH = "https://graph.microsoft.com/v1.0"


class MailService:
    def __init__(self, auth: M365Auth, get_session: Callable[[], aiohttp.ClientSession]) -> None:
        self._auth = auth
        self._get_session = get_session
        self.folders = MailFolderService(auth, get_session)
        self.delta = MailDeltaService(auth, get_session)

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
                raise _to_typed(resp.status, "mail.send")

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
                raise _to_typed(resp.status, "mail.list_inbox")
            data = await resp.json()
            return data.get("value", [])

    async def get_message(self, mailbox: str, message_id: str) -> dict:
        """Fetches a single message including full body."""
        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/users/{mailbox}/messages/{message_id}",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id,subject,from,receivedDateTime,body,isRead,hasAttachments"},
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "mail.get_message")
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
                raise _to_typed(resp.status, "mail.mark_as_read")

    async def move_to_folder(self, mailbox: str, message_id: str, folder: str) -> None:
        """Moves a message to a folder (e.g. 'deleteditems', 'archive')."""
        token = await self._auth.get_token()
        async with self._get_session().post(
            f"{_GRAPH}/users/{mailbox}/messages/{message_id}/move",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"destinationId": folder},
        ) as resp:
            if resp.status != 201:
                raise _to_typed(resp.status, "mail.move_to_folder")

    async def list_messages(
        self,
        mailbox: str,
        folder: str | None = None,
        limit: int = 50,
        unread_only: bool = False,
        page_token: str | None = None,
    ) -> tuple[list[dict], str | None]:
        """Lists messages with pagination.

        Args:
            mailbox: Postfach-Adresse oder User-ID.
            folder: Folder-ID oder Wellname ('inbox', 'archive', 'sentitems', etc.).
                    None = alle Nachrichten im Postfach.
            limit: Seitengröße.
            unread_only: Filtert auf ungelesene Mails.
            page_token: @odata.nextLink aus vorherigem Aufruf zum Weiterblättern.

        Returns:
            Tuple aus (messages, next_page_token). next_page_token=None wenn letzte Seite.
        """
        token = await self._auth.get_token()
        headers = {"Authorization": f"Bearer {token}"}

        if page_token:
            url = page_token
            params: dict = {}
        else:
            if folder:
                url = f"{_GRAPH}/users/{mailbox}/mailFolders/{folder}/messages"
            else:
                url = f"{_GRAPH}/users/{mailbox}/messages"
            params = {
                "$top": limit,
                "$select": "id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments",
                "$orderby": "receivedDateTime DESC",
            }
            if unread_only:
                params["$filter"] = "isRead eq false"

        async with self._get_session().get(url, headers=headers, params=params) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "mail.list_messages")
            data = await resp.json()
            return data.get("value", []), data.get("@odata.nextLink")

    async def fetch_attachments(self, mailbox: str, message_id: str) -> list[dict]:
        """Returns all attachments of a message (with content as base64)."""
        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/users/{mailbox}/messages/{message_id}/attachments",
            headers={"Authorization": f"Bearer {token}"},
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "mail.fetch_attachments")
            data = await resp.json()
            return data.get("value", [])

    async def send_forward(
        self,
        mailbox: str,
        message_id: str,
        to: str | list[str],
        comment: str | None = None,
    ) -> None:
        """Forwards a message to one or more recipients."""
        recipients = [to] if isinstance(to, str) else to
        payload: dict = {
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in recipients
            ],
        }
        if comment is not None:
            payload["comment"] = comment

        token = await self._auth.get_token()
        async with self._get_session().post(
            f"{_GRAPH}/users/{mailbox}/messages/{message_id}/forward",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
        ) as resp:
            if resp.status not in (200, 202):
                raise _to_typed(resp.status, "mail.send_forward")

    async def move_batch(
        self,
        mailbox: str,
        message_ids: list[str],
        destination_folder: str,
    ) -> list[dict]:
        """Moves multiple messages in a single Graph $batch request.

        Returns:
            Liste der Batch-Responses (eine pro message_id, in gleicher Reihenfolge).
            Jeder Eintrag enthält id, status und ggf. body.
        """
        if not message_ids:
            return []
        if len(message_ids) > 20:
            raise ValueError("Graph $batch supports max 20 requests per call")

        requests = [
            {
                "id": str(idx),
                "method": "POST",
                "url": f"/users/{mailbox}/messages/{mid}/move",
                "headers": {"Content-Type": "application/json"},
                "body": {"destinationId": destination_folder},
            }
            for idx, mid in enumerate(message_ids)
        ]

        token = await self._auth.get_token()
        async with self._get_session().post(
            f"{_GRAPH}/$batch",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"requests": requests},
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "mail.move_batch")
            data = await resp.json()
            responses = data.get("responses", [])
            responses.sort(key=lambda r: int(r["id"]))
            return responses
