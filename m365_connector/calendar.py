"""Microsoft Graph — Kalender-Operationen."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Callable

import aiohttp

from .auth import M365Auth
from ._http import to_typed as _to_typed

_GRAPH = "https://graph.microsoft.com/v1.0"


def _to_utc_iso(dt: datetime) -> str:
    """Konvertiert datetime zu UTC-ISO-String für die Graph API.
    Naive datetimes werden als UTC interpretiert.
    """
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")


class CalendarService:
    def __init__(self, auth: M365Auth, get_session: Callable[[], aiohttp.ClientSession]) -> None:
        self._auth = auth
        self._get_session = get_session

    async def list_events(
        self,
        user: str,
        start: datetime,
        end: datetime,
        limit: int = 50,
    ) -> list[dict]:
        """Lists calendar events in the given time range."""
        token = await self._auth.get_token()
        async with self._get_session().get(
            f"{_GRAPH}/users/{user}/calendarView",
            headers={
                "Authorization": f"Bearer {token}",
                "Prefer": 'outlook.timezone="UTC"',
            },
            params={
                "startDateTime": _to_utc_iso(start),
                "endDateTime": _to_utc_iso(end),
                "$top": limit,
                "$select": "id,subject,start,end,organizer,attendees,bodyPreview",
                "$orderby": "start/dateTime ASC",
            },
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "calendar.list_events")
            data = await resp.json()
            return data.get("value", [])

    async def create_event(
        self,
        user: str,
        subject: str,
        start: datetime,
        end: datetime,
        body: str = "",
        attendees: list[str] | None = None,
        location: str | None = None,
    ) -> dict:
        """Creates a new calendar event."""
        payload: dict = {
            "subject": subject,
            "start": {"dateTime": _to_utc_iso(start), "timeZone": "UTC"},
            "end": {"dateTime": _to_utc_iso(end), "timeZone": "UTC"},
        }
        if body:
            payload["body"] = {"contentType": "HTML", "content": body}
        if attendees:
            payload["attendees"] = [
                {"emailAddress": {"address": addr}, "type": "required"}
                for addr in attendees
            ]
        if location:
            payload["location"] = {"displayName": location}

        token = await self._auth.get_token()
        async with self._get_session().post(
            f"{_GRAPH}/users/{user}/events",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
        ) as resp:
            if resp.status != 201:
                raise _to_typed(resp.status, "calendar.create_event")
            return await resp.json()

    async def delete_event(self, user: str, event_id: str) -> None:
        """Deletes a calendar event."""
        token = await self._auth.get_token()
        async with self._get_session().delete(
            f"{_GRAPH}/users/{user}/events/{event_id}",
            headers={"Authorization": f"Bearer {token}"},
        ) as resp:
            if resp.status != 204:
                raise _to_typed(resp.status, "calendar.delete_event")
