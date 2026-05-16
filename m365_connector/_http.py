"""
Interne HTTP-Helfer fuer M365Connector.
_request_with_retry: respektiert Retry-After-Header bei 429, exponential backoff fuer 5xx.
"""

from __future__ import annotations

import asyncio
import logging
from typing import Any, Callable

import aiohttp

from .auth import M365Auth
from .exceptions import (
    M365NotFoundError,
    M365RateLimitError,
    M365ServiceError,
    M365ValidationError,
)

logger = logging.getLogger(__name__)


# Default-Timeout fuer aiohttp.ClientSession — verhindert haengende Calls bei grossen Attachments
DEFAULT_TIMEOUT_SECONDS = 60


def make_session(timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> aiohttp.ClientSession:
    """Erzeugt aiohttp.ClientSession mit konservativem Total-Timeout."""
    return aiohttp.ClientSession(
        timeout=aiohttp.ClientTimeout(total=timeout_seconds),
    )


def _raise_for_status(status: int, body_text: str) -> None:
    """Wirft typisierte M365-Exception passend zum HTTP-Status."""
    if status == 404:
        raise M365NotFoundError(body_text[:500])
    if status in (400, 422):
        raise M365ValidationError(status, body_text[:500])
    if 200 <= status < 300:
        return
    raise M365ServiceError(status, body_text[:500])


async def request_with_retry(
    session: aiohttp.ClientSession,
    method: str,
    url: str,
    *,
    auth: M365Auth,
    json_body: Any = None,
    params: dict | None = None,
    extra_headers: dict | None = None,
    max_attempts: int = 3,
    backoff_base_seconds: float = 2.0,
) -> tuple[int, str, dict]:
    """
    Fuehrt HTTP-Request mit Retry-Logik aus.
    Bei 429: parst Retry-After-Header (Sekunden oder HTTP-Date), wartet, retry.
    Bei 5xx: exponential backoff (2, 4, 8 Sekunden).
    Bei 4xx (ausser 429): wirft sofort die passende Exception.

    Returns: (status, response_text, headers_dict)
    Headers werden zurueckgegeben damit Caller z.B. ETag, X-Ms-* lesen koennen.
    """
    last_exc: Exception | None = None
    for attempt in range(1, max_attempts + 1):
        token = await auth.get_token()
        headers = {"Authorization": f"Bearer {token}"}
        if json_body is not None:
            headers["Content-Type"] = "application/json"
        if extra_headers:
            headers.update(extra_headers)

        try:
            async with session.request(
                method, url,
                headers=headers,
                json=json_body,
                params=params,
            ) as resp:
                body_text = await resp.text()
                hdrs = dict(resp.headers)

                if resp.status == 429:
                    retry_after = _parse_retry_after(resp.headers.get("Retry-After"))
                    logger.warning(
                        "M365 429 throttled (attempt %d/%d) — wait %.1fs",
                        attempt, max_attempts, retry_after,
                    )
                    if attempt >= max_attempts:
                        raise M365RateLimitError(body_text[:500], retry_after, body_text)
                    await asyncio.sleep(retry_after)
                    continue

                if 500 <= resp.status < 600:
                    if attempt < max_attempts:
                        wait = backoff_base_seconds ** attempt
                        logger.warning(
                            "M365 %d server error (attempt %d/%d) — backoff %.1fs",
                            resp.status, attempt, max_attempts, wait,
                        )
                        await asyncio.sleep(wait)
                        continue
                    _raise_for_status(resp.status, body_text)

                # 2xx oder andere 4xx → return / raise typed
                _raise_for_status(resp.status, body_text)
                return resp.status, body_text, hdrs

        except aiohttp.ClientError as exc:
            last_exc = exc
            if attempt < max_attempts:
                wait = backoff_base_seconds ** attempt
                logger.warning(
                    "M365 network error (attempt %d/%d): %s — backoff %.1fs",
                    attempt, max_attempts, exc, wait,
                )
                await asyncio.sleep(wait)
                continue
            raise M365ServiceError(0, f"Network error after {max_attempts} attempts: {exc}") from exc

    # Sollte nie erreicht werden
    raise M365ServiceError(0, f"request_with_retry exhausted: {last_exc}")


def to_typed(status: int, op: str, body: str = "") -> M365ServiceError:
    """Mapped HTTP-Status auf passende typed Exception. Helper fuer Service-Methoden,
    die `raise to_typed(resp.status, 'mail.send')` schreiben statt eines RuntimeError-Mustes."""
    msg = f"{op} ({status})" if not body else f"{op} ({status}): {body[:200]}"
    if status == 404:
        return M365NotFoundError(msg, body)
    if status in (400, 422):
        return M365ValidationError(status, msg, body)
    if status == 429:
        return M365RateLimitError(msg, 5.0, body)
    return M365ServiceError(status, msg, body)


def _parse_retry_after(value: str | None) -> float:
    """
    Retry-After kann eine Anzahl Sekunden ODER ein HTTP-Datum sein.
    Default-Wartezeit: 5s wenn nicht parsebar.
    """
    if not value:
        return 5.0
    try:
        return float(value)
    except ValueError:
        pass
    try:
        from email.utils import parsedate_to_datetime
        from datetime import datetime, timezone
        target = parsedate_to_datetime(value)
        now = datetime.now(timezone.utc)
        delta = (target - now).total_seconds()
        return max(1.0, delta)
    except Exception:
        return 5.0
