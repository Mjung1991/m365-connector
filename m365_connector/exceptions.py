"""
M365 Connector Exceptions — typisierte Fehler statt generischer RuntimeError.
Erlaubt Callers `except M365ServiceError as e: if e.status_code == 404: ...`.
"""

from __future__ import annotations


class M365Error(Exception):
    """Basis-Klasse fuer alle M365-spezifischen Fehler."""


class M365ServiceError(M365Error):
    """
    HTTP-Fehler beim Aufruf von Microsoft Graph.
    Enthaelt status_code + Original-Message + (optional) Response-Body.
    """

    def __init__(self, status_code: int, message: str, response_body: str | None = None) -> None:
        self.status_code = status_code
        self.message = message
        self.response_body = response_body
        super().__init__(f"M365 {status_code}: {message}")


class M365AuthError(M365Error):
    """Token-Erwerb oder Refresh fehlgeschlagen."""


class M365NotFoundError(M365ServiceError):
    """404 — Resource nicht vorhanden (z.B. Mail bereits verschoben)."""

    def __init__(self, message: str, response_body: str | None = None) -> None:
        super().__init__(404, message, response_body)


class M365RateLimitError(M365ServiceError):
    """429 — Throttling. retry_after_seconds aus Retry-After-Header."""

    def __init__(self, message: str, retry_after_seconds: float, response_body: str | None = None) -> None:
        self.retry_after_seconds = retry_after_seconds
        super().__init__(429, message, response_body)


class M365ValidationError(M365ServiceError):
    """400/422 — Request invalid (z.B. ungueltige Subscription-URL)."""
