"""Logging setup – configures a rich-powered handler for the ``charter`` logger."""

from __future__ import annotations

import logging

from rich.logging import RichHandler

_configured = False


def setup(verbose: bool = False) -> None:
    """Initialise the ``charter`` logger.

    Call once at startup.  When *verbose* is ``True`` the level is set to
    ``DEBUG``; otherwise only ``WARNING`` and above are shown.
    """
    global _configured
    if _configured:
        return
    _configured = True

    level = logging.DEBUG if verbose else logging.WARNING
    handler = RichHandler(
        rich_tracebacks=True,
        show_time=True,
        show_path=False,
        markup=True,
    )
    handler.setLevel(level)

    logger = logging.getLogger("charter")
    logger.setLevel(level)
    logger.addHandler(handler)


def get(name: str) -> logging.Logger:
    """Return a child logger under ``charter.<name>``."""
    return logging.getLogger(f"charter.{name}")
