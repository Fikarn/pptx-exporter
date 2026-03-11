"""Persistent user settings (~/.pptx-exporter-settings.json)."""

import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

_SETTINGS_PATH = Path.home() / ".pptx-exporter-settings.json"


def load_settings() -> dict:
    try:
        with open(_SETTINGS_PATH) as fh:
            return json.load(fh)
    except Exception:
        return {}


def save_settings(data: dict) -> None:
    try:
        with open(_SETTINGS_PATH, "w") as fh:
            json.dump(data, fh, indent=2)
    except Exception as exc:
        logger.debug("Could not save settings: %s", exc)
