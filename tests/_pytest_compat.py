"""Provide a Pylance-friendly pytest import for repository tests."""

from __future__ import annotations

import importlib
from typing import Any, cast

pytest = cast(Any, importlib.import_module("pytest"))

__all__ = ["pytest"]
