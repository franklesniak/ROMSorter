"""Provide a Pylance-friendly pytest import for repository tests."""

from __future__ import annotations

import importlib
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import pytest
else:
    pytest = importlib.import_module("pytest")

__all__ = ["pytest"]
