"""
Shared guardrail limits and error taxonomy for resource safety.
"""

from __future__ import annotations

import os
from typing import Tuple


# Shared error taxonomy tags for deterministic client handling.
LIMIT_EXCEEDED = "LIMIT_EXCEEDED"
REGEX_COMPLEXITY_BLOCKED = "REGEX_COMPLEXITY_BLOCKED"
DOC_TOO_LARGE = "DOC_TOO_LARGE"
UNDO_BUDGET_EXCEEDED = "UNDO_BUDGET_EXCEEDED"
SESSION_CONSISTENCY_WARNING = "SESSION_CONSISTENCY_WARNING"


def env_int(name: str, default: int, minimum: int = 0) -> int:
    """Read an integer from the environment with a safe lower bound."""
    raw = os.getenv(name, str(default))
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return default
    if value < minimum:
        return minimum
    return value


def get_max_matches_per_call() -> int:
    return env_int("EW_MAX_MATCHES_PER_CALL", 1000, minimum=1)


def get_max_search_output_chars() -> int:
    return env_int("EW_MAX_SEARCH_OUTPUT_CHARS", 200_000, minimum=256)


def get_max_regex_pattern_chars() -> int:
    return env_int("EW_MAX_REGEX_PATTERN_CHARS", 5_000, minimum=1)


def get_max_regex_scan_chars() -> int:
    return env_int("EW_MAX_REGEX_SCAN_CHARS", 2_000_000, minimum=1)


def get_max_doc_bytes_per_operation() -> int:
    return env_int("EW_MAX_DOC_BYTES_PER_OPERATION", 50_000_000, minimum=1_024)


def get_max_undo_bytes_total() -> int:
    return env_int("EW_MAX_UNDO_BYTES_TOTAL", 200_000_000, minimum=1_024)


def get_regex_timeout_ms() -> int:
    return env_int("EW_REGEX_TIMEOUT_MS", 0, minimum=0)


def get_long_session_op_limit() -> int:
    return env_int("EW_LONG_SESSION_OP_LIMIT", 2000, minimum=1)


def is_regex_timeout_supported() -> bool:
    """Report whether runtime regex timeout handling is available."""
    try:
        import regex  # noqa: F401
    except Exception:
        return False
    return True


def check_doc_size_for_operation(path: str, operation_name: str) -> Tuple[bool, str]:
    """Validate document size before expensive operations."""
    if not os.path.exists(path):
        return True, ""

    limit = get_max_doc_bytes_per_operation()
    try:
        size = os.path.getsize(path)
    except OSError as exc:
        return False, f"Failed to inspect document size for '{path}': {exc}"

    if size <= limit:
        return True, ""

    return (
        False,
        f"[{DOC_TOO_LARGE}] Operation '{operation_name}' refused: "
        f"document size {size} bytes exceeds EW_MAX_DOC_BYTES_PER_OPERATION={limit}.",
    )
