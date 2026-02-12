from __future__ import annotations

"""Undo / Redo management for Enhanced Word server.

This module keeps an in-memory history (list of binary snapshots) for every
.docx file path that is being modified by the tools.  The history is captured
just before a file is written – see utils.file_utils.check_file_writeable,
which calls ``snapshot``.

A single public tool (see tools/undo_tools.py) exposes user-facing operations
without increasing the overall MCP tool count dramatically.
"""

# Standard Library
import os
from pathlib import Path
from threading import Lock
from typing import Dict, List, Optional
from word_document_server.utils.limits import (
    UNDO_BUDGET_EXCEEDED,
    get_max_undo_bytes_total,
)


class _Snapshot:
    """Single history snapshot with creation order for deterministic eviction."""

    __slots__ = ("seq", "data")

    def __init__(self, seq: int, data: bytes) -> None:
        self.seq = seq
        self.data = data


class _HistoryStacks:
    """Container holding undo/redo stacks for one file path."""

    __slots__ = ("undo", "redo")

    def __init__(self) -> None:
        self.undo: List[_Snapshot] = []
        self.redo: List[_Snapshot] = []


class UndoManager:
    """Keeps binary snapshots of files to support undo / redo."""

    def __init__(self, max_depth: int = 20):
        # path -> _HistoryStacks
        self._stacks: Dict[str, _HistoryStacks] = {}
        self._max_depth = max_depth
        self._lock = Lock()
        self._next_seq = 0
        self._budget_evictions = 0

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _get_stacks(self, path: str) -> _HistoryStacks:
        if path not in self._stacks:
            self._stacks[path] = _HistoryStacks()
        return self._stacks[path]

    def _new_snapshot(self, data: bytes) -> _Snapshot:
        snapshot = _Snapshot(self._next_seq, data)
        self._next_seq += 1
        return snapshot

    def _total_snapshot_bytes(self) -> int:
        total = 0
        for stacks in self._stacks.values():
            total += sum(len(s.data) for s in stacks.undo)
            total += sum(len(s.data) for s in stacks.redo)
        return total

    def _enforce_total_budget(self) -> None:
        limit = get_max_undo_bytes_total()
        while self._total_snapshot_bytes() > limit:
            oldest_path = None
            oldest_stack = None
            oldest_seq = None

            for path, stacks in self._stacks.items():
                if stacks.undo:
                    seq = stacks.undo[0].seq
                    if oldest_seq is None or seq < oldest_seq:
                        oldest_path = path
                        oldest_stack = "undo"
                        oldest_seq = seq
                if stacks.redo:
                    seq = stacks.redo[0].seq
                    if oldest_seq is None or seq < oldest_seq:
                        oldest_path = path
                        oldest_stack = "redo"
                        oldest_seq = seq

            if oldest_path is None or oldest_stack is None:
                break

            stacks = self._stacks[oldest_path]
            if oldest_stack == "undo":
                stacks.undo.pop(0)
            else:
                stacks.redo.pop(0)
            self._budget_evictions += 1

    # ------------------------------------------------------------------
    # Public history operations (used by tools & tests)
    # ------------------------------------------------------------------

    def snapshot(self, path: str) -> None:
        """Capture current file bytes before it gets modified.

        This should be called as *early* as possible, ideally right before
        a tool opens the document for editing.  New snapshots clear the redo
        stack (standard application behaviour).
        """
        normalized = str(Path(path).resolve())
        if not os.path.isfile(normalized):
            # Nothing to snapshot
            return

        with self._lock:
            stacks = self._get_stacks(normalized)
            try:
                with open(normalized, "rb") as f:
                    data = f.read()
            except Exception:
                return  # Skip snapshot on read error

            stacks.undo.append(self._new_snapshot(data))
            if len(stacks.undo) > self._max_depth:
                stacks.undo.pop(0)
            stacks.redo.clear()
            self._enforce_total_budget()

    def undo(self, path: str, steps: int = 1) -> str:
        """Restore *steps* previous states for *path*.

        Returns a human-readable status string.
        """
        if steps < 1:
            return "Error: steps must be >= 1"

        normalized = str(Path(path).resolve())
        with self._lock:
            stacks = self._stacks.get(normalized)
            if not stacks or not stacks.undo:
                return f"No undo history for {normalized}"

            requested = steps
            processed = 0
            for _ in range(min(steps, len(stacks.undo))):
                previous = stacks.undo.pop()
                try:
                    with open(normalized, "rb") as f:
                        current_bytes = f.read()
                    stacks.redo.append(self._new_snapshot(current_bytes))
                except Exception:
                    pass

                try:
                    with open(normalized, "wb") as f:
                        f.write(previous.data)
                except Exception as e:
                    stacks.undo.append(previous)
                    return f"Failed to restore snapshot: {e}"

                processed += 1
                self._enforce_total_budget()

            return f"Undo successful (requested {requested}, restored {processed}) for {normalized}"

    def redo(self, path: str, steps: int = 1) -> str:
        if steps < 1:
            return "Error: steps must be >= 1"

        normalized = str(Path(path).resolve())
        with self._lock:
            stacks = self._stacks.get(normalized)
            if not stacks or not stacks.redo:
                return f"No redo history for {normalized}"

            requested = steps
            processed = 0
            for _ in range(min(steps, len(stacks.redo))):
                next_state = stacks.redo.pop()
                try:
                    with open(normalized, "rb") as f:
                        current_bytes = f.read()
                    stacks.undo.append(self._new_snapshot(current_bytes))
                    if len(stacks.undo) > self._max_depth:
                        stacks.undo.pop(0)
                except Exception:
                    pass

                try:
                    with open(normalized, "wb") as f:
                        f.write(next_state.data)
                except Exception as e:
                    stacks.redo.append(next_state)
                    return f"Failed to re-apply snapshot: {e}"

                processed += 1
                self._enforce_total_budget()

            return f"Redo successful (requested {requested}, applied {processed}) for {normalized}"

    def list_history(self, path: str) -> str:
        normalized = str(Path(path).resolve())
        with self._lock:
            stacks = self._stacks.get(normalized)
            undo_len = len(stacks.undo) if stacks else 0
            redo_len = len(stacks.redo) if stacks else 0
            total_bytes = self._total_snapshot_bytes()
            evictions = self._budget_evictions
        suffix = f", total_snapshot_bytes={total_bytes}, budget_evictions={evictions}"
        if evictions > 0:
            suffix += f" [{UNDO_BUDGET_EXCEEDED}]"
        return f"History for {normalized}: undo={undo_len}, redo={redo_len}{suffix}"

    def clear_history(self, path: Optional[str] = None) -> str:
        with self._lock:
            if path:
                normalized = str(Path(path).resolve())
                self._stacks.pop(normalized, None)
                return f"Cleared history for {normalized}"
            # Clear all
            self._stacks.clear()
            return "Cleared history for all documents"


# ----------------------------------------------------------------------
# Global accessor
# ----------------------------------------------------------------------

_undo_manager = UndoManager()

def get_undo_manager() -> UndoManager:  # noqa: D401 (simple function)
    """Return the global singleton UndoManager."""
    return _undo_manager 
