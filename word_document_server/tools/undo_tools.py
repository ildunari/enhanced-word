"""Undo / Redo tools for Enhanced Word server.

This module registers a **single** public tool, ``session_undo``, that can:

• undo one or more steps
• redo
• list history sizes
• clear history

Keeping everything behind a single entry point helps us stay below the global
25-tool limit requested by the project guidelines.
"""

from typing import Optional

from word_document_server.utils.session_utils import resolve_document_path
from word_document_server.undo_manager import get_undo_manager


def session_undo(
    action: str = "undo",
    steps: int = 1,
    document_id: Optional[str] = None,
    filename: Optional[str] = None,
) -> str:
    """Perform undo-related actions.

    Parameters
    ----------
    action : str
        One of:
        * ``"undo"``  – revert *steps* modifications.
        * ``"redo"``  – re-apply *steps* previously undone modifications.
        * ``"list"``  – report how many undo & redo snapshots exist.
        * ``"clear"`` – forget history (for given document or all).
    steps : int, optional
        How many steps to undo / redo (ignored for *list* & *clear*).
    document_id, filename : str, optional
        Identify the target document.  If both are provided, *document_id*
        wins.  If neither is provided the active document (from
        ``DocumentSessionManager``) will be used automatically by
        ``resolve_document_path``.
    """

    # Validate action
    valid = {"undo", "redo", "list", "clear"}
    if action not in valid:
        return f"Invalid action: {action}. Must be one of: {', '.join(sorted(valid))}"

    # Reject negative step counts early
    if steps < 0:
        return "Error: steps must be non-negative"

    # For undo/redo we require a concrete target path
    if action in {"undo", "redo"}:
        # Need a concrete target path
        file_path, err = resolve_document_path(document_id, filename)
        if err:
            return err
    else:
        # list/clear may omit path
        if document_id or filename:
            file_path, err = resolve_document_path(document_id, filename)
            if err:
                return err
        else:
            file_path = None  # Means "all documents"

    undo_mgr = get_undo_manager()

    if action == "undo":
        return undo_mgr.undo(file_path, steps)
    if action == "redo":
        return undo_mgr.redo(file_path, steps)
    if action == "list":
        if file_path is None:
            # Aggregate list for all documents
            summaries = [undo_mgr.list_history(p) for p in list(undo_mgr._stacks.keys())]  # pylint: disable=protected-access
            return "\n".join(summaries) if summaries else "No undo history tracked yet"
        return undo_mgr.list_history(file_path)
    if action == "clear":
        return undo_mgr.clear_history(file_path)


# Exported tools list (for auto-registration systems)
__all__ = [
    "session_undo",
] 