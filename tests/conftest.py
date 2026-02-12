from __future__ import annotations

import pytest


@pytest.fixture(autouse=True)
def _isolate_session_manager():
    # Tools use a global singleton session manager. Keep tests isolated.
    from word_document_server.session_manager import get_session_manager
    from word_document_server.undo_manager import get_undo_manager

    mgr = get_session_manager()
    undo_mgr = get_undo_manager()
    mgr.close_all_documents()
    undo_mgr.clear_history()
    yield
    mgr.close_all_documents()
    undo_mgr.clear_history()
