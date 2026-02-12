from __future__ import annotations

import word_document_server.tools.session_tools as session_tools


def test_session_tools___all___symbols_exist():
    for name in session_tools.__all__:
        assert hasattr(session_tools, name), f"missing export: {name}"

