"""Office Word MCP Server package entry point."""

def run_server(*args, **kwargs):
    """Lazy wrapper for starting the MCP server."""
    from word_document_server.main import run_server as _run_server
    return _run_server(*args, **kwargs)

__all__ = ["run_server"]
