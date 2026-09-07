"""COM error formatting."""
import logging

_logger = logging.getLogger("outlook_desktop_mcp.errors")


def format_com_error(e: Exception) -> str:
    try:
        import pythoncom
        if isinstance(e, pythoncom.com_error):
            hr, msg, exc, arg = e.args
            details = exc[2] if exc else "No details"
            _logger.debug("COM error detail: %s", details)  # internal only
            # The outer HRESULT of a MAPI failure is almost always the generic
            # DISP_E_EXCEPTION. The inner scode is the one that says what actually
            # happened (for example 0x8004010F, message not found, when an EntryID
            # has gone stale). It is a status code, not item content, and callers
            # need it to tell a permanent condition from a transient one. Without
            # it every failure looks alike and gets retried forever.
            scode = exc[5] if exc and len(exc) > 5 and isinstance(exc[5], int) else None
            if scode is not None and scode != hr:
                return f"COM Error (0x{hr & 0xFFFFFFFF:08X}/0x{scode & 0xFFFFFFFF:08X}): {msg}"
            return f"COM Error (0x{hr & 0xFFFFFFFF:08X}): {msg}"
    except Exception:
        pass
    _logger.warning("Unexpected non-COM exception: %s: %s", type(e).__name__, e)
    # Name the exception type but not its message: the type is enough to tell a
    # permanent defect from a transient COM failure, and carries no item content.
    return f"An unexpected error occurred ({type(e).__name__})."
