"""
Shared response helpers for MCP tools.
"""
import os
from typing import Optional


def sanitize_presentation_name(presentation_file_name: Optional[str]) -> Optional[str]:
    """
    Return only the filename component for presentation identifiers to avoid leaking paths.

    Args:
        presentation_file_name: Raw user-supplied presentation identifier.

    Returns:
        The basename of the identifier with trailing separators removed. If the provided
        value is empty or None, it is returned unchanged.
    """
    if not presentation_file_name:
        return presentation_file_name

    stripped = presentation_file_name.strip()
    if not stripped:
        return stripped

    stripped = stripped.rstrip("\\/")
    if not stripped:
        return stripped

    basename = os.path.basename(stripped)
    return basename or stripped

