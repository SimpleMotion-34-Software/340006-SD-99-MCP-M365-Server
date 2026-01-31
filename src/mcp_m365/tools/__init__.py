"""MCP Tools for Microsoft 365 operations."""

from typing import Any, Callable, Dict, List, Optional

from mcp.types import Tool

from .auth import AUTH_TOOLS, AUTH_HANDLERS
from .messages import MESSAGE_TOOLS, MESSAGE_HANDLERS
from .send import SEND_TOOLS, SEND_HANDLERS
from .drafts import DRAFT_TOOLS, DRAFT_HANDLERS
from .folders import FOLDER_TOOLS, FOLDER_HANDLERS
from .contacts import CONTACT_TOOLS, CONTACT_HANDLERS
from .planner import PLANNER_TOOLS, PLANNER_HANDLERS


# Combine all tools
ALL_TOOLS: List[Tool] = (
    AUTH_TOOLS +
    MESSAGE_TOOLS +
    SEND_TOOLS +
    DRAFT_TOOLS +
    FOLDER_TOOLS +
    CONTACT_TOOLS +
    PLANNER_TOOLS
)

# Combine all handlers
ALL_HANDLERS: Dict[str, Callable] = {
    **AUTH_HANDLERS,
    **MESSAGE_HANDLERS,
    **SEND_HANDLERS,
    **DRAFT_HANDLERS,
    **FOLDER_HANDLERS,
    **CONTACT_HANDLERS,
    **PLANNER_HANDLERS,
}


def get_tool_handler(name: str) -> Optional[Callable]:
    """Get the handler function for a tool.

    Args:
        name: The tool name

    Returns:
        The handler function if found, None otherwise.
    """
    return ALL_HANDLERS.get(name)


__all__ = [
    "ALL_TOOLS",
    "ALL_HANDLERS",
    "get_tool_handler",
]
