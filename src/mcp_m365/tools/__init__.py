"""MCP tools for Microsoft Outlook and Teams integration."""

from .auth import AUTH_TOOLS, handle_auth_tool
from .chat import CHAT_TOOLS, handle_chat_tool
from .contacts import CONTACT_TOOLS, handle_contact_tool
from .drafts import DRAFT_TOOLS, handle_draft_tool
from .folders import FOLDER_TOOLS, handle_folder_tool
from .messages import MESSAGE_TOOLS, handle_message_tool
from .send import SEND_TOOLS, handle_send_tool

ALL_TOOLS = AUTH_TOOLS + MESSAGE_TOOLS + SEND_TOOLS + DRAFT_TOOLS + FOLDER_TOOLS + CHAT_TOOLS + CONTACT_TOOLS

__all__ = [
    "ALL_TOOLS",
    "AUTH_TOOLS",
    "MESSAGE_TOOLS",
    "SEND_TOOLS",
    "DRAFT_TOOLS",
    "FOLDER_TOOLS",
    "CHAT_TOOLS",
    "CONTACT_TOOLS",
    "handle_auth_tool",
    "handle_message_tool",
    "handle_send_tool",
    "handle_draft_tool",
    "handle_folder_tool",
    "handle_chat_tool",
    "handle_contact_tool",
]
