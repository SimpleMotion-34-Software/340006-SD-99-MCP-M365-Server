"""Draft management tools for M365 MCP Server."""

from typing import Any, Dict, List

from mcp.types import Tool

from ..auth import M365OAuth
from ..graph import GraphClient


DRAFT_TOOLS: List[Tool] = [
    Tool(
        name="m365_list_drafts",
        description="List draft email messages",
        inputSchema={
            "type": "object",
            "properties": {
                "top": {
                    "type": "integer",
                    "description": "Number of drafts to return",
                    "default": 25,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_create_draft",
        description="Create a new draft email message",
        inputSchema={
            "type": "object",
            "properties": {
                "subject": {
                    "type": "string",
                    "description": "Email subject",
                },
                "body": {
                    "type": "string",
                    "description": "Email body content",
                },
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of recipient email addresses",
                },
                "cc": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of CC email addresses",
                },
                "is_html": {
                    "type": "boolean",
                    "description": "Whether body is HTML formatted",
                    "default": False,
                },
            },
            "required": ["subject", "body"],
        },
    ),
    Tool(
        name="m365_update_draft",
        description="Update an existing draft email message",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The draft message ID",
                },
                "subject": {
                    "type": "string",
                    "description": "New subject (optional)",
                },
                "body": {
                    "type": "string",
                    "description": "New body content (optional)",
                },
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "New recipient list (optional)",
                },
                "is_html": {
                    "type": "boolean",
                    "description": "Whether body is HTML formatted",
                    "default": False,
                },
            },
            "required": ["message_id"],
        },
    ),
    Tool(
        name="m365_delete_draft",
        description="Delete a draft email message",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The draft message ID",
                },
            },
            "required": ["message_id"],
        },
    ),
    Tool(
        name="m365_send_draft",
        description="Send a draft email message",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The draft message ID",
                },
            },
            "required": ["message_id"],
        },
    ),
]


def _format_draft(msg: Dict[str, Any]) -> Dict[str, Any]:
    """Format a draft message for display."""
    to_recipients = []
    for r in msg.get("toRecipients", []):
        if r.get("emailAddress"):
            to_recipients.append(r["emailAddress"].get("address", ""))

    return {
        "id": msg.get("id"),
        "subject": msg.get("subject"),
        "to": to_recipients,
        "created": msg.get("createdDateTime"),
        "modified": msg.get("lastModifiedDateTime"),
        "preview": msg.get("bodyPreview", "")[:200],
    }


async def handle_list_drafts(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_list_drafts tool call."""
    top = min(arguments.get("top", 25), 50)

    drafts = await client.list_drafts(top=top)

    return {
        "count": len(drafts),
        "drafts": [_format_draft(d) for d in drafts],
    }


async def handle_create_draft(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_create_draft tool call."""
    subject = arguments["subject"]
    body = arguments["body"]
    to_recipients = arguments.get("to")
    cc_recipients = arguments.get("cc")
    is_html = arguments.get("is_html", False)

    draft = await client.create_draft(
        subject=subject,
        body=body,
        to_recipients=to_recipients,
        cc_recipients=cc_recipients,
        is_html=is_html,
    )

    return {
        "status": "created",
        "draft": _format_draft(draft),
    }


async def handle_update_draft(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_update_draft tool call."""
    message_id = arguments["message_id"]
    subject = arguments.get("subject")
    body = arguments.get("body")
    to_recipients = arguments.get("to")
    is_html = arguments.get("is_html", False)

    draft = await client.update_draft(
        message_id=message_id,
        subject=subject,
        body=body,
        to_recipients=to_recipients,
        is_html=is_html,
    )

    return {
        "status": "updated",
        "draft": _format_draft(draft),
    }


async def handle_delete_draft(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_delete_draft tool call."""
    message_id = arguments["message_id"]

    await client.delete_draft(message_id)

    return {
        "status": "deleted",
        "message_id": message_id,
    }


async def handle_send_draft(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_send_draft tool call."""
    message_id = arguments["message_id"]

    await client.send_draft(message_id)

    return {
        "status": "sent",
        "message_id": message_id,
    }


DRAFT_HANDLERS = {
    "m365_list_drafts": handle_list_drafts,
    "m365_create_draft": handle_create_draft,
    "m365_update_draft": handle_update_draft,
    "m365_delete_draft": handle_delete_draft,
    "m365_send_draft": handle_send_draft,
}
