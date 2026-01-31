"""Send/reply/forward tools for M365 MCP Server."""

from typing import Any, Dict, List

from mcp.types import Tool

from ..auth import M365OAuth
from ..graph import GraphClient


SEND_TOOLS: List[Tool] = [
    Tool(
        name="m365_send_message",
        description="Send a new email message",
        inputSchema={
            "type": "object",
            "properties": {
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of recipient email addresses",
                },
                "subject": {
                    "type": "string",
                    "description": "Email subject",
                },
                "body": {
                    "type": "string",
                    "description": "Email body content",
                },
                "cc": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of CC email addresses",
                },
                "bcc": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of BCC email addresses",
                },
                "is_html": {
                    "type": "boolean",
                    "description": "Whether body is HTML formatted",
                    "default": False,
                },
            },
            "required": ["to", "subject", "body"],
        },
    ),
    Tool(
        name="m365_reply",
        description="Reply to an email message",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The message ID to reply to",
                },
                "comment": {
                    "type": "string",
                    "description": "Reply text",
                },
                "reply_all": {
                    "type": "boolean",
                    "description": "Reply to all recipients",
                    "default": False,
                },
            },
            "required": ["message_id", "comment"],
        },
    ),
    Tool(
        name="m365_forward",
        description="Forward an email message",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The message ID to forward",
                },
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of recipient email addresses",
                },
                "comment": {
                    "type": "string",
                    "description": "Optional comment to include",
                },
            },
            "required": ["message_id", "to"],
        },
    ),
]


async def handle_send_message(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_send_message tool call."""
    to_recipients = arguments["to"]
    subject = arguments["subject"]
    body = arguments["body"]
    cc_recipients = arguments.get("cc")
    bcc_recipients = arguments.get("bcc")
    is_html = arguments.get("is_html", False)

    await client.send_message(
        subject=subject,
        body=body,
        to_recipients=to_recipients,
        cc_recipients=cc_recipients,
        bcc_recipients=bcc_recipients,
        is_html=is_html,
    )

    return {
        "status": "sent",
        "to": to_recipients,
        "subject": subject,
    }


async def handle_reply(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_reply tool call."""
    message_id = arguments["message_id"]
    comment = arguments["comment"]
    reply_all = arguments.get("reply_all", False)

    await client.reply_to_message(
        message_id=message_id,
        comment=comment,
        reply_all=reply_all,
    )

    return {
        "status": "replied",
        "message_id": message_id,
        "reply_all": reply_all,
    }


async def handle_forward(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_forward tool call."""
    message_id = arguments["message_id"]
    to_recipients = arguments["to"]
    comment = arguments.get("comment")

    await client.forward_message(
        message_id=message_id,
        to_recipients=to_recipients,
        comment=comment,
    )

    return {
        "status": "forwarded",
        "message_id": message_id,
        "to": to_recipients,
    }


SEND_HANDLERS = {
    "m365_send_message": handle_send_message,
    "m365_reply": handle_reply,
    "m365_forward": handle_forward,
}
