"""Email sending tools for M365 MCP server."""

from typing import Any

from mcp.types import Tool

from ..graph import GraphClient

SEND_TOOLS = [
    Tool(
        name="m365_send_message",
        description="Compose and send a new email message.",
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
                    "description": "Email subject line",
                },
                "body": {
                    "type": "string",
                    "description": "Email body content (HTML or plain text)",
                },
                "body_type": {
                    "type": "string",
                    "enum": ["HTML", "Text"],
                    "description": "Body content type. Default: HTML",
                    "default": "HTML",
                },
                "cc": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of CC recipient email addresses",
                },
                "bcc": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of BCC recipient email addresses",
                },
                "importance": {
                    "type": "string",
                    "enum": ["low", "normal", "high"],
                    "description": "Message importance level. Default: normal",
                    "default": "normal",
                },
                "save_to_sent": {
                    "type": "boolean",
                    "description": "Save a copy to Sent Items folder. Default: true",
                    "default": True,
                },
            },
            "required": ["to", "subject", "body"],
        },
    ),
    Tool(
        name="m365_reply",
        description="Reply to an existing email message.",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The ID of the message to reply to",
                },
                "body": {
                    "type": "string",
                    "description": "Reply body content (HTML or plain text)",
                },
                "body_type": {
                    "type": "string",
                    "enum": ["HTML", "Text"],
                    "description": "Body content type. Default: HTML",
                    "default": "HTML",
                },
                "reply_all": {
                    "type": "boolean",
                    "description": "Reply to all recipients instead of just the sender. Default: false",
                    "default": False,
                },
            },
            "required": ["message_id", "body"],
        },
    ),
    Tool(
        name="m365_forward",
        description="Forward an email message to new recipients.",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The ID of the message to forward",
                },
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of recipient email addresses to forward to",
                },
                "comment": {
                    "type": "string",
                    "description": "Optional comment to add above the forwarded message",
                },
            },
            "required": ["message_id", "to"],
        },
    ),
]


async def handle_send_tool(
    name: str, arguments: dict[str, Any], client: GraphClient
) -> dict[str, Any]:
    """Handle email sending tool calls.

    Args:
        name: Tool name
        arguments: Tool arguments
        client: Graph API client

    Returns:
        Tool result
    """
    if name == "m365_send_message":
        to = arguments.get("to", [])
        subject = arguments.get("subject", "")
        body = arguments.get("body", "")
        body_type = arguments.get("body_type", "HTML")
        cc = arguments.get("cc")
        bcc = arguments.get("bcc")
        importance = arguments.get("importance", "normal")
        save_to_sent = arguments.get("save_to_sent", True)

        if not to:
            return {"error": "At least one recipient is required"}
        if not subject:
            return {"error": "Subject is required"}
        if not body:
            return {"error": "Body is required"}

        result = await client.send_message(
            to=to,
            subject=subject,
            body=body,
            body_type=body_type,
            cc=cc,
            bcc=bcc,
            importance=importance,
            save_to_sent=save_to_sent,
        )

        return {
            **result,
            "to": to,
            "subject": subject,
        }

    elif name == "m365_reply":
        message_id = arguments.get("message_id")
        body = arguments.get("body", "")
        body_type = arguments.get("body_type", "HTML")
        reply_all = arguments.get("reply_all", False)

        if not message_id:
            return {"error": "message_id is required"}
        if not body:
            return {"error": "Reply body is required"}

        result = await client.reply_to_message(
            message_id=message_id,
            body=body,
            body_type=body_type,
            reply_all=reply_all,
        )

        return {
            **result,
            "message_id": message_id,
            "reply_all": reply_all,
        }

    elif name == "m365_forward":
        message_id = arguments.get("message_id")
        to = arguments.get("to", [])
        comment = arguments.get("comment")

        if not message_id:
            return {"error": "message_id is required"}
        if not to:
            return {"error": "At least one recipient is required"}

        result = await client.forward_message(
            message_id=message_id,
            to=to,
            comment=comment,
        )

        return {
            **result,
            "message_id": message_id,
            "forwarded_to": to,
        }

    return {"error": f"Unknown send tool: {name}"}
