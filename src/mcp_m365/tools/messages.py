"""Message reading tools for M365 MCP server."""

import base64
from typing import Any

from mcp.types import Tool

from ..graph import GraphClient

MESSAGE_TOOLS = [
    Tool(
        name="m365_list_messages",
        description="List email messages from a folder. Returns message summaries with subject, sender, date, and preview.",
        inputSchema={
            "type": "object",
            "properties": {
                "folder": {
                    "type": "string",
                    "description": "Folder to list messages from: inbox, sentItems, drafts, deletedItems, or a folder ID. Default: inbox",
                    "default": "inbox",
                },
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of messages to return (1-1000). Default: 25",
                    "default": 25,
                },
                "skip": {
                    "type": "integer",
                    "description": "Number of messages to skip for pagination. Default: 0",
                    "default": 0,
                },
                "unread_only": {
                    "type": "boolean",
                    "description": "Only return unread messages. Default: false",
                    "default": False,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_search_messages",
        description="Search for emails using keywords. Searches in subject, body, and other fields. Use KQL syntax for advanced queries (e.g., 'from:john subject:meeting').",
        inputSchema={
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search query. Use KQL syntax: 'keyword', 'from:email', 'subject:text', 'received:today', etc.",
                },
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of results to return. Default: 25",
                    "default": 25,
                },
                "folder": {
                    "type": "string",
                    "description": "Optional folder to search in. If not specified, searches all folders.",
                },
            },
            "required": ["query"],
        },
    ),
    Tool(
        name="m365_get_message",
        description="Get the full content of a specific email message by its ID, including the complete body.",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The message ID to retrieve",
                },
                "mark_as_read": {
                    "type": "boolean",
                    "description": "Mark the message as read after retrieving. Default: false",
                    "default": False,
                },
            },
            "required": ["message_id"],
        },
    ),
    Tool(
        name="m365_get_thread",
        description="Get all messages in an email conversation thread.",
        inputSchema={
            "type": "object",
            "properties": {
                "conversation_id": {
                    "type": "string",
                    "description": "The conversation ID (from a message's conversationId field)",
                },
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of messages to return. Default: 50",
                    "default": 50,
                },
            },
            "required": ["conversation_id"],
        },
    ),
    Tool(
        name="m365_get_attachment",
        description="Get an attachment from an email message. Returns the attachment content as base64-encoded data.",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The message ID containing the attachment",
                },
                "attachment_id": {
                    "type": "string",
                    "description": "The attachment ID. Use m365_get_message first to see available attachments.",
                },
            },
            "required": ["message_id", "attachment_id"],
        },
    ),
]


def _format_message_summary(msg: dict[str, Any]) -> dict[str, Any]:
    """Format a message into a summary object."""
    sender = msg.get("from", {}).get("emailAddress", {})
    return {
        "id": msg.get("id"),
        "subject": msg.get("subject", "(No subject)"),
        "from": sender.get("address", "Unknown"),
        "from_name": sender.get("name"),
        "to": [r.get("emailAddress", {}).get("address") for r in msg.get("toRecipients", [])],
        "received": msg.get("receivedDateTime"),
        "is_read": msg.get("isRead", False),
        "has_attachments": msg.get("hasAttachments", False),
        "importance": msg.get("importance", "normal"),
        "preview": msg.get("bodyPreview", "")[:200],
    }


def _format_message_detail(msg: dict[str, Any]) -> dict[str, Any]:
    """Format a message with full details."""
    sender = msg.get("from", {}).get("emailAddress", {})
    body = msg.get("body", {})

    result = {
        "id": msg.get("id"),
        "conversation_id": msg.get("conversationId"),
        "subject": msg.get("subject", "(No subject)"),
        "from": sender.get("address", "Unknown"),
        "from_name": sender.get("name"),
        "to": [
            {"email": r.get("emailAddress", {}).get("address"), "name": r.get("emailAddress", {}).get("name")}
            for r in msg.get("toRecipients", [])
        ],
        "cc": [
            {"email": r.get("emailAddress", {}).get("address"), "name": r.get("emailAddress", {}).get("name")}
            for r in msg.get("ccRecipients", [])
        ],
        "received": msg.get("receivedDateTime"),
        "sent": msg.get("sentDateTime"),
        "is_read": msg.get("isRead", False),
        "has_attachments": msg.get("hasAttachments", False),
        "importance": msg.get("importance", "normal"),
        "body_type": body.get("contentType", "text"),
        "body": body.get("content", ""),
    }

    return result


async def handle_message_tool(
    name: str, arguments: dict[str, Any], client: GraphClient
) -> dict[str, Any]:
    """Handle message reading tool calls.

    Args:
        name: Tool name
        arguments: Tool arguments
        client: Graph API client

    Returns:
        Tool result
    """
    if name == "m365_list_messages":
        folder = arguments.get("folder", "inbox")
        limit = min(arguments.get("limit", 25), 1000)
        skip = arguments.get("skip", 0)
        unread_only = arguments.get("unread_only", False)

        filter_query = "isRead eq false" if unread_only else None

        result = await client.list_messages(
            folder=folder,
            top=limit,
            skip=skip,
            filter_query=filter_query,
        )

        messages = [_format_message_summary(msg) for msg in result.get("value", [])]

        return {
            "messages": messages,
            "count": len(messages),
            "has_more": "@odata.nextLink" in result,
            "folder": folder,
        }

    elif name == "m365_search_messages":
        query = arguments.get("query", "")
        limit = min(arguments.get("limit", 25), 1000)
        folder = arguments.get("folder")

        if not query:
            return {"error": "Search query is required"}

        result = await client.search_messages(
            query=query,
            top=limit,
            folder=folder,
        )

        messages = [_format_message_summary(msg) for msg in result.get("value", [])]

        return {
            "messages": messages,
            "count": len(messages),
            "query": query,
        }

    elif name == "m365_get_message":
        message_id = arguments.get("message_id")
        mark_as_read = arguments.get("mark_as_read", False)

        if not message_id:
            return {"error": "message_id is required"}

        msg = await client.get_message(message_id)

        if mark_as_read and not msg.get("isRead"):
            await client.mark_as_read(message_id, True)

        # Get attachments if present
        attachments = []
        if msg.get("hasAttachments"):
            attach_result = await client.get_attachments(message_id)
            for att in attach_result.get("value", []):
                attachments.append({
                    "id": att.get("id"),
                    "name": att.get("name"),
                    "content_type": att.get("contentType"),
                    "size": att.get("size"),
                })

        result = _format_message_detail(msg)
        result["attachments"] = attachments

        return result

    elif name == "m365_get_thread":
        conversation_id = arguments.get("conversation_id")
        limit = arguments.get("limit", 50)

        if not conversation_id:
            return {"error": "conversation_id is required"}

        result = await client.get_thread(conversation_id, top=limit)

        messages = []
        for msg in result.get("value", []):
            sender = msg.get("from", {}).get("emailAddress", {})
            body = msg.get("body", {})
            messages.append({
                "id": msg.get("id"),
                "subject": msg.get("subject", "(No subject)"),
                "from": sender.get("address"),
                "from_name": sender.get("name"),
                "received": msg.get("receivedDateTime"),
                "body_preview": msg.get("bodyPreview", "")[:500],
            })

        return {
            "messages": messages,
            "count": len(messages),
            "conversation_id": conversation_id,
        }

    elif name == "m365_get_attachment":
        message_id = arguments.get("message_id")
        attachment_id = arguments.get("attachment_id")

        if not message_id:
            return {"error": "message_id is required"}
        if not attachment_id:
            return {"error": "attachment_id is required"}

        attachment = await client.get_attachment_content(message_id, attachment_id)

        return {
            "id": attachment.get("id"),
            "name": attachment.get("name"),
            "content_type": attachment.get("contentType"),
            "size": attachment.get("size"),
            "content_bytes": attachment.get("contentBytes"),  # Base64 encoded
        }

    return {"error": f"Unknown message tool: {name}"}
