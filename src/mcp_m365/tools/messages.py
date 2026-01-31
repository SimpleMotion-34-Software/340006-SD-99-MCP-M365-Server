"""Message tools for M365 MCP Server."""

from typing import Any, Dict, List, Optional

from mcp.types import Tool

from ..auth import M365OAuth
from ..graph import GraphClient


MESSAGE_TOOLS: List[Tool] = [
    Tool(
        name="m365_list_messages",
        description="List email messages in a mailbox folder",
        inputSchema={
            "type": "object",
            "properties": {
                "folder": {
                    "type": "string",
                    "description": "Folder name (inbox, drafts, sentitems, deleteditems, junkemail) or folder ID",
                    "default": "inbox",
                },
                "top": {
                    "type": "integer",
                    "description": "Number of messages to return (max 50)",
                    "default": 25,
                },
                "skip": {
                    "type": "integer",
                    "description": "Number of messages to skip for pagination",
                    "default": 0,
                },
                "unread_only": {
                    "type": "boolean",
                    "description": "Only return unread messages",
                    "default": False,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_get_message",
        description="Get a specific email message with full body content",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The message ID",
                },
                "include_attachments": {
                    "type": "boolean",
                    "description": "Include attachment metadata",
                    "default": False,
                },
            },
            "required": ["message_id"],
        },
    ),
    Tool(
        name="m365_search_messages",
        description="Search email messages by keyword",
        inputSchema={
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search query (searches subject, body, and from)",
                },
                "top": {
                    "type": "integer",
                    "description": "Number of results to return",
                    "default": 25,
                },
            },
            "required": ["query"],
        },
    ),
    Tool(
        name="m365_get_thread",
        description="Get all messages in a conversation thread",
        inputSchema={
            "type": "object",
            "properties": {
                "conversation_id": {
                    "type": "string",
                    "description": "The conversation ID",
                },
                "top": {
                    "type": "integer",
                    "description": "Maximum messages to return",
                    "default": 25,
                },
            },
            "required": ["conversation_id"],
        },
    ),
    Tool(
        name="m365_get_attachment",
        description="Get an attachment from a message",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The message ID",
                },
                "attachment_id": {
                    "type": "string",
                    "description": "The attachment ID",
                },
            },
            "required": ["message_id", "attachment_id"],
        },
    ),
]


def _format_message_summary(msg: Dict[str, Any]) -> Dict[str, Any]:
    """Format a message for summary display."""
    from_email = ""
    if msg.get("from", {}).get("emailAddress"):
        from_email = msg["from"]["emailAddress"].get("address", "")

    return {
        "id": msg.get("id"),
        "subject": msg.get("subject"),
        "from": from_email,
        "received": msg.get("receivedDateTime"),
        "is_read": msg.get("isRead"),
        "has_attachments": msg.get("hasAttachments"),
        "preview": msg.get("bodyPreview", "")[:200],
    }


def _format_message_full(msg: Dict[str, Any]) -> Dict[str, Any]:
    """Format a message with full details."""
    from_email = ""
    if msg.get("from", {}).get("emailAddress"):
        from_email = msg["from"]["emailAddress"].get("address", "")

    to_recipients = []
    for r in msg.get("toRecipients", []):
        if r.get("emailAddress"):
            to_recipients.append(r["emailAddress"].get("address", ""))

    cc_recipients = []
    for r in msg.get("ccRecipients", []):
        if r.get("emailAddress"):
            cc_recipients.append(r["emailAddress"].get("address", ""))

    body = msg.get("body", {})

    return {
        "id": msg.get("id"),
        "conversation_id": msg.get("conversationId"),
        "subject": msg.get("subject"),
        "from": from_email,
        "to": to_recipients,
        "cc": cc_recipients,
        "received": msg.get("receivedDateTime"),
        "sent": msg.get("sentDateTime"),
        "is_read": msg.get("isRead"),
        "has_attachments": msg.get("hasAttachments"),
        "body_type": body.get("contentType", "text"),
        "body": body.get("content", ""),
        "importance": msg.get("importance"),
    }


async def handle_list_messages(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_list_messages tool call."""
    folder = arguments.get("folder", "inbox")
    top = min(arguments.get("top", 25), 50)
    skip = arguments.get("skip", 0)
    unread_only = arguments.get("unread_only", False)

    filter_query = None
    if unread_only:
        filter_query = "isRead eq false"

    messages = await client.list_messages(
        folder=folder,
        top=top,
        skip=skip,
        filter_query=filter_query,
    )

    return {
        "folder": folder,
        "count": len(messages),
        "messages": [_format_message_summary(m) for m in messages],
    }


async def handle_get_message(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_get_message tool call."""
    message_id = arguments["message_id"]
    include_attachments = arguments.get("include_attachments", False)

    message = await client.get_message(message_id)
    result = _format_message_full(message)

    if include_attachments and message.get("hasAttachments"):
        attachments = await client.list_attachments(message_id)
        result["attachments"] = [
            {
                "id": a.get("id"),
                "name": a.get("name"),
                "content_type": a.get("contentType"),
                "size": a.get("size"),
            }
            for a in attachments
        ]

    return result


async def handle_search_messages(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_search_messages tool call."""
    query = arguments["query"]
    top = min(arguments.get("top", 25), 50)

    messages = await client.search_messages(query=query, top=top)

    return {
        "query": query,
        "count": len(messages),
        "messages": [_format_message_summary(m) for m in messages],
    }


async def handle_get_thread(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_get_thread tool call."""
    conversation_id = arguments["conversation_id"]
    top = min(arguments.get("top", 25), 50)

    # Filter by conversation ID
    filter_query = f"conversationId eq '{conversation_id}'"

    messages = await client.list_messages(
        folder="inbox",
        top=top,
        filter_query=filter_query,
        order_by="receivedDateTime asc",
    )

    return {
        "conversation_id": conversation_id,
        "count": len(messages),
        "messages": [_format_message_summary(m) for m in messages],
    }


async def handle_get_attachment(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_get_attachment tool call."""
    message_id = arguments["message_id"]
    attachment_id = arguments["attachment_id"]

    attachment = await client.get_attachment(message_id, attachment_id)

    return {
        "id": attachment.get("id"),
        "name": attachment.get("name"),
        "content_type": attachment.get("contentType"),
        "size": attachment.get("size"),
        "content_bytes": attachment.get("contentBytes"),  # Base64 encoded
    }


MESSAGE_HANDLERS = {
    "m365_list_messages": handle_list_messages,
    "m365_get_message": handle_get_message,
    "m365_search_messages": handle_search_messages,
    "m365_get_thread": handle_get_thread,
    "m365_get_attachment": handle_get_attachment,
}
