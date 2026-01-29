"""Teams chat tools for M365 MCP server."""

from typing import Any

from mcp.types import Tool

from ..graph import GraphClient

CHAT_TOOLS = [
    Tool(
        name="m365_list_chats",
        description="List all Teams chats (1:1, group, meeting). Returns chat summaries with type, topic, and last message preview.",
        inputSchema={
            "type": "object",
            "properties": {
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of chats to return (1-50). Default: 50",
                    "default": 50,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_get_chat",
        description="Get details of a specific Teams chat by its ID, including members and last message.",
        inputSchema={
            "type": "object",
            "properties": {
                "chat_id": {
                    "type": "string",
                    "description": "The chat ID to retrieve",
                },
            },
            "required": ["chat_id"],
        },
    ),
    Tool(
        name="m365_get_chat_messages",
        description="Get messages from a specific Teams chat. Returns messages with sender, content, and timestamp.",
        inputSchema={
            "type": "object",
            "properties": {
                "chat_id": {
                    "type": "string",
                    "description": "The chat ID to get messages from",
                },
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of messages to return (1-50). Default: 50",
                    "default": 50,
                },
            },
            "required": ["chat_id"],
        },
    ),
    Tool(
        name="m365_get_chat_members",
        description="List all members of a Teams chat.",
        inputSchema={
            "type": "object",
            "properties": {
                "chat_id": {
                    "type": "string",
                    "description": "The chat ID to get members from",
                },
            },
            "required": ["chat_id"],
        },
    ),
    Tool(
        name="m365_send_chat_message",
        description="Send a message to a Teams chat.",
        inputSchema={
            "type": "object",
            "properties": {
                "chat_id": {
                    "type": "string",
                    "description": "The chat ID to send the message to",
                },
                "content": {
                    "type": "string",
                    "description": "Message content (HTML or plain text)",
                },
                "content_type": {
                    "type": "string",
                    "enum": ["html", "text"],
                    "description": "Content type. Default: html",
                    "default": "html",
                },
            },
            "required": ["chat_id", "content"],
        },
    ),
    Tool(
        name="m365_search_chat_messages",
        description="Search for messages across all Teams chats. Returns matching messages with sender, content, and chat context.",
        inputSchema={
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search query to find in chat messages",
                },
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of results to return (1-50). Default: 25",
                    "default": 25,
                },
            },
            "required": ["query"],
        },
    ),
]


def _format_chat_summary(chat: dict[str, Any]) -> dict[str, Any]:
    """Format a chat into a summary object."""
    last_message = chat.get("lastMessagePreview", {})
    last_msg_from = last_message.get("from", {}).get("user", {})

    return {
        "id": chat.get("id"),
        "chat_type": chat.get("chatType"),
        "topic": chat.get("topic") or _get_chat_display_name(chat),
        "created": chat.get("createdDateTime"),
        "last_updated": chat.get("lastUpdatedDateTime"),
        "last_message": {
            "preview": last_message.get("body", {}).get("content", "")[:200] if last_message else None,
            "from": last_msg_from.get("displayName") if last_msg_from else None,
            "created": last_message.get("createdDateTime") if last_message else None,
        } if last_message else None,
    }


def _get_chat_display_name(chat: dict[str, Any]) -> str:
    """Get a display name for a chat without a topic."""
    chat_type = chat.get("chatType", "")
    if chat_type == "oneOnOne":
        return "1:1 Chat"
    elif chat_type == "meeting":
        return "Meeting Chat"
    elif chat_type == "group":
        return "Group Chat"
    return "Chat"


def _format_chat_detail(chat: dict[str, Any]) -> dict[str, Any]:
    """Format a chat with full details."""
    last_message = chat.get("lastMessagePreview", {})
    members = chat.get("members", [])

    return {
        "id": chat.get("id"),
        "chat_type": chat.get("chatType"),
        "topic": chat.get("topic") or _get_chat_display_name(chat),
        "created": chat.get("createdDateTime"),
        "last_updated": chat.get("lastUpdatedDateTime"),
        "web_url": chat.get("webUrl"),
        "tenant_id": chat.get("tenantId"),
        "members": [
            {
                "id": m.get("id"),
                "display_name": m.get("displayName"),
                "email": m.get("email"),
                "roles": m.get("roles", []),
            }
            for m in members
        ],
        "last_message": {
            "id": last_message.get("id"),
            "content": last_message.get("body", {}).get("content"),
            "content_type": last_message.get("body", {}).get("contentType"),
            "from": last_message.get("from", {}).get("user", {}).get("displayName"),
            "created": last_message.get("createdDateTime"),
        } if last_message else None,
    }


def _format_chat_message(message: dict[str, Any]) -> dict[str, Any]:
    """Format a chat message."""
    sender = message.get("from", {})
    user = sender.get("user", {}) if sender else {}

    return {
        "id": message.get("id"),
        "content": message.get("body", {}).get("content"),
        "content_type": message.get("body", {}).get("contentType"),
        "from": user.get("displayName") if user else None,
        "from_email": user.get("email") if user else None,
        "created": message.get("createdDateTime"),
        "last_modified": message.get("lastModifiedDateTime"),
        "message_type": message.get("messageType"),
        "importance": message.get("importance"),
        "mentions": [
            {
                "id": m.get("id"),
                "mentioned": m.get("mentioned", {}).get("user", {}).get("displayName"),
            }
            for m in message.get("mentions", [])
        ],
        "attachments": [
            {
                "id": a.get("id"),
                "name": a.get("name"),
                "content_type": a.get("contentType"),
            }
            for a in message.get("attachments", [])
        ],
    }


def _format_chat_member(member: dict[str, Any]) -> dict[str, Any]:
    """Format a chat member."""
    return {
        "id": member.get("id"),
        "display_name": member.get("displayName"),
        "email": member.get("email"),
        "roles": member.get("roles", []),
        "visible_history_start": member.get("visibleHistoryStartDateTime"),
    }


async def handle_chat_tool(
    name: str, arguments: dict[str, Any], client: GraphClient
) -> dict[str, Any]:
    """Handle Teams chat tool calls.

    Args:
        name: Tool name
        arguments: Tool arguments
        client: Graph API client

    Returns:
        Tool result
    """
    if name == "m365_list_chats":
        limit = min(arguments.get("limit", 50), 50)

        result = await client.list_chats(top=limit)

        chats = [_format_chat_summary(chat) for chat in result.get("value", [])]

        return {
            "chats": chats,
            "count": len(chats),
            "has_more": "@odata.nextLink" in result,
        }

    elif name == "m365_get_chat":
        chat_id = arguments.get("chat_id")

        if not chat_id:
            return {"error": "chat_id is required"}

        chat = await client.get_chat(chat_id)

        return _format_chat_detail(chat)

    elif name == "m365_get_chat_messages":
        chat_id = arguments.get("chat_id")
        limit = min(arguments.get("limit", 50), 50)

        if not chat_id:
            return {"error": "chat_id is required"}

        result = await client.get_chat_messages(chat_id, top=limit)

        messages = [_format_chat_message(msg) for msg in result.get("value", [])]

        return {
            "messages": messages,
            "count": len(messages),
            "chat_id": chat_id,
            "has_more": "@odata.nextLink" in result,
        }

    elif name == "m365_get_chat_members":
        chat_id = arguments.get("chat_id")

        if not chat_id:
            return {"error": "chat_id is required"}

        result = await client.get_chat_members(chat_id)

        members = [_format_chat_member(m) for m in result.get("value", [])]

        return {
            "members": members,
            "count": len(members),
            "chat_id": chat_id,
        }

    elif name == "m365_send_chat_message":
        chat_id = arguments.get("chat_id")
        content = arguments.get("content")
        content_type = arguments.get("content_type", "html")

        if not chat_id:
            return {"error": "chat_id is required"}
        if not content:
            return {"error": "content is required"}

        result = await client.send_chat_message(
            chat_id=chat_id,
            content=content,
            content_type=content_type,
        )

        return {
            "success": True,
            "message_id": result.get("id"),
            "created": result.get("createdDateTime"),
        }

    elif name == "m365_search_chat_messages":
        query = arguments.get("query")
        limit = min(arguments.get("limit", 25), 50)

        if not query:
            return {"error": "query is required"}

        result = await client.search_chat_messages(query=query, top=limit)

        messages = [_format_chat_message(msg) for msg in result.get("value", [])]

        return {
            "messages": messages,
            "count": len(messages),
            "query": query,
            "has_more": "@odata.nextLink" in result,
        }

    return {"error": f"Unknown chat tool: {name}"}
