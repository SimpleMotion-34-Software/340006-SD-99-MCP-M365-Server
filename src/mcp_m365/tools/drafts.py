"""Draft management tools for M365 MCP server."""

from typing import Any

from mcp.types import Tool

from ..graph import GraphClient

DRAFT_TOOLS = [
    Tool(
        name="m365_list_drafts",
        description="List all draft email messages.",
        inputSchema={
            "type": "object",
            "properties": {
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of drafts to return. Default: 25",
                    "default": 25,
                },
                "skip": {
                    "type": "integer",
                    "description": "Number of drafts to skip for pagination. Default: 0",
                    "default": 0,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_create_draft",
        description="Create a new draft email message. The draft can be edited later or sent using m365_send_draft.",
        inputSchema={
            "type": "object",
            "properties": {
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of recipient email addresses (optional for drafts)",
                },
                "subject": {
                    "type": "string",
                    "description": "Email subject line",
                    "default": "",
                },
                "body": {
                    "type": "string",
                    "description": "Email body content (HTML or plain text)",
                    "default": "",
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
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_update_draft",
        description="Update an existing draft email message. Only the fields you provide will be updated.",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The draft message ID to update",
                },
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Updated list of recipient email addresses",
                },
                "subject": {
                    "type": "string",
                    "description": "Updated email subject",
                },
                "body": {
                    "type": "string",
                    "description": "Updated email body content",
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
                    "description": "Updated list of CC recipients",
                },
                "bcc": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Updated list of BCC recipients",
                },
                "importance": {
                    "type": "string",
                    "enum": ["low", "normal", "high"],
                    "description": "Updated message importance",
                },
            },
            "required": ["message_id"],
        },
    ),
    Tool(
        name="m365_delete_draft",
        description="Delete a draft email message.",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The draft message ID to delete",
                },
            },
            "required": ["message_id"],
        },
    ),
    Tool(
        name="m365_send_draft",
        description="Send an existing draft email message.",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The draft message ID to send",
                },
            },
            "required": ["message_id"],
        },
    ),
]


def _format_draft_summary(msg: dict[str, Any]) -> dict[str, Any]:
    """Format a draft message into a summary object."""
    return {
        "id": msg.get("id"),
        "subject": msg.get("subject", "(No subject)"),
        "to": [r.get("emailAddress", {}).get("address") for r in msg.get("toRecipients", [])],
        "created": msg.get("createdDateTime"),
        "modified": msg.get("lastModifiedDateTime"),
        "has_attachments": msg.get("hasAttachments", False),
        "preview": msg.get("bodyPreview", "")[:200],
    }


async def handle_draft_tool(
    name: str, arguments: dict[str, Any], client: GraphClient
) -> dict[str, Any]:
    """Handle draft management tool calls.

    Args:
        name: Tool name
        arguments: Tool arguments
        client: Graph API client

    Returns:
        Tool result
    """
    if name == "m365_list_drafts":
        limit = min(arguments.get("limit", 25), 1000)
        skip = arguments.get("skip", 0)

        result = await client.list_drafts(top=limit, skip=skip)

        drafts = [_format_draft_summary(msg) for msg in result.get("value", [])]

        return {
            "drafts": drafts,
            "count": len(drafts),
            "has_more": "@odata.nextLink" in result,
        }

    elif name == "m365_create_draft":
        to = arguments.get("to")
        subject = arguments.get("subject", "")
        body = arguments.get("body", "")
        body_type = arguments.get("body_type", "HTML")
        cc = arguments.get("cc")
        bcc = arguments.get("bcc")
        importance = arguments.get("importance", "normal")

        draft = await client.create_draft(
            to=to,
            subject=subject,
            body=body,
            body_type=body_type,
            cc=cc,
            bcc=bcc,
            importance=importance,
        )

        return {
            "success": True,
            "message": "Draft created successfully",
            "draft_id": draft.get("id"),
            "subject": draft.get("subject"),
        }

    elif name == "m365_update_draft":
        message_id = arguments.get("message_id")
        if not message_id:
            return {"error": "message_id is required"}

        # Build update kwargs - only include fields that were provided
        update_kwargs: dict[str, Any] = {"message_id": message_id}

        if "to" in arguments:
            update_kwargs["to"] = arguments["to"]
        if "subject" in arguments:
            update_kwargs["subject"] = arguments["subject"]
        if "body" in arguments:
            update_kwargs["body"] = arguments["body"]
            update_kwargs["body_type"] = arguments.get("body_type", "HTML")
        if "cc" in arguments:
            update_kwargs["cc"] = arguments["cc"]
        if "bcc" in arguments:
            update_kwargs["bcc"] = arguments["bcc"]
        if "importance" in arguments:
            update_kwargs["importance"] = arguments["importance"]

        draft = await client.update_draft(**update_kwargs)

        return {
            "success": True,
            "message": "Draft updated successfully",
            "draft_id": draft.get("id"),
            "subject": draft.get("subject"),
        }

    elif name == "m365_delete_draft":
        message_id = arguments.get("message_id")
        if not message_id:
            return {"error": "message_id is required"}

        result = await client.delete_draft(message_id)

        return {
            **result,
            "draft_id": message_id,
        }

    elif name == "m365_send_draft":
        message_id = arguments.get("message_id")
        if not message_id:
            return {"error": "message_id is required"}

        result = await client.send_draft(message_id)

        return {
            **result,
            "draft_id": message_id,
        }

    return {"error": f"Unknown draft tool: {name}"}
