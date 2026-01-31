"""Folder management tools for M365 MCP Server."""

from typing import Any, Dict, List

from mcp.types import Tool

from ..auth import M365OAuth
from ..graph import GraphClient


FOLDER_TOOLS: List[Tool] = [
    Tool(
        name="m365_list_folders",
        description="List email folders in the mailbox",
        inputSchema={
            "type": "object",
            "properties": {
                "parent_folder_id": {
                    "type": "string",
                    "description": "Parent folder ID for listing child folders (optional)",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_create_folder",
        description="Create a new email folder",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Folder name",
                },
                "parent_folder_id": {
                    "type": "string",
                    "description": "Parent folder ID (optional, creates at root if not specified)",
                },
            },
            "required": ["name"],
        },
    ),
    Tool(
        name="m365_move_message",
        description="Move a message to a different folder",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The message ID to move",
                },
                "destination_folder_id": {
                    "type": "string",
                    "description": "Target folder ID",
                },
            },
            "required": ["message_id", "destination_folder_id"],
        },
    ),
    Tool(
        name="m365_delete_message",
        description="Delete an email message (moves to deleted items)",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The message ID to delete",
                },
            },
            "required": ["message_id"],
        },
    ),
]


def _format_folder(folder: Dict[str, Any]) -> Dict[str, Any]:
    """Format a folder for display."""
    return {
        "id": folder.get("id"),
        "display_name": folder.get("displayName"),
        "parent_folder_id": folder.get("parentFolderId"),
        "child_folder_count": folder.get("childFolderCount", 0),
        "unread_item_count": folder.get("unreadItemCount", 0),
        "total_item_count": folder.get("totalItemCount", 0),
    }


async def handle_list_folders(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_list_folders tool call."""
    parent_folder_id = arguments.get("parent_folder_id")

    folders = await client.list_folders(parent_folder_id=parent_folder_id)

    return {
        "count": len(folders),
        "folders": [_format_folder(f) for f in folders],
    }


async def handle_create_folder(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_create_folder tool call."""
    name = arguments["name"]
    parent_folder_id = arguments.get("parent_folder_id")

    folder = await client.create_folder(
        display_name=name,
        parent_folder_id=parent_folder_id,
    )

    return {
        "status": "created",
        "folder": _format_folder(folder),
    }


async def handle_move_message(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_move_message tool call."""
    message_id = arguments["message_id"]
    destination_folder_id = arguments["destination_folder_id"]

    message = await client.move_message(
        message_id=message_id,
        destination_folder_id=destination_folder_id,
    )

    return {
        "status": "moved",
        "message_id": message.get("id"),
        "destination_folder_id": destination_folder_id,
    }


async def handle_delete_message(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_delete_message tool call."""
    message_id = arguments["message_id"]

    await client.delete_message(message_id)

    return {
        "status": "deleted",
        "message_id": message_id,
    }


FOLDER_HANDLERS = {
    "m365_list_folders": handle_list_folders,
    "m365_create_folder": handle_create_folder,
    "m365_move_message": handle_move_message,
    "m365_delete_message": handle_delete_message,
}
