"""Folder management tools for M365 MCP server."""

from typing import Any

from mcp.types import Tool

from ..graph import GraphClient

FOLDER_TOOLS = [
    Tool(
        name="m365_list_folders",
        description="List all mail folders in the mailbox.",
        inputSchema={
            "type": "object",
            "properties": {
                "include_children": {
                    "type": "boolean",
                    "description": "Include nested child folders. Default: false",
                    "default": False,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_create_folder",
        description="Create a new mail folder.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name for the new folder",
                },
                "parent_folder_id": {
                    "type": "string",
                    "description": "Parent folder ID to create under (optional, defaults to root)",
                },
            },
            "required": ["name"],
        },
    ),
    Tool(
        name="m365_move_message",
        description="Move a message to a different folder.",
        inputSchema={
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The message ID to move",
                },
                "destination_folder_id": {
                    "type": "string",
                    "description": "The destination folder ID. Use m365_list_folders to see available folders.",
                },
            },
            "required": ["message_id", "destination_folder_id"],
        },
    ),
    Tool(
        name="m365_delete_message",
        description="Delete a message by moving it to the Deleted Items folder.",
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


def _format_folder(folder: dict[str, Any], include_children: bool = False) -> dict[str, Any]:
    """Format a folder into a summary object."""
    result = {
        "id": folder.get("id"),
        "name": folder.get("displayName"),
        "parent_folder_id": folder.get("parentFolderId"),
        "total_items": folder.get("totalItemCount", 0),
        "unread_items": folder.get("unreadItemCount", 0),
        "child_folder_count": folder.get("childFolderCount", 0),
    }

    if include_children and folder.get("childFolders"):
        result["child_folders"] = [
            _format_folder(child, include_children=True)
            for child in folder.get("childFolders", [])
        ]

    return result


async def handle_folder_tool(
    name: str, arguments: dict[str, Any], client: GraphClient
) -> dict[str, Any]:
    """Handle folder management tool calls.

    Args:
        name: Tool name
        arguments: Tool arguments
        client: Graph API client

    Returns:
        Tool result
    """
    if name == "m365_list_folders":
        include_children = arguments.get("include_children", False)

        result = await client.list_folders(include_children=include_children)

        folders = [
            _format_folder(folder, include_children=include_children)
            for folder in result.get("value", [])
        ]

        return {
            "folders": folders,
            "count": len(folders),
        }

    elif name == "m365_create_folder":
        name_arg = arguments.get("name")
        parent_folder_id = arguments.get("parent_folder_id")

        if not name_arg:
            return {"error": "Folder name is required"}

        folder = await client.create_folder(
            display_name=name_arg,
            parent_folder_id=parent_folder_id,
        )

        return {
            "success": True,
            "message": f"Folder '{name_arg}' created successfully",
            "folder_id": folder.get("id"),
            "folder_name": folder.get("displayName"),
        }

    elif name == "m365_move_message":
        message_id = arguments.get("message_id")
        destination_folder_id = arguments.get("destination_folder_id")

        if not message_id:
            return {"error": "message_id is required"}
        if not destination_folder_id:
            return {"error": "destination_folder_id is required"}

        moved = await client.move_message(
            message_id=message_id,
            destination_folder_id=destination_folder_id,
        )

        return {
            "success": True,
            "message": "Message moved successfully",
            "message_id": moved.get("id"),
            "new_folder_id": destination_folder_id,
        }

    elif name == "m365_delete_message":
        message_id = arguments.get("message_id")

        if not message_id:
            return {"error": "message_id is required"}

        result = await client.delete_message(message_id)

        return {
            **result,
            "message_id": message_id,
        }

    return {"error": f"Unknown folder tool: {name}"}
