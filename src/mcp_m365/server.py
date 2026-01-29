"""MCP Server for Microsoft 365 integration."""

import asyncio
import json
import logging
import sys
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent

from .auth import M365OAuth
from .graph import GraphClient
from .tools import (
    ALL_TOOLS,
    handle_auth_tool,
    handle_chat_tool,
    handle_contact_tool,
    handle_draft_tool,
    handle_folder_tool,
    handle_message_tool,
    handle_send_tool,
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)


class M365MCPServer:
    """MCP Server for Microsoft 365 integration."""

    def __init__(self):
        """Initialize the M365 MCP server."""
        self.server = Server("m365-mcp")
        self.oauth = M365OAuth()
        self.client = GraphClient(self.oauth)

        # Register handlers
        self._register_handlers()

    def _register_handlers(self) -> None:
        """Register MCP server handlers."""

        @self.server.list_tools()
        async def list_tools():
            """List all available M365 tools."""
            return ALL_TOOLS

        @self.server.call_tool()
        async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
            """Handle tool calls."""
            logger.info(f"Tool call: {name} with arguments: {arguments}")

            try:
                result = await self._handle_tool(name, arguments)
            except Exception as e:
                logger.exception(f"Error handling tool {name}")
                result = {"error": str(e)}

            return [TextContent(type="text", text=json.dumps(result, indent=2))]

    async def _handle_tool(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        """Route tool calls to appropriate handlers.

        Args:
            name: Tool name
            arguments: Tool arguments

        Returns:
            Tool result
        """
        # Authentication tools
        if name.startswith("m365_auth") or name == "m365_connect" or name == "m365_disconnect":
            return await handle_auth_tool(name, arguments, self.oauth)

        # Check authentication for other tools
        tokens = await self.oauth.get_valid_tokens()
        if not tokens:
            return {
                "error": "Not authenticated with Microsoft 365",
                "message": "Use m365_connect to connect to Microsoft 365 first",
            }

        # Message reading tools
        if name.startswith("m365_") and any(
            x in name for x in ["list_messages", "search_messages", "get_message", "get_thread", "get_attachment"]
        ):
            return await handle_message_tool(name, arguments, self.client)

        # Send tools
        if name.startswith("m365_") and any(
            x in name for x in ["send_message", "reply", "forward"]
        ):
            return await handle_send_tool(name, arguments, self.client)

        # Draft tools
        if name.startswith("m365_") and "draft" in name:
            return await handle_draft_tool(name, arguments, self.client)

        # Folder tools
        if name.startswith("m365_") and any(
            x in name for x in ["list_folders", "create_folder", "move_message", "delete_message"]
        ):
            return await handle_folder_tool(name, arguments, self.client)

        # Teams chat tools
        if name.startswith("m365_") and "chat" in name:
            return await handle_chat_tool(name, arguments, self.client)

        # Contact tools
        if name.startswith("m365_") and "contact" in name:
            return await handle_contact_tool(name, arguments, self.client)

        return {"error": f"Unknown tool: {name}"}

    async def run(self) -> None:
        """Run the MCP server."""
        logger.info("Starting M365 MCP server")

        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                self.server.create_initialization_options(),
            )


def main() -> None:
    """Main entry point."""
    server = M365MCPServer()
    asyncio.run(server.run())


if __name__ == "__main__":
    main()
