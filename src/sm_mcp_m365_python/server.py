"""MCP Server for Microsoft 365 integration."""

import asyncio
import json
from typing import Any, Dict, List, Optional

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from .auth import M365OAuth, get_active_profile, set_active_profile, CREDENTIAL_PROFILES
from .graph import GraphClient
from .tools import ALL_TOOLS, get_tool_handler


class M365MCPServer:
    """MCP Server for Microsoft 365 operations."""

    def __init__(self):
        """Initialize the server."""
        self.server = Server("m365")
        self.oauth: Optional[M365OAuth] = None
        self.client: Optional[GraphClient] = None

        # Register handlers
        self._register_handlers()

    def _get_oauth(self) -> M365OAuth:
        """Get or create OAuth handler for current profile."""
        profile = get_active_profile()
        if self.oauth is None or self.oauth.profile != profile:
            self.oauth = M365OAuth(profile)
            self.client = GraphClient(self.oauth)
        return self.oauth

    def _get_client(self) -> GraphClient:
        """Get or create Graph client for current profile."""
        self._get_oauth()
        return self.client

    def _register_handlers(self):
        """Register MCP handlers."""

        @self.server.list_tools()
        async def list_tools() -> List[Tool]:
            """List available tools."""
            return ALL_TOOLS

        @self.server.call_tool()
        async def call_tool(name: str, arguments: Dict[str, Any]) -> List[TextContent]:
            """Call a tool by name."""
            try:
                handler = get_tool_handler(name)
                if handler is None:
                    return [TextContent(
                        type="text",
                        text=json.dumps({"error": f"Unknown tool: {name}"}),
                    )]

                # Pass server context to handler
                result = await handler(
                    arguments,
                    oauth=self._get_oauth(),
                    client=self._get_client(),
                )

                # Handle profile changes
                if name == "m365_set_profile":
                    profile = arguments.get("profile", "SM")
                    self.oauth = M365OAuth(profile)
                    self.client = GraphClient(self.oauth)

                return [TextContent(
                    type="text",
                    text=json.dumps(result, indent=2, default=str),
                )]

            except Exception as e:
                return [TextContent(
                    type="text",
                    text=json.dumps({"error": str(e)}),
                )]

    async def run(self):
        """Run the MCP server."""
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                self.server.create_initialization_options(),
            )
