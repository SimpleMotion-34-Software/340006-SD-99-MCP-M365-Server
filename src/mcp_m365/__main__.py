"""Entry point for running the M365 MCP server."""

import asyncio
import sys

from .server import M365MCPServer


def main():
    """Run the M365 MCP server."""
    server = M365MCPServer()
    asyncio.run(server.run())


if __name__ == "__main__":
    main()
