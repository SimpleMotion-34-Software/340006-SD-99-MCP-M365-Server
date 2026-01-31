"""Contact management tools for M365 MCP Server."""

from typing import Any, Dict, List

from mcp.types import Tool

from ..auth import M365OAuth
from ..graph import GraphClient


CONTACT_TOOLS: List[Tool] = [
    Tool(
        name="m365_list_contacts",
        description="List contacts from the user's address book",
        inputSchema={
            "type": "object",
            "properties": {
                "top": {
                    "type": "integer",
                    "description": "Number of contacts to return",
                    "default": 50,
                },
                "skip": {
                    "type": "integer",
                    "description": "Number of contacts to skip for pagination",
                    "default": 0,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_get_contact",
        description="Get a specific contact by ID",
        inputSchema={
            "type": "object",
            "properties": {
                "contact_id": {
                    "type": "string",
                    "description": "The contact ID",
                },
            },
            "required": ["contact_id"],
        },
    ),
    Tool(
        name="m365_search_contacts",
        description="Search contacts by name or email",
        inputSchema={
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search query (name or email)",
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
        name="m365_create_contact",
        description="Create a new contact",
        inputSchema={
            "type": "object",
            "properties": {
                "display_name": {
                    "type": "string",
                    "description": "Full display name",
                },
                "given_name": {
                    "type": "string",
                    "description": "First name",
                },
                "surname": {
                    "type": "string",
                    "description": "Last name",
                },
                "email_addresses": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of email addresses",
                },
                "business_phones": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of business phone numbers",
                },
                "mobile_phone": {
                    "type": "string",
                    "description": "Mobile phone number",
                },
                "company_name": {
                    "type": "string",
                    "description": "Company name",
                },
                "job_title": {
                    "type": "string",
                    "description": "Job title",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_update_contact",
        description="Update an existing contact",
        inputSchema={
            "type": "object",
            "properties": {
                "contact_id": {
                    "type": "string",
                    "description": "The contact ID",
                },
                "display_name": {
                    "type": "string",
                    "description": "Full display name",
                },
                "given_name": {
                    "type": "string",
                    "description": "First name",
                },
                "surname": {
                    "type": "string",
                    "description": "Last name",
                },
                "email_addresses": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of email addresses",
                },
                "business_phones": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of business phone numbers",
                },
                "mobile_phone": {
                    "type": "string",
                    "description": "Mobile phone number",
                },
                "company_name": {
                    "type": "string",
                    "description": "Company name",
                },
                "job_title": {
                    "type": "string",
                    "description": "Job title",
                },
            },
            "required": ["contact_id"],
        },
    ),
    Tool(
        name="m365_delete_contact",
        description="Delete a contact",
        inputSchema={
            "type": "object",
            "properties": {
                "contact_id": {
                    "type": "string",
                    "description": "The contact ID",
                },
            },
            "required": ["contact_id"],
        },
    ),
]


def _format_contact(contact: Dict[str, Any]) -> Dict[str, Any]:
    """Format a contact for display."""
    email_addresses = []
    for e in contact.get("emailAddresses", []):
        if e.get("address"):
            email_addresses.append(e["address"])

    return {
        "id": contact.get("id"),
        "display_name": contact.get("displayName"),
        "given_name": contact.get("givenName"),
        "surname": contact.get("surname"),
        "email_addresses": email_addresses,
        "business_phones": contact.get("businessPhones", []),
        "mobile_phone": contact.get("mobilePhone"),
        "company_name": contact.get("companyName"),
        "job_title": contact.get("jobTitle"),
    }


async def handle_list_contacts(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_list_contacts tool call."""
    top = min(arguments.get("top", 50), 100)
    skip = arguments.get("skip", 0)

    contacts = await client.list_contacts(top=top, skip=skip)

    return {
        "count": len(contacts),
        "contacts": [_format_contact(c) for c in contacts],
    }


async def handle_get_contact(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_get_contact tool call."""
    contact_id = arguments["contact_id"]

    contact = await client.get_contact(contact_id)

    return _format_contact(contact)


async def handle_search_contacts(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_search_contacts tool call."""
    query = arguments["query"]
    top = min(arguments.get("top", 25), 50)

    contacts = await client.search_contacts(query=query, top=top)

    return {
        "query": query,
        "count": len(contacts),
        "contacts": [_format_contact(c) for c in contacts],
    }


async def handle_create_contact(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_create_contact tool call."""
    contact = await client.create_contact(
        display_name=arguments.get("display_name"),
        given_name=arguments.get("given_name"),
        surname=arguments.get("surname"),
        email_addresses=arguments.get("email_addresses"),
        business_phones=arguments.get("business_phones"),
        mobile_phone=arguments.get("mobile_phone"),
        company_name=arguments.get("company_name"),
        job_title=arguments.get("job_title"),
    )

    return {
        "status": "created",
        "contact": _format_contact(contact),
    }


async def handle_update_contact(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_update_contact tool call."""
    contact_id = arguments["contact_id"]

    contact = await client.update_contact(
        contact_id=contact_id,
        display_name=arguments.get("display_name"),
        given_name=arguments.get("given_name"),
        surname=arguments.get("surname"),
        email_addresses=arguments.get("email_addresses"),
        business_phones=arguments.get("business_phones"),
        mobile_phone=arguments.get("mobile_phone"),
        company_name=arguments.get("company_name"),
        job_title=arguments.get("job_title"),
    )

    return {
        "status": "updated",
        "contact": _format_contact(contact),
    }


async def handle_delete_contact(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_delete_contact tool call."""
    contact_id = arguments["contact_id"]

    await client.delete_contact(contact_id)

    return {
        "status": "deleted",
        "contact_id": contact_id,
    }


CONTACT_HANDLERS = {
    "m365_list_contacts": handle_list_contacts,
    "m365_get_contact": handle_get_contact,
    "m365_search_contacts": handle_search_contacts,
    "m365_create_contact": handle_create_contact,
    "m365_update_contact": handle_update_contact,
    "m365_delete_contact": handle_delete_contact,
}
