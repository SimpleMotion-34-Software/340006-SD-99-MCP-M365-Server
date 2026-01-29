"""Contact management tools for M365 MCP server."""

from typing import Any

from mcp.types import Tool

from ..graph import GraphClient

CONTACT_TOOLS = [
    Tool(
        name="m365_list_contacts",
        description="List contacts from the user's address book. Returns contact summaries with name, email, phone, and company.",
        inputSchema={
            "type": "object",
            "properties": {
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of contacts to return (1-1000). Default: 50",
                    "default": 50,
                },
                "skip": {
                    "type": "integer",
                    "description": "Number of contacts to skip for pagination. Default: 0",
                    "default": 0,
                },
                "search": {
                    "type": "string",
                    "description": "Optional search query to filter contacts by name or email.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_get_contact",
        description="Get the full details of a specific contact by its ID.",
        inputSchema={
            "type": "object",
            "properties": {
                "contact_id": {
                    "type": "string",
                    "description": "The contact ID to retrieve",
                },
            },
            "required": ["contact_id"],
        },
    ),
    Tool(
        name="m365_create_contact",
        description="Create a new contact in the user's address book.",
        inputSchema={
            "type": "object",
            "properties": {
                "given_name": {
                    "type": "string",
                    "description": "Contact's first name",
                },
                "surname": {
                    "type": "string",
                    "description": "Contact's last name",
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
                "department": {
                    "type": "string",
                    "description": "Department",
                },
                "notes": {
                    "type": "string",
                    "description": "Notes about the contact",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_update_contact",
        description="Update an existing contact. Only the fields you provide will be updated.",
        inputSchema={
            "type": "object",
            "properties": {
                "contact_id": {
                    "type": "string",
                    "description": "The contact ID to update",
                },
                "given_name": {
                    "type": "string",
                    "description": "Updated first name",
                },
                "surname": {
                    "type": "string",
                    "description": "Updated last name",
                },
                "email_addresses": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Updated list of email addresses",
                },
                "business_phones": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Updated list of business phone numbers",
                },
                "mobile_phone": {
                    "type": "string",
                    "description": "Updated mobile phone number",
                },
                "company_name": {
                    "type": "string",
                    "description": "Updated company name",
                },
                "job_title": {
                    "type": "string",
                    "description": "Updated job title",
                },
                "department": {
                    "type": "string",
                    "description": "Updated department",
                },
                "notes": {
                    "type": "string",
                    "description": "Updated notes",
                },
            },
            "required": ["contact_id"],
        },
    ),
    Tool(
        name="m365_delete_contact",
        description="Delete a contact from the user's address book.",
        inputSchema={
            "type": "object",
            "properties": {
                "contact_id": {
                    "type": "string",
                    "description": "The contact ID to delete",
                },
            },
            "required": ["contact_id"],
        },
    ),
    Tool(
        name="m365_search_contacts",
        description="Search for contacts by name, email, or company.",
        inputSchema={
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search query to find contacts",
                },
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of results to return. Default: 25",
                    "default": 25,
                },
            },
            "required": ["query"],
        },
    ),
]


def _format_contact_summary(contact: dict[str, Any]) -> dict[str, Any]:
    """Format a contact into a summary object."""
    emails = contact.get("emailAddresses", [])
    primary_email = emails[0].get("address") if emails else None

    return {
        "id": contact.get("id"),
        "display_name": contact.get("displayName"),
        "given_name": contact.get("givenName"),
        "surname": contact.get("surname"),
        "email": primary_email,
        "email_addresses": [e.get("address") for e in emails],
        "mobile_phone": contact.get("mobilePhone"),
        "business_phones": contact.get("businessPhones", []),
        "company_name": contact.get("companyName"),
        "job_title": contact.get("jobTitle"),
    }


def _format_contact_detail(contact: dict[str, Any]) -> dict[str, Any]:
    """Format a contact with full details."""
    emails = contact.get("emailAddresses", [])

    return {
        "id": contact.get("id"),
        "display_name": contact.get("displayName"),
        "given_name": contact.get("givenName"),
        "surname": contact.get("surname"),
        "email_addresses": [{"address": e.get("address"), "name": e.get("name")} for e in emails],
        "mobile_phone": contact.get("mobilePhone"),
        "business_phones": contact.get("businessPhones", []),
        "home_phones": contact.get("homePhones", []),
        "company_name": contact.get("companyName"),
        "job_title": contact.get("jobTitle"),
        "department": contact.get("department"),
        "office_location": contact.get("officeLocation"),
        "personal_notes": contact.get("personalNotes"),
        "birthday": contact.get("birthday"),
        "business_address": contact.get("businessAddress"),
        "home_address": contact.get("homeAddress"),
        "created_datetime": contact.get("createdDateTime"),
        "last_modified_datetime": contact.get("lastModifiedDateTime"),
    }


async def handle_contact_tool(
    name: str, arguments: dict[str, Any], client: GraphClient
) -> dict[str, Any]:
    """Handle contact management tool calls.

    Args:
        name: Tool name
        arguments: Tool arguments
        client: Graph API client

    Returns:
        Tool result
    """
    if name == "m365_list_contacts":
        limit = min(arguments.get("limit", 50), 1000)
        skip = arguments.get("skip", 0)
        search = arguments.get("search")

        result = await client.list_contacts(
            top=limit,
            skip=skip,
            search=search,
        )

        contacts = [_format_contact_summary(c) for c in result.get("value", [])]

        return {
            "contacts": contacts,
            "count": len(contacts),
            "has_more": "@odata.nextLink" in result,
        }

    elif name == "m365_get_contact":
        contact_id = arguments.get("contact_id")

        if not contact_id:
            return {"error": "contact_id is required"}

        contact = await client.get_contact(contact_id)
        return _format_contact_detail(contact)

    elif name == "m365_create_contact":
        result = await client.create_contact(
            given_name=arguments.get("given_name"),
            surname=arguments.get("surname"),
            email_addresses=arguments.get("email_addresses"),
            business_phones=arguments.get("business_phones"),
            mobile_phone=arguments.get("mobile_phone"),
            company_name=arguments.get("company_name"),
            job_title=arguments.get("job_title"),
            department=arguments.get("department"),
            notes=arguments.get("notes"),
        )

        return {
            "success": True,
            "message": "Contact created successfully",
            "contact_id": result.get("id"),
            "display_name": result.get("displayName"),
        }

    elif name == "m365_update_contact":
        contact_id = arguments.get("contact_id")

        if not contact_id:
            return {"error": "contact_id is required"}

        result = await client.update_contact(
            contact_id=contact_id,
            given_name=arguments.get("given_name"),
            surname=arguments.get("surname"),
            email_addresses=arguments.get("email_addresses"),
            business_phones=arguments.get("business_phones"),
            mobile_phone=arguments.get("mobile_phone"),
            company_name=arguments.get("company_name"),
            job_title=arguments.get("job_title"),
            department=arguments.get("department"),
            notes=arguments.get("notes"),
        )

        return {
            "success": True,
            "message": "Contact updated successfully",
            "contact_id": result.get("id"),
            "display_name": result.get("displayName"),
        }

    elif name == "m365_delete_contact":
        contact_id = arguments.get("contact_id")

        if not contact_id:
            return {"error": "contact_id is required"}

        await client.delete_contact(contact_id)

        return {
            "success": True,
            "message": "Contact deleted successfully",
        }

    elif name == "m365_search_contacts":
        query = arguments.get("query")
        limit = min(arguments.get("limit", 25), 1000)

        if not query:
            return {"error": "Search query is required"}

        result = await client.search_contacts(
            query=query,
            top=limit,
        )

        contacts = [_format_contact_summary(c) for c in result.get("value", [])]

        return {
            "contacts": contacts,
            "count": len(contacts),
            "query": query,
        }

    return {"error": f"Unknown contact tool: {name}"}
