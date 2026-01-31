"""Microsoft Graph API client."""

from typing import Any, Dict, List, Optional
import aiohttp

from ..auth import M365OAuth


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class GraphClient:
    """Async client for Microsoft Graph API."""

    def __init__(self, oauth: M365OAuth):
        """Initialize the Graph client.

        Args:
            oauth: The OAuth handler for authentication.
        """
        self.oauth = oauth

    async def _get_headers(self) -> Dict[str, str]:
        """Get authorization headers.

        Returns:
            Dictionary of headers including Authorization.

        Raises:
            RuntimeError: If not authenticated.
        """
        tokens = await self.oauth.get_valid_tokens()
        if not tokens:
            raise RuntimeError("Not authenticated. Use m365_connect to authenticate.")

        return {
            "Authorization": f"Bearer {tokens.access_token}",
            "Content-Type": "application/json",
        }

    async def _request(
        self,
        method: str,
        endpoint: str,
        params: Optional[Dict[str, Any]] = None,
        json: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Make a request to the Graph API.

        Args:
            method: HTTP method (GET, POST, PATCH, DELETE)
            endpoint: API endpoint (e.g., '/me/messages')
            params: Query parameters
            json: JSON body for POST/PATCH

        Returns:
            Response JSON as dictionary.

        Raises:
            RuntimeError: If the request fails.
        """
        headers = await self._get_headers()
        url = f"{GRAPH_BASE_URL}{endpoint}"

        async with aiohttp.ClientSession() as session:
            async with session.request(
                method,
                url,
                headers=headers,
                params=params,
                json=json,
            ) as resp:
                if resp.status == 204:
                    return {}

                if resp.status >= 400:
                    error = await resp.text()
                    raise RuntimeError(f"Graph API error ({resp.status}): {error}")

                return await resp.json()

    async def get(
        self,
        endpoint: str,
        params: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Make a GET request.

        Args:
            endpoint: API endpoint
            params: Query parameters

        Returns:
            Response JSON.
        """
        return await self._request("GET", endpoint, params=params)

    async def post(
        self,
        endpoint: str,
        json: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Make a POST request.

        Args:
            endpoint: API endpoint
            json: JSON body

        Returns:
            Response JSON.
        """
        return await self._request("POST", endpoint, json=json)

    async def patch(
        self,
        endpoint: str,
        json: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Make a PATCH request.

        Args:
            endpoint: API endpoint
            json: JSON body

        Returns:
            Response JSON.
        """
        return await self._request("PATCH", endpoint, json=json)

    async def delete(self, endpoint: str) -> Dict[str, Any]:
        """Make a DELETE request.

        Args:
            endpoint: API endpoint

        Returns:
            Response JSON (empty for 204).
        """
        return await self._request("DELETE", endpoint)

    # ========== Messages ==========

    async def list_messages(
        self,
        folder: str = "inbox",
        top: int = 25,
        skip: int = 0,
        select: Optional[List[str]] = None,
        filter_query: Optional[str] = None,
        order_by: str = "receivedDateTime desc",
    ) -> List[Dict[str, Any]]:
        """List messages in a folder.

        Args:
            folder: Folder name or ID (default: inbox)
            top: Number of messages to return
            skip: Number of messages to skip
            select: Fields to select
            filter_query: OData filter query
            order_by: Sort order

        Returns:
            List of message objects.
        """
        if select is None:
            select = [
                "id",
                "subject",
                "from",
                "toRecipients",
                "receivedDateTime",
                "isRead",
                "hasAttachments",
                "bodyPreview",
            ]

        params = {
            "$top": top,
            "$skip": skip,
            "$select": ",".join(select),
            "$orderby": order_by,
        }

        if filter_query:
            params["$filter"] = filter_query

        # Use well-known folder names or folder ID
        if folder.lower() in ["inbox", "drafts", "sentitems", "deleteditems", "junkemail"]:
            endpoint = f"/me/mailFolders/{folder}/messages"
        else:
            endpoint = f"/me/mailFolders/{folder}/messages"

        result = await self.get(endpoint, params)
        return result.get("value", [])

    async def get_message(
        self,
        message_id: str,
        select: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Get a specific message.

        Args:
            message_id: The message ID
            select: Fields to select

        Returns:
            Message object.
        """
        params = {}
        if select:
            params["$select"] = ",".join(select)

        return await self.get(f"/me/messages/{message_id}", params)

    async def search_messages(
        self,
        query: str,
        top: int = 25,
    ) -> List[Dict[str, Any]]:
        """Search messages.

        Args:
            query: Search query
            top: Number of results

        Returns:
            List of matching messages.
        """
        params = {
            "$search": f'"{query}"',
            "$top": top,
            "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview",
        }

        result = await self.get("/me/messages", params)
        return result.get("value", [])

    async def send_message(
        self,
        subject: str,
        body: str,
        to_recipients: List[str],
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        is_html: bool = False,
        save_to_sent: bool = True,
    ) -> Dict[str, Any]:
        """Send a new message.

        Args:
            subject: Message subject
            body: Message body
            to_recipients: List of email addresses
            cc_recipients: List of CC email addresses
            bcc_recipients: List of BCC email addresses
            is_html: Whether body is HTML
            save_to_sent: Whether to save to sent items

        Returns:
            Empty dict on success.
        """
        message = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body,
            },
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to_recipients
            ],
        }

        if cc_recipients:
            message["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc_recipients
            ]

        if bcc_recipients:
            message["bccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in bcc_recipients
            ]

        payload = {
            "message": message,
            "saveToSentItems": save_to_sent,
        }

        return await self.post("/me/sendMail", payload)

    async def reply_to_message(
        self,
        message_id: str,
        comment: str,
        reply_all: bool = False,
    ) -> Dict[str, Any]:
        """Reply to a message.

        Args:
            message_id: The message ID to reply to
            comment: Reply text
            reply_all: Whether to reply all

        Returns:
            Empty dict on success.
        """
        endpoint = f"/me/messages/{message_id}/{'replyAll' if reply_all else 'reply'}"
        return await self.post(endpoint, {"comment": comment})

    async def forward_message(
        self,
        message_id: str,
        to_recipients: List[str],
        comment: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Forward a message.

        Args:
            message_id: The message ID to forward
            to_recipients: List of email addresses
            comment: Optional comment to add

        Returns:
            Empty dict on success.
        """
        payload = {
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to_recipients
            ],
        }
        if comment:
            payload["comment"] = comment

        return await self.post(f"/me/messages/{message_id}/forward", payload)

    # ========== Drafts ==========

    async def list_drafts(self, top: int = 25) -> List[Dict[str, Any]]:
        """List draft messages.

        Args:
            top: Number of drafts to return

        Returns:
            List of draft messages.
        """
        return await self.list_messages(folder="drafts", top=top)

    async def create_draft(
        self,
        subject: str,
        body: str,
        to_recipients: Optional[List[str]] = None,
        cc_recipients: Optional[List[str]] = None,
        is_html: bool = False,
    ) -> Dict[str, Any]:
        """Create a draft message.

        Args:
            subject: Message subject
            body: Message body
            to_recipients: List of email addresses
            cc_recipients: List of CC email addresses
            is_html: Whether body is HTML

        Returns:
            Created draft message.
        """
        message = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body,
            },
        }

        if to_recipients:
            message["toRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in to_recipients
            ]

        if cc_recipients:
            message["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc_recipients
            ]

        return await self.post("/me/messages", message)

    async def update_draft(
        self,
        message_id: str,
        subject: Optional[str] = None,
        body: Optional[str] = None,
        to_recipients: Optional[List[str]] = None,
        is_html: bool = False,
    ) -> Dict[str, Any]:
        """Update a draft message.

        Args:
            message_id: The draft message ID
            subject: New subject
            body: New body
            to_recipients: New recipients
            is_html: Whether body is HTML

        Returns:
            Updated draft message.
        """
        message = {}

        if subject is not None:
            message["subject"] = subject

        if body is not None:
            message["body"] = {
                "contentType": "HTML" if is_html else "Text",
                "content": body,
            }

        if to_recipients is not None:
            message["toRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in to_recipients
            ]

        return await self.patch(f"/me/messages/{message_id}", message)

    async def delete_draft(self, message_id: str) -> Dict[str, Any]:
        """Delete a draft message.

        Args:
            message_id: The draft message ID

        Returns:
            Empty dict on success.
        """
        return await self.delete(f"/me/messages/{message_id}")

    async def send_draft(self, message_id: str) -> Dict[str, Any]:
        """Send a draft message.

        Args:
            message_id: The draft message ID

        Returns:
            Empty dict on success.
        """
        return await self.post(f"/me/messages/{message_id}/send", {})

    # ========== Folders ==========

    async def list_folders(
        self,
        parent_folder_id: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """List mail folders.

        Args:
            parent_folder_id: Parent folder ID for child folders

        Returns:
            List of folder objects.
        """
        if parent_folder_id:
            endpoint = f"/me/mailFolders/{parent_folder_id}/childFolders"
        else:
            endpoint = "/me/mailFolders"

        result = await self.get(endpoint)
        return result.get("value", [])

    async def create_folder(
        self,
        display_name: str,
        parent_folder_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Create a mail folder.

        Args:
            display_name: Folder name
            parent_folder_id: Parent folder ID

        Returns:
            Created folder object.
        """
        if parent_folder_id:
            endpoint = f"/me/mailFolders/{parent_folder_id}/childFolders"
        else:
            endpoint = "/me/mailFolders"

        return await self.post(endpoint, {"displayName": display_name})

    async def move_message(
        self,
        message_id: str,
        destination_folder_id: str,
    ) -> Dict[str, Any]:
        """Move a message to another folder.

        Args:
            message_id: The message ID
            destination_folder_id: Target folder ID

        Returns:
            Moved message object.
        """
        return await self.post(
            f"/me/messages/{message_id}/move",
            {"destinationId": destination_folder_id},
        )

    async def delete_message(self, message_id: str) -> Dict[str, Any]:
        """Delete a message.

        Args:
            message_id: The message ID

        Returns:
            Empty dict on success.
        """
        return await self.delete(f"/me/messages/{message_id}")

    # ========== Contacts ==========

    async def list_contacts(
        self,
        top: int = 100,
        skip: int = 0,
        select: Optional[List[str]] = None,
        filter_query: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """List contacts.

        Args:
            top: Number of contacts to return
            skip: Number to skip
            select: Fields to select
            filter_query: OData filter query

        Returns:
            List of contact objects.
        """
        if select is None:
            select = [
                "id",
                "displayName",
                "givenName",
                "surname",
                "emailAddresses",
                "businessPhones",
                "mobilePhone",
                "companyName",
                "jobTitle",
            ]

        params = {
            "$top": top,
            "$skip": skip,
            "$select": ",".join(select),
        }

        if filter_query:
            params["$filter"] = filter_query

        result = await self.get("/me/contacts", params)
        return result.get("value", [])

    async def get_contact(self, contact_id: str) -> Dict[str, Any]:
        """Get a specific contact.

        Args:
            contact_id: The contact ID

        Returns:
            Contact object.
        """
        return await self.get(f"/me/contacts/{contact_id}")

    async def search_contacts(
        self,
        query: str,
        top: int = 25,
    ) -> List[Dict[str, Any]]:
        """Search contacts.

        Args:
            query: Search query
            top: Number of results

        Returns:
            List of matching contacts.
        """
        # Use filter for simple search
        filter_query = f"contains(displayName, '{query}') or contains(emailAddresses/any(e:e/address), '{query}')"

        return await self.list_contacts(
            top=top,
            filter_query=filter_query,
        )

    async def create_contact(
        self,
        display_name: Optional[str] = None,
        given_name: Optional[str] = None,
        surname: Optional[str] = None,
        email_addresses: Optional[List[str]] = None,
        business_phones: Optional[List[str]] = None,
        mobile_phone: Optional[str] = None,
        company_name: Optional[str] = None,
        job_title: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Create a contact.

        Args:
            display_name: Full name
            given_name: First name
            surname: Last name
            email_addresses: List of email addresses
            business_phones: List of business phone numbers
            mobile_phone: Mobile phone number
            company_name: Company name
            job_title: Job title

        Returns:
            Created contact object.
        """
        contact = {}

        if display_name:
            contact["displayName"] = display_name
        if given_name:
            contact["givenName"] = given_name
        if surname:
            contact["surname"] = surname
        if email_addresses:
            contact["emailAddresses"] = [
                {"address": addr, "name": addr} for addr in email_addresses
            ]
        if business_phones:
            contact["businessPhones"] = business_phones
        if mobile_phone:
            contact["mobilePhone"] = mobile_phone
        if company_name:
            contact["companyName"] = company_name
        if job_title:
            contact["jobTitle"] = job_title

        return await self.post("/me/contacts", contact)

    async def update_contact(
        self,
        contact_id: str,
        display_name: Optional[str] = None,
        given_name: Optional[str] = None,
        surname: Optional[str] = None,
        email_addresses: Optional[List[str]] = None,
        business_phones: Optional[List[str]] = None,
        mobile_phone: Optional[str] = None,
        company_name: Optional[str] = None,
        job_title: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Update a contact.

        Args:
            contact_id: The contact ID
            display_name: Full name
            given_name: First name
            surname: Last name
            email_addresses: List of email addresses
            business_phones: List of business phone numbers
            mobile_phone: Mobile phone number
            company_name: Company name
            job_title: Job title

        Returns:
            Updated contact object.
        """
        contact = {}

        if display_name is not None:
            contact["displayName"] = display_name
        if given_name is not None:
            contact["givenName"] = given_name
        if surname is not None:
            contact["surname"] = surname
        if email_addresses is not None:
            contact["emailAddresses"] = [
                {"address": addr, "name": addr} for addr in email_addresses
            ]
        if business_phones is not None:
            contact["businessPhones"] = business_phones
        if mobile_phone is not None:
            contact["mobilePhone"] = mobile_phone
        if company_name is not None:
            contact["companyName"] = company_name
        if job_title is not None:
            contact["jobTitle"] = job_title

        return await self.patch(f"/me/contacts/{contact_id}", contact)

    async def delete_contact(self, contact_id: str) -> Dict[str, Any]:
        """Delete a contact.

        Args:
            contact_id: The contact ID

        Returns:
            Empty dict on success.
        """
        return await self.delete(f"/me/contacts/{contact_id}")

    # ========== Attachments ==========

    async def get_attachment(
        self,
        message_id: str,
        attachment_id: str,
    ) -> Dict[str, Any]:
        """Get a message attachment.

        Args:
            message_id: The message ID
            attachment_id: The attachment ID

        Returns:
            Attachment object with content.
        """
        return await self.get(f"/me/messages/{message_id}/attachments/{attachment_id}")

    async def list_attachments(self, message_id: str) -> List[Dict[str, Any]]:
        """List attachments for a message.

        Args:
            message_id: The message ID

        Returns:
            List of attachment objects.
        """
        result = await self.get(f"/me/messages/{message_id}/attachments")
        return result.get("value", [])
