"""Microsoft Graph API client wrapper with rate limiting and error handling."""

import asyncio
import base64
import logging
from datetime import datetime
from typing import Any

import aiohttp

from ..auth import M365OAuth

logger = logging.getLogger(__name__)

# Microsoft Graph API base URL
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

# Rate limit configuration (Microsoft Graph: 10,000 requests per 10 minutes per app)
RATE_LIMIT_REQUESTS = 100  # per minute (conservative)
RATE_LIMIT_WINDOW = 60  # seconds


class GraphAPIError(Exception):
    """Microsoft Graph API error."""

    def __init__(self, message: str, status_code: int | None = None, details: Any = None):
        super().__init__(message)
        self.status_code = status_code
        self.details = details


class GraphClient:
    """Wrapper around Microsoft Graph API with rate limiting and automatic token refresh."""

    def __init__(self, oauth: M365OAuth):
        """Initialize Graph client.

        Args:
            oauth: OAuth handler for authentication
        """
        self.oauth = oauth
        self._request_times: list[float] = []

    async def _get_user_path(self) -> str:
        """Get the appropriate user path for API calls.

        For delegated auth (with refresh token): returns "me"
        For client credentials (no refresh token): returns "users/{user_email}"

        Returns:
            User path string for Graph API endpoints
        """
        tokens = await self.oauth.get_valid_tokens()
        if not tokens:
            return "me"  # Will fail at auth check anyway

        # Client credentials flow has no refresh token
        if not tokens.refresh_token and tokens.user_email:
            return f"users/{tokens.user_email}"

        return "me"

    def _parse_error_message(self, error_text: str) -> str:
        """Parse Microsoft Graph API error into readable message.

        Args:
            error_text: Raw error response text

        Returns:
            Human-readable error message
        """
        import json

        try:
            error_data = json.loads(error_text)
            error_obj = error_data.get("error", {})

            if "message" in error_obj:
                return f"Graph API error: {error_obj['message']}"

            if "code" in error_obj:
                return f"Graph API error ({error_obj['code']}): {error_text[:200]}"

        except json.JSONDecodeError:
            pass

        return f"Graph API error: {error_text[:200]}"

    async def _check_rate_limit(self) -> None:
        """Check and enforce rate limiting."""
        now = datetime.now().timestamp()

        # Remove old requests outside the window
        self._request_times = [t for t in self._request_times if now - t < RATE_LIMIT_WINDOW]

        # If at limit, wait
        if len(self._request_times) >= RATE_LIMIT_REQUESTS:
            wait_time = RATE_LIMIT_WINDOW - (now - self._request_times[0])
            if wait_time > 0:
                logger.info(f"Rate limit reached, waiting {wait_time:.1f}s")
                await asyncio.sleep(wait_time)

        self._request_times.append(now)

    async def _request(
        self,
        method: str,
        endpoint: str,
        data: dict | None = None,
        params: dict | None = None,
        raw_response: bool = False,
    ) -> dict[str, Any] | bytes:
        """Make authenticated request to Microsoft Graph API.

        Args:
            method: HTTP method
            endpoint: API endpoint (without base URL)
            data: Request body data
            params: Query parameters
            raw_response: If True, return raw bytes instead of JSON

        Returns:
            Response data (dict or bytes)

        Raises:
            GraphAPIError: If request fails
        """
        tokens = await self.oauth.get_valid_tokens()
        if not tokens:
            raise GraphAPIError("Not authenticated with Microsoft 365", status_code=401)

        await self._check_rate_limit()

        url = f"{GRAPH_API_BASE}/{endpoint}"
        headers = {
            "Authorization": f"Bearer {tokens.access_token}",
            "Accept": "application/json",
        }

        if data:
            headers["Content-Type"] = "application/json"

        async with aiohttp.ClientSession() as session:
            for attempt in range(3):  # Retry up to 3 times
                async with session.request(
                    method,
                    url,
                    json=data,
                    params=params,
                    headers=headers,
                ) as response:
                    # Handle rate limiting
                    if response.status == 429:
                        retry_after = int(response.headers.get("Retry-After", 60))
                        logger.warning(f"Rate limited, retrying after {retry_after}s")
                        await asyncio.sleep(retry_after)
                        continue

                    # Handle other errors
                    if response.status >= 400:
                        error_text = await response.text()
                        user_message = self._parse_error_message(error_text)
                        raise GraphAPIError(
                            user_message,
                            status_code=response.status,
                            details=error_text,
                        )

                    if raw_response:
                        return await response.read()
                    return await response.json()

            raise GraphAPIError("Max retries exceeded", status_code=429)

    # ==================== Messages ====================

    async def list_messages(
        self,
        folder: str = "inbox",
        top: int = 25,
        skip: int = 0,
        select: list[str] | None = None,
        filter_query: str | None = None,
        order_by: str = "receivedDateTime desc",
    ) -> dict[str, Any]:
        """List messages in a folder.

        Args:
            folder: Folder name (inbox, sentItems, drafts, deletedItems) or folder ID
            top: Number of messages to return (max 1000)
            skip: Number of messages to skip (for pagination)
            select: Fields to return (default: common fields)
            filter_query: OData filter query
            order_by: Sort order

        Returns:
            Response with messages array and pagination info
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
                "importance",
                "flag",
            ]

        params: dict[str, Any] = {
            "$top": min(top, 1000),
            "$skip": skip,
            "$select": ",".join(select),
            "$orderby": order_by,
        }

        if filter_query:
            params["$filter"] = filter_query

        user_path = await self._get_user_path()
        endpoint = f"{user_path}/mailFolders/{folder}/messages"
        return await self._request("GET", endpoint, params=params)

    async def search_messages(
        self,
        query: str,
        top: int = 25,
        folder: str | None = None,
    ) -> dict[str, Any]:
        """Search messages using KQL (Keyword Query Language).

        Args:
            query: Search query (e.g., "from:john subject:meeting")
            top: Number of results to return
            folder: Optional folder to search in (default: all folders)

        Returns:
            Response with matching messages
        """
        params: dict[str, Any] = {
            "$search": f'"{query}"',
            "$top": min(top, 1000),
            "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview",
        }

        user_path = await self._get_user_path()
        if folder:
            endpoint = f"{user_path}/mailFolders/{folder}/messages"
        else:
            endpoint = f"{user_path}/messages"

        return await self._request("GET", endpoint, params=params)

    async def get_message(
        self,
        message_id: str,
        include_body: bool = True,
    ) -> dict[str, Any]:
        """Get a message by ID.

        Args:
            message_id: Message ID
            include_body: Whether to include full body content

        Returns:
            Message details
        """
        select = [
            "id",
            "subject",
            "from",
            "toRecipients",
            "ccRecipients",
            "bccRecipients",
            "replyTo",
            "receivedDateTime",
            "sentDateTime",
            "isRead",
            "hasAttachments",
            "importance",
            "flag",
            "conversationId",
            "conversationIndex",
        ]

        if include_body:
            select.extend(["body", "bodyPreview"])

        params = {"$select": ",".join(select)}
        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/messages/{message_id}", params=params)

    async def get_thread(
        self,
        conversation_id: str,
        top: int = 50,
    ) -> dict[str, Any]:
        """Get all messages in a conversation thread.

        Args:
            conversation_id: Conversation ID from a message
            top: Maximum messages to return

        Returns:
            Response with messages in the thread
        """
        params: dict[str, Any] = {
            "$filter": f"conversationId eq '{conversation_id}'",
            "$top": top,
            "$orderby": "receivedDateTime asc",
            "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,body",
        }

        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/messages", params=params)

    async def get_attachments(self, message_id: str) -> dict[str, Any]:
        """Get attachments for a message.

        Args:
            message_id: Message ID

        Returns:
            Response with attachments array
        """
        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/messages/{message_id}/attachments")

    async def get_attachment_content(
        self,
        message_id: str,
        attachment_id: str,
    ) -> dict[str, Any]:
        """Get a specific attachment with content.

        Args:
            message_id: Message ID
            attachment_id: Attachment ID

        Returns:
            Attachment details including contentBytes
        """
        user_path = await self._get_user_path()
        return await self._request(
            "GET",
            f"{user_path}/messages/{message_id}/attachments/{attachment_id}",
        )

    # ==================== Send ====================

    async def send_message(
        self,
        to: list[str],
        subject: str,
        body: str,
        body_type: str = "HTML",
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        importance: str = "normal",
        attachments: list[dict[str, Any]] | None = None,
        save_to_sent: bool = True,
    ) -> dict[str, Any]:
        """Send a new email message.

        Args:
            to: List of recipient email addresses
            subject: Email subject
            body: Email body content
            body_type: Body content type ("HTML" or "Text")
            cc: List of CC recipients
            bcc: List of BCC recipients
            importance: Message importance ("low", "normal", "high")
            attachments: List of attachments (each with name, contentType, contentBytes)
            save_to_sent: Whether to save to Sent Items folder

        Returns:
            Send result
        """
        message: dict[str, Any] = {
            "subject": subject,
            "body": {
                "contentType": body_type,
                "content": body,
            },
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
            "importance": importance,
        }

        if cc:
            message["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in cc]
        if bcc:
            message["bccRecipients"] = [{"emailAddress": {"address": addr}} for addr in bcc]
        if attachments:
            message["attachments"] = attachments

        data = {
            "message": message,
            "saveToSentItems": save_to_sent,
        }

        # sendMail endpoint returns 202 Accepted with no body
        user_path = await self._get_user_path()
        await self._request("POST", f"{user_path}/sendMail", data=data)
        return {"success": True, "message": "Email sent successfully"}

    async def reply_to_message(
        self,
        message_id: str,
        body: str,
        body_type: str = "HTML",
        reply_all: bool = False,
    ) -> dict[str, Any]:
        """Reply to an existing message.

        Args:
            message_id: Original message ID to reply to
            body: Reply body content
            body_type: Body content type ("HTML" or "Text")
            reply_all: Whether to reply to all recipients

        Returns:
            Reply result
        """
        endpoint = "replyAll" if reply_all else "reply"

        data = {
            "message": {
                "body": {
                    "contentType": body_type,
                    "content": body,
                },
            },
        }

        user_path = await self._get_user_path()
        await self._request("POST", f"{user_path}/messages/{message_id}/{endpoint}", data=data)
        return {"success": True, "message": "Reply sent successfully"}

    async def forward_message(
        self,
        message_id: str,
        to: list[str],
        comment: str | None = None,
    ) -> dict[str, Any]:
        """Forward a message.

        Args:
            message_id: Message ID to forward
            to: List of recipient email addresses
            comment: Optional comment to add to the forwarded message

        Returns:
            Forward result
        """
        data: dict[str, Any] = {
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
        }

        if comment:
            data["comment"] = comment

        user_path = await self._get_user_path()
        await self._request("POST", f"{user_path}/messages/{message_id}/forward", data=data)
        return {"success": True, "message": "Message forwarded successfully"}

    # ==================== Drafts ====================

    async def list_drafts(self, top: int = 25, skip: int = 0) -> dict[str, Any]:
        """List draft messages.

        Args:
            top: Number of drafts to return
            skip: Number to skip for pagination

        Returns:
            Response with drafts array
        """
        return await self.list_messages(folder="drafts", top=top, skip=skip)

    async def create_draft(
        self,
        to: list[str] | None = None,
        subject: str = "",
        body: str = "",
        body_type: str = "HTML",
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        importance: str = "normal",
    ) -> dict[str, Any]:
        """Create a new draft message.

        Args:
            to: List of recipient email addresses
            subject: Email subject
            body: Email body content
            body_type: Body content type ("HTML" or "Text")
            cc: List of CC recipients
            bcc: List of BCC recipients
            importance: Message importance

        Returns:
            Created draft message
        """
        message: dict[str, Any] = {
            "subject": subject,
            "body": {
                "contentType": body_type,
                "content": body,
            },
            "importance": importance,
        }

        if to:
            message["toRecipients"] = [{"emailAddress": {"address": addr}} for addr in to]
        if cc:
            message["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in cc]
        if bcc:
            message["bccRecipients"] = [{"emailAddress": {"address": addr}} for addr in bcc]

        user_path = await self._get_user_path()
        return await self._request("POST", f"{user_path}/messages", data=message)

    async def update_draft(
        self,
        message_id: str,
        to: list[str] | None = None,
        subject: str | None = None,
        body: str | None = None,
        body_type: str = "HTML",
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        importance: str | None = None,
    ) -> dict[str, Any]:
        """Update an existing draft.

        Args:
            message_id: Draft message ID
            to: Updated recipients
            subject: Updated subject
            body: Updated body content
            body_type: Body content type
            cc: Updated CC recipients
            bcc: Updated BCC recipients
            importance: Updated importance

        Returns:
            Updated draft message
        """
        message: dict[str, Any] = {}

        if to is not None:
            message["toRecipients"] = [{"emailAddress": {"address": addr}} for addr in to]
        if subject is not None:
            message["subject"] = subject
        if body is not None:
            message["body"] = {
                "contentType": body_type,
                "content": body,
            }
        if cc is not None:
            message["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in cc]
        if bcc is not None:
            message["bccRecipients"] = [{"emailAddress": {"address": addr}} for addr in bcc]
        if importance is not None:
            message["importance"] = importance

        user_path = await self._get_user_path()
        return await self._request("PATCH", f"{user_path}/messages/{message_id}", data=message)

    async def delete_draft(self, message_id: str) -> dict[str, Any]:
        """Delete a draft message.

        Args:
            message_id: Draft message ID

        Returns:
            Deletion result
        """
        user_path = await self._get_user_path()
        await self._request("DELETE", f"{user_path}/messages/{message_id}")
        return {"success": True, "message": "Draft deleted"}

    async def send_draft(self, message_id: str) -> dict[str, Any]:
        """Send a draft message.

        Args:
            message_id: Draft message ID

        Returns:
            Send result
        """
        user_path = await self._get_user_path()
        await self._request("POST", f"{user_path}/messages/{message_id}/send")
        return {"success": True, "message": "Draft sent successfully"}

    # ==================== Folders ====================

    async def list_folders(self, include_children: bool = False) -> dict[str, Any]:
        """List mail folders.

        Args:
            include_children: Whether to include child folders

        Returns:
            Response with folders array
        """
        params: dict[str, Any] = {
            "$select": "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount",
        }

        if include_children:
            params["$expand"] = "childFolders"

        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/mailFolders", params=params)

    async def create_folder(
        self,
        display_name: str,
        parent_folder_id: str | None = None,
    ) -> dict[str, Any]:
        """Create a new mail folder.

        Args:
            display_name: Name of the folder
            parent_folder_id: Parent folder ID (optional, defaults to root)

        Returns:
            Created folder
        """
        data = {"displayName": display_name}

        user_path = await self._get_user_path()
        if parent_folder_id:
            endpoint = f"{user_path}/mailFolders/{parent_folder_id}/childFolders"
        else:
            endpoint = f"{user_path}/mailFolders"

        return await self._request("POST", endpoint, data=data)

    async def move_message(
        self,
        message_id: str,
        destination_folder_id: str,
    ) -> dict[str, Any]:
        """Move a message to a different folder.

        Args:
            message_id: Message ID to move
            destination_folder_id: Target folder ID

        Returns:
            Moved message
        """
        data = {"destinationId": destination_folder_id}
        user_path = await self._get_user_path()
        return await self._request("POST", f"{user_path}/messages/{message_id}/move", data=data)

    async def delete_message(self, message_id: str) -> dict[str, Any]:
        """Delete a message (move to Deleted Items).

        Args:
            message_id: Message ID to delete

        Returns:
            Deletion result
        """
        # Move to deletedItems folder
        folders = await self.list_folders()
        deleted_folder = None
        for folder in folders.get("value", []):
            if folder.get("displayName") == "Deleted Items":
                deleted_folder = folder.get("id")
                break

        if deleted_folder:
            return await self.move_message(message_id, deleted_folder)
        else:
            # Fallback: permanent delete
            user_path = await self._get_user_path()
            await self._request("DELETE", f"{user_path}/messages/{message_id}")
            return {"success": True, "message": "Message permanently deleted"}

    async def mark_as_read(
        self,
        message_id: str,
        is_read: bool = True,
    ) -> dict[str, Any]:
        """Mark a message as read or unread.

        Args:
            message_id: Message ID
            is_read: True to mark as read, False to mark as unread

        Returns:
            Updated message
        """
        user_path = await self._get_user_path()
        return await self._request(
            "PATCH",
            f"{user_path}/messages/{message_id}",
            data={"isRead": is_read},
        )

    # ==================== Teams Chat ====================

    async def list_chats(self, top: int = 50) -> dict[str, Any]:
        """List user's Teams chats (1:1, group, meeting chats).

        Args:
            top: Maximum number of chats to return (default: 50)

        Returns:
            Response with chats array
        """
        params: dict[str, Any] = {
            "$top": min(top, 50),
            "$orderby": "lastMessagePreview/createdDateTime desc",
            "$expand": "lastMessagePreview",
        }

        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/chats", params=params)

    async def get_chat(self, chat_id: str) -> dict[str, Any]:
        """Get details of a specific chat.

        Args:
            chat_id: The chat ID

        Returns:
            Chat details including members and last message preview
        """
        params: dict[str, Any] = {
            "$expand": "lastMessagePreview,members",
        }

        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/chats/{chat_id}", params=params)

    async def get_chat_messages(
        self,
        chat_id: str,
        top: int = 50,
    ) -> dict[str, Any]:
        """Get messages from a specific chat.

        Args:
            chat_id: The chat ID
            top: Maximum number of messages to return (default: 50)

        Returns:
            Response with messages array
        """
        params: dict[str, Any] = {
            "$top": min(top, 50),
            "$orderby": "createdDateTime desc",
        }

        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/chats/{chat_id}/messages", params=params)

    async def get_chat_members(self, chat_id: str) -> dict[str, Any]:
        """Get members of a chat.

        Args:
            chat_id: The chat ID

        Returns:
            Response with members array
        """
        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/chats/{chat_id}/members")

    async def send_chat_message(
        self,
        chat_id: str,
        content: str,
        content_type: str = "html",
    ) -> dict[str, Any]:
        """Send a message to a chat.

        Args:
            chat_id: The chat ID
            content: Message content
            content_type: Content type ("html" or "text")

        Returns:
            Created message
        """
        data = {
            "body": {
                "contentType": content_type,
                "content": content,
            }
        }

        user_path = await self._get_user_path()
        return await self._request("POST", f"{user_path}/chats/{chat_id}/messages", data=data)

    async def search_chat_messages(
        self,
        query: str,
        top: int = 25,
    ) -> dict[str, Any]:
        """Search messages across all chats.

        For client credentials flow, this iterates through user's chats
        and searches messages in each chat.

        Args:
            query: Search query string
            top: Maximum number of results to return (default: 25)

        Returns:
            Response with matching messages
        """
        user_path = await self._get_user_path()
        query_lower = query.lower()

        # For client credentials, we need to iterate through chats
        # getAllMessages endpoint doesn't work with application permissions
        if user_path != "me":
            # Get user's chats
            chats_response = await self.list_chats(top=50)
            chats = chats_response.get("value", [])

            matching_messages = []
            for chat in chats:
                if len(matching_messages) >= top:
                    break

                chat_id = chat.get("id")
                if not chat_id:
                    continue

                # Get messages from this chat
                try:
                    messages_response = await self.get_chat_messages(chat_id, top=50)
                    messages = messages_response.get("value", [])

                    # Filter messages that match the query
                    for msg in messages:
                        if len(matching_messages) >= top:
                            break

                        body = msg.get("body", {})
                        content = body.get("content", "") or ""

                        if query_lower in content.lower():
                            # Add chat context to message
                            msg["chatId"] = chat_id
                            msg["chatTopic"] = chat.get("topic")
                            matching_messages.append(msg)

                except Exception:
                    # Skip chats we can't access
                    continue

            return {"value": matching_messages}

        # For delegated auth, use the getAllMessages endpoint
        params: dict[str, Any] = {
            "$search": f'"{query}"',
            "$top": min(top, 50),
        }

        return await self._request("GET", f"{user_path}/chats/getAllMessages", params=params)

    # ==================== Contacts ====================

    async def list_contacts(
        self,
        top: int = 50,
        skip: int = 0,
        search: str | None = None,
    ) -> dict[str, Any]:
        """List contacts from the user's address book.

        Args:
            top: Number of contacts to return (max 1000)
            skip: Number of contacts to skip (for pagination)
            search: Optional search query to filter contacts

        Returns:
            Response with contacts array
        """
        params: dict[str, Any] = {
            "$top": min(top, 1000),
            "$skip": skip,
            "$orderby": "displayName",
            "$select": "id,displayName,givenName,surname,emailAddresses,mobilePhone,businessPhones,companyName,jobTitle",
        }

        if search:
            # Search in displayName, givenName, surname, and emailAddresses
            params["$filter"] = (
                f"startswith(displayName,'{search}') or "
                f"startswith(givenName,'{search}') or "
                f"startswith(surname,'{search}')"
            )

        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/contacts", params=params)

    async def get_contact(self, contact_id: str) -> dict[str, Any]:
        """Get a contact by ID.

        Args:
            contact_id: Contact ID

        Returns:
            Contact details
        """
        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/contacts/{contact_id}")

    async def create_contact(
        self,
        given_name: str | None = None,
        surname: str | None = None,
        email_addresses: list[str] | None = None,
        business_phones: list[str] | None = None,
        mobile_phone: str | None = None,
        company_name: str | None = None,
        job_title: str | None = None,
        department: str | None = None,
        notes: str | None = None,
    ) -> dict[str, Any]:
        """Create a new contact.

        Args:
            given_name: Contact's first name
            surname: Contact's last name
            email_addresses: List of email addresses
            business_phones: List of business phone numbers
            mobile_phone: Mobile phone number
            company_name: Company name
            job_title: Job title
            department: Department
            notes: Notes about the contact

        Returns:
            Created contact
        """
        contact: dict[str, Any] = {}

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
        if department:
            contact["department"] = department
        if notes:
            contact["personalNotes"] = notes

        user_path = await self._get_user_path()
        return await self._request("POST", f"{user_path}/contacts", data=contact)

    async def update_contact(
        self,
        contact_id: str,
        given_name: str | None = None,
        surname: str | None = None,
        email_addresses: list[str] | None = None,
        business_phones: list[str] | None = None,
        mobile_phone: str | None = None,
        company_name: str | None = None,
        job_title: str | None = None,
        department: str | None = None,
        notes: str | None = None,
    ) -> dict[str, Any]:
        """Update an existing contact.

        Args:
            contact_id: Contact ID to update
            given_name: Updated first name
            surname: Updated last name
            email_addresses: Updated list of email addresses
            business_phones: Updated list of business phone numbers
            mobile_phone: Updated mobile phone number
            company_name: Updated company name
            job_title: Updated job title
            department: Updated department
            notes: Updated notes

        Returns:
            Updated contact
        """
        contact: dict[str, Any] = {}

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
        if department is not None:
            contact["department"] = department
        if notes is not None:
            contact["personalNotes"] = notes

        user_path = await self._get_user_path()
        return await self._request("PATCH", f"{user_path}/contacts/{contact_id}", data=contact)

    async def delete_contact(self, contact_id: str) -> None:
        """Delete a contact.

        Args:
            contact_id: Contact ID to delete
        """
        user_path = await self._get_user_path()
        await self._request("DELETE", f"{user_path}/contacts/{contact_id}")

    async def search_contacts(
        self,
        query: str,
        top: int = 25,
    ) -> dict[str, Any]:
        """Search contacts by name, email, or company.

        Args:
            query: Search query
            top: Maximum number of results to return

        Returns:
            Response with matching contacts
        """
        params: dict[str, Any] = {
            "$top": min(top, 1000),
            "$select": "id,displayName,givenName,surname,emailAddresses,mobilePhone,businessPhones,companyName,jobTitle",
            "$filter": (
                f"startswith(displayName,'{query}') or "
                f"startswith(givenName,'{query}') or "
                f"startswith(surname,'{query}') or "
                f"startswith(companyName,'{query}')"
            ),
        }

        user_path = await self._get_user_path()
        return await self._request("GET", f"{user_path}/contacts", params=params)
