# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added
- Initial MCP server implementation for Microsoft 365 email operations
- Azure AD OAuth 2.0 authentication with PKCE
- Encrypted token storage at `~/.m365/tokens.enc`
- Message operations: list, search, get, get_thread, get_attachment
- Send operations: send_message, reply, forward
- Draft operations: list, create, update, delete, send
- Folder operations: list, create, move_message, delete_message
- Rate limiting for Microsoft Graph API compliance
- Cross-platform credential storage (macOS Keychain, Linux libsecret)

## Version History

| Version | Hash | Date | Author | Message |
|---------|------|------|--------|---------|
| (unreleased) | - | - | Greg Gowans | Initial implementation |
