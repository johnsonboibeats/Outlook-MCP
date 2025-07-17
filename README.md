# Outlook MCP Server

A Model Context Protocol (MCP) server that provides Claude with access to Microsoft Outlook through the Microsoft Graph API. Features multi-account support, remote deployment capabilities, and comprehensive email, calendar, and folder management.

## Quick Start

### 1. Install Dependencies
```bash
npm install
```

### 2. Configure Azure App Registration
1. Create an Azure App Registration at [Azure Portal](https://portal.azure.com)
2. Set redirect URIs:
   - `http://localhost:3333/auth/callback`
   - For remote: `https://your-domain.com/auth/callback`
3. Grant permissions: `Mail.ReadWrite`, `Calendars.ReadWrite`, `offline_access`

### 3. Environment Setup
Copy `.env.example` to `.env` and configure:
```bash
# Azure App Registration
MS_CLIENT_ID=your-azure-client-id
MS_CLIENT_SECRET=your-azure-client-secret
OUTLOOK_CLIENT_ID=your-azure-client-id
OUTLOOK_CLIENT_SECRET=your-azure-client-secret

# Server Configuration
USE_TEST_MODE=false
MCP_PORT=3001
```

### 4. Claude Desktop Configuration
Add to your Claude Desktop config:
```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": {
        "USE_TEST_MODE": "false"
      }
    }
  }
}
```

## Usage Modes

### Local Mode (Default)
```bash
npm start
```
Runs via stdio for Claude Desktop integration.

### Remote Mode
```bash
npm run start-remote
```
Runs HTTP server on port 3001 with MCP endpoint at `/mcp`.

### Test Mode
```bash
npm run test-mode
```
Runs with mock data for development.

## Features

### Multi-Account Support
- Manage multiple Outlook accounts simultaneously
- Separate authentication and token storage per account
- Account-specific operations and data isolation

### Email Management
- List, search, and read emails
- Send emails and create drafts
- Reply to messages
- Manage folders and move emails
- Handle attachments

### Calendar Management
- List calendar events
- Create, update, and delete events
- Accept/decline meeting invitations
- Manage calendar permissions

### Folder & Rules Management
- Create and manage email folders
- Set up email rules and filters
- Organize mailbox structure

## MCP Tools

### Authentication
- `authenticate` - Add new Outlook account
- `check-auth-status` - View account authentication status
- `list-accounts` - List all configured accounts
- `remove-account` - Remove an account

### Email Tools
- `list-emails` - List emails with filtering options
- `read-email` - Read specific email content
- `send-email` - Send new email
- `create-draft` - Create email draft
- `reply-to-email` - Reply to existing email
- `search-emails` - Search emails by criteria
- `mark-as-read` - Mark emails as read/unread

### Calendar Tools
- `list-calendar-events` - List calendar events
- `create-calendar-event` - Create new event
- `delete-calendar-event` - Delete event
- `accept-calendar-event` - Accept meeting invitation
- `decline-calendar-event` - Decline meeting invitation

### Folder Tools
- `list-folders` - List email folders
- `create-folder` - Create new folder
- `move-email` - Move email to folder

## Deployment

### Railway Deployment
1. Connect your GitHub repository to [Railway](https://railway.app)
2. Set environment variables in Railway dashboard
3. Deploy automatically on git push

Environment variables for Railway:
```bash
MS_CLIENT_ID=your-azure-client-id
MS_CLIENT_SECRET=your-azure-client-secret
OUTLOOK_CLIENT_ID=your-azure-client-id
OUTLOOK_CLIENT_SECRET=your-azure-client-secret
NODE_ENV=production
USE_TEST_MODE=false
```

### Docker Deployment
```dockerfile
FROM node:18-alpine
WORKDIR /app
COPY package*.json ./
RUN npm ci --production
COPY . .
EXPOSE 3001
CMD ["npm", "run", "start-remote"]
```

### Manual Server Setup
1. Clone repository
2. Install dependencies: `npm ci --production`
3. Configure environment variables
4. Start server: `npm run start-remote`
5. Configure reverse proxy (nginx/Apache) if needed

## Architecture

```
outlook-mcp/
├── auth/                  # Authentication & account management
│   ├── index.js          # Main auth module exports
│   ├── server.js         # OAuth server (moved from root)
│   ├── account-manager.js # Multi-account management
│   ├── token-manager.js  # Token storage & refresh
│   └── tools.js          # Auth MCP tools
├── calendar/             # Calendar operations
├── email/               # Email operations
├── folder/              # Folder management
├── rules/               # Email rules
├── utils/               # Shared utilities
│   ├── graph-api.js     # Graph API helpers
│   ├── odata-helpers.js # OData query builders
│   ├── permissions.js   # Permission checking
│   └── mock-data.js     # Test data
├── Scripts/             # Deployment & management scripts
├── config.js            # Configuration management
├── index.js             # Main server entry point
└── package.json         # Dependencies & scripts
```

## Security

### Best Practices
- Environment variables for all secrets
- Separate token storage per account
- Automatic token refresh
- HTTPS enforcement in production
- CORS configuration for remote access

### Token Management
- Tokens stored in `~/.outlook-mcp-accounts/`
- Automatic refresh before expiration
- Secure token storage with unique account IDs
- Cleanup on account removal

## Troubleshooting

### Common Issues

**"No accounts configured"**
- Solution: Use `authenticate` tool to add an account

**"Authentication required for account X"**
- Solution: Re-authenticate the specific account

**CORS errors in browser**
- Solution: Check CORS configuration for your domain

**Port already in use**
- Solution: Change `MCP_PORT` environment variable

### Health Checking
- Local: Use `check-auth-status` tool
- Remote: Visit `http://localhost:3001/health`

### Logs
- Server logs to stderr
- Authentication events logged with account IDs
- Error details for debugging

## Development

### Prerequisites
- Node.js 14+
- Azure App Registration
- Microsoft Graph API permissions

### Scripts
- `npm start` - Local stdio mode
- `npm run start-remote` - Remote HTTP mode
- `npm run auth-server` - Start OAuth server only
- `npm run test-mode` - Run with mock data
- `npm run inspect` - Run with MCP inspector

### Contributing
1. Fork the repository
2. Create feature branch
3. Make changes with tests
4. Update documentation
5. Submit pull request

## License

MIT License - see LICENSE file for details.

## Support

- GitHub Issues: [Report bugs](https://github.com/your-repo/issues)
- Documentation: This README
- MCP Protocol: [modelcontextprotocol.io](https://modelcontextprotocol.io)
- Microsoft Graph: [docs.microsoft.com/graph](https://docs.microsoft.com/graph)