# Outlook MCP - Avoiding Authentication Issues

## Quick Start
To avoid authentication server issues, use the provided scripts:

```bash
# Start the services
/Users/bj/Library/CloudStorage/GoogleDrive-johnsonboibeats@gmail.com/My Drive/Repos/Outlook-MCP/Scripts/start-outlook-mcp.sh

# Stop the services
/Users/bj/Library/CloudStorage/GoogleDrive-johnsonboibeats@gmail.com/My Drive/Repos/Outlook-MCP/Scripts/stop-outlook-mcp.sh
```

## Permanent Solutions

### Option 1: Auto-Start on Mac Login (Recommended)
Install the LaunchAgent to automatically start the auth server when you log in:

```bash
# Copy the plist file to LaunchAgents
cp /Users/bj/Library/CloudStorage/GoogleDrive-johnsonboibeats@gmail.com/My Drive/Repos/Outlook-MCP/Scripts/com.outlook-mcp.auth-server.plist ~/Library/LaunchAgents/

# Load the service
launchctl load ~/Library/LaunchAgents/com.outlook-mcp.auth-server.plist

# To uninstall later:
# launchctl unload ~/Library/LaunchAgents/com.outlook-mcp.auth-server.plist
# rm ~/Library/LaunchAgents/com.outlook-mcp.auth-server.plist
```

### Option 2: Add to Shell Profile
Add this to your ~/.zshrc or ~/.bash_profile:

```bash
# Auto-start Outlook MCP auth server
alias outlook-start="/Users/bj/Library/CloudStorage/GoogleDrive-johnsonboibeats@gmail.com/My Drive/Repos/Outlook-MCP/Scripts/start-outlook-mcp.sh"
alias outlook-stop="/Users/bj/Library/CloudStorage/GoogleDrive-johnsonboibeats@gmail.com/My Drive/Repos/Outlook-MCP/Scripts/stop-outlook-mcp.sh"

# Optional: Auto-start on new terminal sessions
# /Users/bj/Library/CloudStorage/GoogleDrive-johnsonboibeats@gmail.com/My Drive/Repos/Outlook-MCP/Scripts/start-outlook-mcp.sh >/dev/null 2>&1
```

### Option 3: Create a Desktop App
Create an Automator app that runs the start script:
1. Open Automator
2. Create new "Application"
3. Add "Run Shell Script" action
4. Enter: `/Users/bj/Library/CloudStorage/GoogleDrive-johnsonboibeats@gmail.com/My Drive/Repos/Outlook-MCP/Scripts/start-outlook-mcp.sh`
5. Save as "Outlook MCP.app" in Applications folder

## Troubleshooting

### Check if auth server is running:
```bash
lsof -i :3333
```

### View logs:
```bash
tail -f ~/Library/Logs/Claude/outlook-auth-server.log
```

### Force re-authentication:
```bash
rm ~/.outlook-mcp-tokens.json
/Users/bj/Library/CloudStorage/GoogleDrive-johnsonboibeats@gmail.com/My Drive/Repos/Outlook-MCP/Scripts/start-outlook-mcp.sh
# Then visit http://localhost:3333/auth
```

### Common Issues:
1. **Port 3333 already in use**: Another process is using the port
   - Solution: Kill the process or change the port in the config

2. **Token expired**: The refresh token has expired
   - Solution: Delete ~/.outlook-mcp-tokens.json and re-authenticate

3. **Invalid client secret**: The Azure app secret has expired
   - Solution: Generate a new client secret in Azure Portal and update .env file

## Best Practices
1. Use the LaunchAgent for automatic startup
2. Keep the auth server running in the background
3. Check logs regularly for any issues
4. Backup your .env file with credentials
5. Set calendar reminders to renew Azure client secrets before they expire
