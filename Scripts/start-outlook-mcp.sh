#!/bin/bash
# Outlook MCP Startup Script
# This script ensures the authentication server is running before starting the MCP server

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
AUTH_SERVER_PORT=3333
LOG_FILE="$HOME/Library/Logs/Claude/outlook-auth-server.log"

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${GREEN}Starting Outlook MCP Services...${NC}"

# Check if auth server is already running
if lsof -i :$AUTH_SERVER_PORT >/dev/null 2>&1; then
    echo -e "${YELLOW}Authentication server already running on port $AUTH_SERVER_PORT${NC}"
else
    echo -e "${GREEN}Starting authentication server...${NC}"
    cd "$PROJECT_DIR"
    nohup node auth/server.js > "$LOG_FILE" 2>&1 &
    AUTH_PID=$!
    
    # Wait for server to start
    sleep 2
    
    if lsof -i :$AUTH_SERVER_PORT >/dev/null 2>&1; then
        echo -e "${GREEN}✓ Authentication server started successfully (PID: $AUTH_PID)${NC}"
        echo $AUTH_PID > "$PROJECT_DIR/.auth-server.pid"
    else
        echo -e "${RED}✗ Failed to start authentication server${NC}"
        exit 1
    fi
fi

# Check authentication status
echo -e "${GREEN}Checking authentication status...${NC}"
cd "$PROJECT_DIR"
node -e "
const tokenManager = require('./auth/token-manager.js');
tokenManager.loadTokens().then(tokens => {
    if (tokens && tokens.access_token) {
        console.log('✓ Already authenticated');
    } else {
        console.log('⚠ Not authenticated - please visit: http://localhost:3333/auth');
    }
}).catch(err => {
    console.log('⚠ Not authenticated - please visit: http://localhost:3333/auth');
});
"

echo -e "${GREEN}MCP server is ready to use!${NC}"
echo -e "${YELLOW}Note: Keep this terminal open or run with 'nohup' to keep the auth server running${NC}"
