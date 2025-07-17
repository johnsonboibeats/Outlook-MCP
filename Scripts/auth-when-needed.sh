#!/bin/bash
# Smart Outlook Authentication Script
# Only starts auth server when actually needed

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
TOKEN_FILE="$HOME/.outlook-mcp-tokens.json"
AUTH_SERVER_PORT=3333

# Colors
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

# Check if tokens exist and are valid
if [ -f "$TOKEN_FILE" ]; then
    # Check token validity using Node.js
    TOKEN_STATUS=$(node -e "
    const fs = require('fs');
    try {
        const tokens = JSON.parse(fs.readFileSync('$TOKEN_FILE', 'utf8'));
        const now = Date.now();
        const expiresAt = tokens.expires_at || 0;
        const refreshToken = tokens.refresh_token;
        
        if (refreshToken && expiresAt > now - 300000) { // 5 min buffer
            console.log('VALID');
        } else if (refreshToken) {
            console.log('NEEDS_REFRESH');
        } else {
            console.log('NEEDS_AUTH');
        }
    } catch (e) {
        console.log('NEEDS_AUTH');
    }
    " 2>/dev/null)
    
    if [ "$TOKEN_STATUS" = "VALID" ]; then
        echo -e "${GREEN}✓ Authentication is valid. No auth server needed.${NC}"
        exit 0
    elif [ "$TOKEN_STATUS" = "NEEDS_REFRESH" ]; then
        echo -e "${YELLOW}Token needs refresh, but auth server not required.${NC}"
        echo -e "${YELLOW}The MCP server will auto-refresh when needed.${NC}"
        exit 0
    fi
fi

# If we get here, we need the auth server
echo -e "${YELLOW}Authentication required. Starting auth server...${NC}"

# Check if already running
if lsof -i :$AUTH_SERVER_PORT >/dev/null 2>&1; then
    echo -e "${GREEN}Auth server already running on port $AUTH_SERVER_PORT${NC}"
else
    cd "$PROJECT_DIR"
    node auth/server.js &
    AUTH_PID=$!
    sleep 2
    echo -e "${GREEN}✓ Auth server started (PID: $AUTH_PID)${NC}"
fi

echo -e "${YELLOW}Please visit: ${GREEN}http://localhost:3333/auth${NC}"
echo -e "${YELLOW}After authentication, the server will stop automatically.${NC}"

# Wait for authentication to complete
echo -e "${YELLOW}Waiting for authentication...${NC}"
while true; do
    sleep 5
    if [ -f "$TOKEN_FILE" ]; then
        TOKEN_STATUS=$(node -e "
        const fs = require('fs');
        try {
            const tokens = JSON.parse(fs.readFileSync('$TOKEN_FILE', 'utf8'));
            if (tokens.access_token && tokens.refresh_token) {
                console.log('AUTHENTICATED');
            }
        } catch (e) {}
        " 2>/dev/null)
        
        if [ "$TOKEN_STATUS" = "AUTHENTICATED" ]; then
            echo -e "${GREEN}✓ Authentication successful!${NC}"
            
            # Stop the auth server
            if [ ! -z "$AUTH_PID" ]; then
                kill $AUTH_PID 2>/dev/null
                echo -e "${GREEN}✓ Auth server stopped${NC}"
            fi
            break
        fi
    fi
done
