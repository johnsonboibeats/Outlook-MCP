#!/bin/bash
# Outlook MCP Shutdown Script
# This script cleanly stops the authentication server

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
PID_FILE="$PROJECT_DIR/.auth-server.pid"

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${YELLOW}Stopping Outlook MCP Services...${NC}"

# Stop auth server using PID file
if [ -f "$PID_FILE" ]; then
    PID=$(cat "$PID_FILE")
    if kill -0 $PID 2>/dev/null; then
        kill $PID
        echo -e "${GREEN}✓ Stopped authentication server (PID: $PID)${NC}"
        rm "$PID_FILE"
    else
        echo -e "${YELLOW}Authentication server not running (stale PID file)${NC}"
        rm "$PID_FILE"
    fi
else
    # Try to find and stop by port
    AUTH_PID=$(lsof -ti :3333)
    if [ ! -z "$AUTH_PID" ]; then
        kill $AUTH_PID
        echo -e "${GREEN}✓ Stopped authentication server (PID: $AUTH_PID)${NC}"
    else
        echo -e "${YELLOW}No authentication server found running${NC}"
    fi
fi

echo -e "${GREEN}All services stopped${NC}"
