#!/usr/bin/env node
/**
 * Outlook MCP Server - Main entry point
 * 
 * A Model Context Protocol server that provides access to
 * Microsoft Outlook through the Microsoft Graph API.
 * 
 * Supports both local stdio and remote HTTP streamable transport.
 */
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const { StreamableHTTPServerTransport } = require("@modelcontextprotocol/sdk/server/streamableHttp.js");
const express = require('express');
const cors = require('cors');
const { randomUUID } = require('crypto');
const config = require('./config');

// Import module tools
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');

// Log startup information
console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

// Combine all tools
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools
  // Future modules: contactsTools, etc.
];

// Create server with tools capabilities
const server = new Server(
  { name: config.SERVER_NAME, version: config.SERVER_VERSION },
  { 
    capabilities: { 
      tools: TOOLS.reduce((acc, tool) => {
        acc[tool.name] = {};
        return acc;
      }, {})
    } 
  }
);

// Handle all requests
server.fallbackRequestHandler = async (request) => {
  try {
    const { method, params, id } = request;
    console.error(`REQUEST: ${method} [${id}]`);
    
    // Initialize handler
    if (method === "initialize") {
      console.error(`INITIALIZE REQUEST: ID [${id}]`);
      return {
        protocolVersion: "2024-11-05",
        capabilities: { 
          tools: TOOLS.reduce((acc, tool) => {
            acc[tool.name] = {};
            return acc;
          }, {})
        },
        serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION }
      };
    }
    
    // Tools list handler
    if (method === "tools/list") {
      console.error(`TOOLS LIST REQUEST: ID [${id}]`);
      console.error(`TOOLS COUNT: ${TOOLS.length}`);
      console.error(`TOOLS NAMES: ${TOOLS.map(t => t.name).join(', ')}`);
      
      return {
        tools: TOOLS.map(tool => ({
          name: tool.name,
          description: tool.description,
          inputSchema: tool.inputSchema
        }))
      };
    }
    
    // Required empty responses for other capabilities
    if (method === "resources/list") return { resources: [] };
    if (method === "prompts/list") return { prompts: [] };
    
    // Tool call handler
    if (method === "tools/call") {
      try {
        const { name, arguments: args = {} } = params || {};
        
        console.error(`TOOL CALL: ${name}`);
        
        // Find the tool handler
        const tool = TOOLS.find(t => t.name === name);
        
        if (tool && tool.handler) {
          return await tool.handler(args);
        }
        
        // Tool not found
        return {
          error: {
            code: -32601,
            message: `Tool not found: ${name}`
          }
        };
      } catch (error) {
        console.error(`Error in tools/call:`, error);
        return {
          error: {
            code: -32603,
            message: `Error processing tool call: ${error.message}`
          }
        };
      }
    }
    
    // For any other method, return method not found
    return {
      error: {
        code: -32601,
        message: `Method not found: ${method}`
      }
    };
  } catch (error) {
    console.error(`Error in fallbackRequestHandler:`, error);
    return {
      error: {
        code: -32603,
        message: `Error processing request: ${error.message}`
      }
    };
  }
};

// Make the script executable
process.on('SIGTERM', () => {
  console.error('SIGTERM received but staying alive');
});

// Check for remote mode
const isRemoteMode = process.argv.includes('--remote');

if (isRemoteMode) {
  // Start remote HTTP server with streamable transport
  const app = express();
  app.use(cors(config.CORS));
  app.use(express.json());
  
  // Store transports by session ID
  const transports = {};
  
  // Health check endpoint
  app.get('/health', (req, res) => {
    res.json({ status: 'ok', server: config.SERVER_NAME, version: config.SERVER_VERSION });
  });

  // Auth endpoints for Railway deployment
  app.get('/auth', (req, res) => {
    if (!config.AUTH_CONFIG.clientId || !config.AUTH_CONFIG.clientSecret) {
      return res.status(500).json({ error: 'Microsoft Graph API credentials not configured' });
    }
    
    const clientId = req.query.client_id || config.AUTH_CONFIG.clientId;
    const authParams = new URLSearchParams({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: config.AUTH_CONFIG.redirectUri,
      scope: config.AUTH_CONFIG.scopes.join(' '),
      response_mode: 'query',
      state: Date.now().toString()
    });
    
    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${authParams}`;
    res.redirect(authUrl);
  });

  app.get('/auth/callback', async (req, res) => {
    if (req.query.error) {
      return res.status(400).json({ 
        error: req.query.error, 
        description: req.query.error_description 
      });
    }
    
    if (!req.query.code) {
      return res.status(400).json({ error: 'No authorization code received' });
    }
    
    try {
      // Exchange code for tokens
      const tokenData = new URLSearchParams({
        client_id: config.AUTH_CONFIG.clientId,
        client_secret: config.AUTH_CONFIG.clientSecret,
        code: req.query.code,
        redirect_uri: config.AUTH_CONFIG.redirectUri,
        grant_type: 'authorization_code'
      });
      
      const tokenResponse = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: tokenData
      });
      
      const tokens = await tokenResponse.json();
      
      if (tokens.error) {
        return res.status(400).json({ error: tokens.error, description: tokens.error_description });
      }
      
      // Store tokens (simplified for Railway - in production use secure storage)
      // Note: Railway ephemeral storage means tokens won't persist across restarts
      global.outlookTokens = tokens;
      
      res.json({ 
        message: 'Authentication completed successfully!',
        expires_in: tokens.expires_in,
        scope: tokens.scope
      });
    } catch (error) {
      console.error('Token exchange error:', error);
      res.status(500).json({ error: 'Failed to exchange code for tokens' });
    }
  });
  
  // MCP endpoint - handles POST, GET, DELETE
  app.all('/mcp', async (req, res) => {
    try {
      const sessionId = req.headers['mcp-session-id'];
      let transport;
      
      if (sessionId && transports[sessionId]) {
        // Reuse existing transport
        transport = transports[sessionId];
      } else if (!sessionId && req.method === 'POST' && req.body && req.body.method === 'initialize') {
        // New initialization request
        transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
          onsessioninitialized: (sessionId) => {
            transports[sessionId] = transport;
          }
        });
        
        // Clean up on close
        transport.onclose = () => {
          if (transport.sessionId) {
            delete transports[transport.sessionId];
          }
        };
        
        // Connect server to transport
        await server.connect(transport);
      } else {
        // Invalid request
        return res.status(400).json({
          jsonrpc: '2.0',
          error: {
            code: -32000,
            message: 'Bad Request: No valid session ID provided'
          },
          id: null
        });
      }
      
      // Handle the request
      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      console.error('MCP request error:', error);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: '2.0',
          error: {
            code: -32603,
            message: `Internal server error: ${error.message}`
          },
          id: null
        });
      }
    }
  });
  
  const PORT = config.MCP_PORT;
  const HOST = process.env.RAILWAY_PUBLIC_DOMAIN || `localhost:${PORT}`;
  const PROTOCOL = process.env.RAILWAY_PUBLIC_DOMAIN ? 'https' : 'http';
  
  app.listen(PORT, () => {
    console.error(`${config.SERVER_NAME} remote server listening on port ${PORT}`);
    console.error(`MCP endpoint: ${PROTOCOL}://${HOST}/mcp`);
    console.error(`Health check: ${PROTOCOL}://${HOST}/health`);
  });
} else {
  // Start local stdio server
  const transport = new StdioServerTransport();
  server.connect(transport)
    .then(() => console.error(`${config.SERVER_NAME} connected and listening`))
    .catch(error => {
      console.error(`Connection error: ${error.message}`);
      process.exit(1);
    });
}
