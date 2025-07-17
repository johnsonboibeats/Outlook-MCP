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
