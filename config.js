/**
 * Configuration for Outlook MCP Server
 * Centralized configuration management with environment variable support
 */
const path = require('path');
const os = require('os');

// Load environment variables
require('dotenv').config();

// Ensure we have a home directory path
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

// Environment configuration
const isDevelopment = process.env.NODE_ENV !== 'production';
const isTestMode = process.env.USE_TEST_MODE === 'true';

module.exports = {
  // Server information
  SERVER_NAME: 'outlook-mcp',
  SERVER_VERSION: '1.0.0',
  
  // Environment settings
  NODE_ENV: process.env.NODE_ENV || 'development',
  IS_DEVELOPMENT: isDevelopment,
  USE_TEST_MODE: isTestMode,
  
  // Server ports
  AUTH_SERVER_PORT: process.env.AUTH_SERVER_PORT || 3333,
  MCP_PORT: process.env.PORT || process.env.MCP_PORT || 3001,
  
  // Authentication configuration
  AUTH_CONFIG: {
    clientId: process.env.MS_CLIENT_ID || process.env.OUTLOOK_CLIENT_ID || '',
    clientSecret: process.env.MS_CLIENT_SECRET || process.env.OUTLOOK_CLIENT_SECRET || '',
    redirectUri: process.env.AUTH_REDIRECT_URI || (process.env.RAILWAY_PUBLIC_DOMAIN ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN}/auth/callback` : 'http://localhost:3333/auth/callback'),
    scopes: [
      'Mail.Read',
      'Mail.ReadWrite', 
      'Mail.Send',
      'User.Read',
      'Calendars.Read',
      'Calendars.ReadWrite',
      'offline_access'
    ],
    tokenStorePath: path.join(homeDir, '.outlook-mcp-tokens.json'),
    authServerUrl: process.env.AUTH_SERVER_URL || (process.env.RAILWAY_PUBLIC_DOMAIN ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN}` : 'http://localhost:3333')
  },
  
  // Microsoft Graph API configuration
  GRAPH_API: {
    endpoint: 'https://graph.microsoft.com/v1.0',
    timeout: 30000,
    retryAttempts: 3,
    retryDelay: 1000
  },
  
  // Field selections for optimal performance
  FIELDS: {
    email: {
      list: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
      detail: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders'
    },
    calendar: {
      list: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled',
      detail: 'id,subject,body,start,end,location,organizer,attendees,isAllDay,isCancelled,recurrence'
    },
    folder: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount'
  },
  
  // Pagination and limits
  PAGINATION: {
    defaultPageSize: 25,
    maxPageSize: 100,
    maxResults: 1000
  },
  
  // CORS configuration for remote mode
  CORS: {
    origin: process.env.CORS_ORIGIN || '*',
    exposedHeaders: ['Mcp-Session-Id'],
    allowedHeaders: ['Content-Type', 'mcp-session-id']
  },
  
  // Logging configuration
  LOGGING: {
    level: process.env.LOG_LEVEL || (isDevelopment ? 'debug' : 'info'),
    enableConsole: true,
    enableFile: false
  }
};