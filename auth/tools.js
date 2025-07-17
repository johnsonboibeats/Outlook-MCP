/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const accountManager = require('./account-manager');

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `📧 MODULAR Outlook Assistant MCP Server v${config.SERVER_VERSION} 📧\n\nProvides access to Microsoft Outlook email, calendar, and contacts through Microsoft Graph API.\nImplemented with a modular architecture for improved maintainability.`
    }]
  };
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;
  
  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    // Create a test token with a 1-hour expiry
    tokenManager.createTestTokens();
    
    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }
  
  // For real authentication, generate an auth URL and instruct the user to visit it
  const authUrl = `${config.AUTH_CONFIG.authServerUrl}/auth?client_id=${config.AUTH_CONFIG.clientId}`;
  
  return {
    content: [{
      type: "text",
      text: `Authentication required. Please visit the following URL to authenticate with Microsoft: ${authUrl}\n\nAfter authentication, you will be redirected back to this application.`
    }]
  };
}

/**
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');
  
  const accounts = accountManager.getAllAccounts();
  
  if (accounts.length === 0) {
    return {
      content: [{ type: "text", text: "No accounts configured. Use 'list-accounts' to see available accounts or 'authenticate' to add a new account." }]
    };
  }
  
  const accountsStatus = accounts.map(account => {
    const status = account.hasValidTokens ? '✅ Ready' : '❌ Needs authentication';
    return `• ${account.displayName} (${account.userPrincipalName}): ${status}`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Account Status:\n${accountsStatus}\n\nTotal accounts: ${accounts.length}` 
    }]
  };
}

/**
 * List accounts tool handler
 * @returns {object} - MCP response
 */
async function handleListAccounts() {
  const accounts = accountManager.getAllAccounts();
  
  if (accounts.length === 0) {
    return {
      content: [{ 
        type: "text", 
        text: "No accounts configured. Use the 'authenticate' tool to add your first account." 
      }]
    };
  }
  
  const accountsList = accounts.map((account, index) => {
    const status = account.hasValidTokens ? '✅' : '❌';
    const lastUsed = new Date(account.lastUsed).toLocaleDateString();
    return `${index + 1}. ${status} ${account.displayName}\n   Email: ${account.userPrincipalName}\n   Last used: ${lastUsed}\n   Account ID: ${account.id}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Configured Accounts:\n\n${accountsList}\n\n✅ = Ready to use, ❌ = Needs authentication` 
    }]
  };
}

/**
 * Remove account tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRemoveAccount(args) {
  const { accountId } = args;
  
  if (!accountId) {
    return {
      content: [{ 
        type: "text", 
        text: "Please provide an accountId. Use 'list-accounts' to see available account IDs." 
      }]
    };
  }
  
  const account = accountManager.getAccount(accountId);
  if (!account) {
    return {
      content: [{ 
        type: "text", 
        text: `Account with ID '${accountId}' not found. Use 'list-accounts' to see available accounts.` 
      }]
    };
  }
  
  const success = accountManager.removeAccount(accountId);
  if (success) {
    return {
      content: [{ 
        type: "text", 
        text: `Successfully removed account: ${account.displayName} (${account.userPrincipalName})` 
      }]
    };
  } else {
    return {
      content: [{ 
        type: "text", 
        text: `Failed to remove account with ID '${accountId}'` 
      }]
    };
  }
}

// Tool definitions
const authTools = [
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status for all configured accounts",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  },
  {
    name: "list-accounts",
    description: "List all configured Outlook accounts with their status",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleListAccounts
  },
  {
    name: "remove-account",
    description: "Remove an Outlook account from the server",
    inputSchema: {
      type: "object",
      properties: {
        accountId: {
          type: "string",
          description: "The unique ID of the account to remove (use list-accounts to see IDs)"
        }
      },
      required: ["accountId"]
    },
    handler: handleRemoveAccount
  }
];

module.exports = {
  authTools,
  handleAuthenticate,
  handleCheckAuthStatus,
  handleListAccounts,
  handleRemoveAccount
};
