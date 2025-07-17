/**
 * Multi-account management for Outlook MCP server
 * Supports multiple Outlook accounts with separate token management
 */
const fs = require('fs');
const path = require('path');
const { randomUUID } = require('crypto');
const config = require('../config');

class AccountManager {
  constructor() {
    this.accounts = new Map();
    this.accountsDir = path.join(process.env.HOME || process.cwd(), '.outlook-mcp-accounts');
    this.configPath = path.join(this.accountsDir, 'accounts.json');
    this.loadAccounts();
  }

  /**
   * Ensures the accounts directory exists
   */
  ensureAccountsDir() {
    if (!fs.existsSync(this.accountsDir)) {
      fs.mkdirSync(this.accountsDir, { recursive: true });
    }
  }

  /**
   * Loads all accounts from storage
   */
  loadAccounts() {
    try {
      this.ensureAccountsDir();
      
      if (fs.existsSync(this.configPath)) {
        const accountsData = JSON.parse(fs.readFileSync(this.configPath, 'utf8'));
        
        for (const [accountId, accountInfo] of Object.entries(accountsData)) {
          this.accounts.set(accountId, {
            ...accountInfo,
            tokens: this.loadAccountTokens(accountId)
          });
        }
        
        console.error(`[AccountManager] Loaded ${this.accounts.size} accounts`);
      }
    } catch (error) {
      console.error('[AccountManager] Error loading accounts:', error);
      this.accounts = new Map();
    }
  }

  /**
   * Saves all accounts to storage
   */
  saveAccounts() {
    try {
      this.ensureAccountsDir();
      
      const accountsData = {};
      for (const [accountId, account] of this.accounts.entries()) {
        // Save account info without tokens (tokens stored separately)
        const { tokens, ...accountInfo } = account;
        accountsData[accountId] = accountInfo;
        
        // Save tokens separately
        if (tokens) {
          this.saveAccountTokens(accountId, tokens);
        }
      }
      
      fs.writeFileSync(this.configPath, JSON.stringify(accountsData, null, 2));
      console.error(`[AccountManager] Saved ${this.accounts.size} accounts`);
    } catch (error) {
      console.error('[AccountManager] Error saving accounts:', error);
    }
  }

  /**
   * Loads tokens for a specific account
   */
  loadAccountTokens(accountId) {
    try {
      const tokenPath = path.join(this.accountsDir, `${accountId}.tokens.json`);
      if (fs.existsSync(tokenPath)) {
        const tokens = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
        
        // Check if token is expired
        const now = Date.now();
        if (tokens.expires_at && now > tokens.expires_at) {
          console.error(`[AccountManager] Tokens for account ${accountId} have expired`);
          return null;
        }
        
        return tokens;
      }
    } catch (error) {
      console.error(`[AccountManager] Error loading tokens for account ${accountId}:`, error);
    }
    return null;
  }

  /**
   * Saves tokens for a specific account
   */
  saveAccountTokens(accountId, tokens) {
    try {
      this.ensureAccountsDir();
      const tokenPath = path.join(this.accountsDir, `${accountId}.tokens.json`);
      fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
    } catch (error) {
      console.error(`[AccountManager] Error saving tokens for account ${accountId}:`, error);
    }
  }

  /**
   * Adds a new account
   */
  addAccount(userPrincipalName, displayName = null, tokens = null) {
    const accountId = randomUUID();
    const account = {
      id: accountId,
      userPrincipalName,
      displayName: displayName || userPrincipalName,
      createdAt: new Date().toISOString(),
      lastUsed: new Date().toISOString(),
      tokens
    };
    
    this.accounts.set(accountId, account);
    this.saveAccounts();
    
    console.error(`[AccountManager] Added account: ${displayName || userPrincipalName} (${accountId})`);
    return accountId;
  }

  /**
   * Removes an account
   */
  removeAccount(accountId) {
    if (this.accounts.has(accountId)) {
      const account = this.accounts.get(accountId);
      this.accounts.delete(accountId);
      
      // Remove token file
      try {
        const tokenPath = path.join(this.accountsDir, `${accountId}.tokens.json`);
        if (fs.existsSync(tokenPath)) {
          fs.unlinkSync(tokenPath);
        }
      } catch (error) {
        console.error(`[AccountManager] Error removing token file for ${accountId}:`, error);
      }
      
      this.saveAccounts();
      console.error(`[AccountManager] Removed account: ${account.displayName} (${accountId})`);
      return true;
    }
    return false;
  }

  /**
   * Updates account tokens
   */
  updateAccountTokens(accountId, tokens) {
    if (this.accounts.has(accountId)) {
      const account = this.accounts.get(accountId);
      account.tokens = tokens;
      account.lastUsed = new Date().toISOString();
      
      this.saveAccountTokens(accountId, tokens);
      this.saveAccounts();
      return true;
    }
    return false;
  }

  /**
   * Gets an account by ID
   */
  getAccount(accountId) {
    return this.accounts.get(accountId);
  }

  /**
   * Gets all accounts
   */
  getAllAccounts() {
    return Array.from(this.accounts.values()).map(account => ({
      id: account.id,
      userPrincipalName: account.userPrincipalName,
      displayName: account.displayName,
      createdAt: account.createdAt,
      lastUsed: account.lastUsed,
      hasValidTokens: account.tokens && account.tokens.access_token && 
                     account.tokens.expires_at > Date.now()
    }));
  }

  /**
   * Gets access token for an account
   */
  async getAccessToken(accountId) {
    const account = this.accounts.get(accountId);
    if (!account || !account.tokens) {
      return null;
    }

    // Check if token is expired or expires soon
    const now = Date.now();
    const expiresAt = account.tokens.expires_at || 0;
    const fiveMinutes = 5 * 60 * 1000;

    if (now + fiveMinutes > expiresAt && account.tokens.refresh_token) {
      console.error(`[AccountManager] Refreshing token for account ${accountId}`);
      const newTokens = await this.refreshAccessToken(accountId, account.tokens.refresh_token);
      if (newTokens) {
        return newTokens.access_token;
      }
      return null;
    }

    return account.tokens.access_token;
  }

  /**
   * Refreshes access token for an account
   */
  async refreshAccessToken(accountId, refreshToken) {
    try {
      const https = require('https');
      
      const tokenData = new URLSearchParams({
        client_id: config.AUTH_CONFIG.clientId,
        client_secret: config.AUTH_CONFIG.clientSecret,
        scope: config.AUTH_CONFIG.scopes.join(' '),
        refresh_token: refreshToken,
        grant_type: 'refresh_token'
      });

      const options = {
        hostname: 'login.microsoftonline.com',
        path: '/common/oauth2/v2.0/token',
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(tokenData.toString())
        }
      };

      return new Promise((resolve) => {
        const req = https.request(options, (res) => {
          let data = '';
          
          res.on('data', (chunk) => {
            data += chunk;
          });
          
          res.on('end', () => {
            try {
              const response = JSON.parse(data);
              
              if (res.statusCode === 200 && response.access_token) {
                const expiresIn = response.expires_in || 3600;
                const expiresAt = Date.now() + (expiresIn * 1000);
                
                const newTokens = {
                  ...response,
                  expires_at: expiresAt
                };
                
                this.updateAccountTokens(accountId, newTokens);
                resolve(newTokens);
              } else {
                console.error(`[AccountManager] Token refresh failed for ${accountId}:`, response);
                resolve(null);
              }
            } catch (error) {
              console.error(`[AccountManager] Error parsing refresh response for ${accountId}:`, error);
              resolve(null);
            }
          });
        });
        
        req.on('error', (error) => {
          console.error(`[AccountManager] Network error during token refresh for ${accountId}:`, error);
          resolve(null);
        });
        
        req.write(tokenData.toString());
        req.end();
      });
    } catch (error) {
      console.error(`[AccountManager] Error refreshing token for ${accountId}:`, error);
      return null;
    }
  }

  /**
   * Finds account by email address
   */
  findAccountByEmail(email) {
    for (const account of this.accounts.values()) {
      if (account.userPrincipalName.toLowerCase() === email.toLowerCase()) {
        return account;
      }
    }
    return null;
  }

  /**
   * Gets the default account (first one with valid tokens)
   */
  getDefaultAccount() {
    for (const account of this.accounts.values()) {
      if (account.tokens && account.tokens.access_token && 
          account.tokens.expires_at > Date.now()) {
        return account;
      }
    }
    return null;
  }
}

// Singleton instance
const accountManager = new AccountManager();

module.exports = accountManager;