/**
 * Authentication module for Outlook MCP server
 */
const tokenManager = require('./token-manager');
const accountManager = require('./account-manager');
const { authTools } = require('./tools');

/**
 * Ensures the user is authenticated and returns an access token
 * @param {string} accountId - Account ID to authenticate (optional, uses default if not provided)
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(accountId = null, forceNew = false) {
  if (forceNew) {
    // Force re-authentication
    throw new Error('Authentication required');
  }
  
  // If accountId is provided, use multi-account manager
  if (accountId) {
    const accessToken = await accountManager.getAccessToken(accountId);
    if (!accessToken) {
      throw new Error(`Authentication required for account ${accountId}`);
    }
    return accessToken;
  }
  
  // Try multi-account manager first (get default account)
  const defaultAccount = accountManager.getDefaultAccount();
  if (defaultAccount) {
    const accessToken = await accountManager.getAccessToken(defaultAccount.id);
    if (accessToken) {
      return accessToken;
    }
  }
  
  // Fallback to legacy single-account token manager
  const accessToken = await tokenManager.getAccessTokenAsync();
  if (!accessToken) {
    throw new Error('Authentication required - no accounts configured');
  }
  
  return accessToken;
}

module.exports = {
  tokenManager,
  accountManager,
  authTools,
  ensureAuthenticated
};
