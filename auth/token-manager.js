/**
 * Token management for Microsoft Graph API authentication
 */
const fs = require('fs');
const https = require('https');
const config = require('../config');

// Global variable to store tokens
let cachedTokens = null;

/**
 * Loads authentication tokens from the token file or global storage
 * @returns {object|null} - The loaded tokens or null if not available
 */
function loadTokenCache() {
  try {
    // For Railway deployment, check global storage first
    if (global.outlookTokens) {
      console.error('[DEBUG] Loading tokens from global storage (Railway mode)');
      cachedTokens = global.outlookTokens;
      return cachedTokens;
    }
    
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    console.error(`[DEBUG] Attempting to load tokens from: ${tokenPath}`);
    console.error(`[DEBUG] HOME directory: ${process.env.HOME}`);
    console.error(`[DEBUG] Full resolved path: ${tokenPath}`);
    
    // Log file existence and details
    if (!fs.existsSync(tokenPath)) {
      console.error('[DEBUG] Token file does not exist');
      return null;
    }
    
    const stats = fs.statSync(tokenPath);
    console.error(`[DEBUG] Token file stats:
      Size: ${stats.size} bytes
      Created: ${stats.birthtime}
      Modified: ${stats.mtime}`);
    
    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    console.error('[DEBUG] Token file contents length:', tokenData.length);
    console.error('[DEBUG] Token file first 200 characters:', tokenData.slice(0, 200));
    
    try {
      const tokens = JSON.parse(tokenData);
      console.error('[DEBUG] Parsed tokens keys:', Object.keys(tokens));
      
      // Log each key's value to see what's present
      Object.keys(tokens).forEach(key => {
        console.error(`[DEBUG] ${key}: ${typeof tokens[key]}`);
      });
      
      // Check for access token presence
      if (!tokens.access_token) {
        console.error('[DEBUG] No access_token found in tokens');
        return null;
      }
      
      // Check token expiration
      const now = Date.now();
      const expiresAt = tokens.expires_at || 0;
      
      console.error(`[DEBUG] Current time: ${now}`);
      console.error(`[DEBUG] Token expires at: ${expiresAt}`);
      
      if (now > expiresAt) {
        console.error('[DEBUG] Token has expired, attempting refresh...');
        // Try to refresh the token
        if (tokens.refresh_token) {
          return refreshAccessToken(tokens.refresh_token);
        }
        return null;
      }
      
      // Update the cache
      cachedTokens = tokens;
      return tokens;
    } catch (parseError) {
      console.error('[DEBUG] Error parsing token JSON:', parseError);
      return null;
    }
  } catch (error) {
    console.error('[DEBUG] Error loading token cache:', error);
    return null;
  }
}

/**
 * Saves authentication tokens to the token file
 * @param {object} tokens - The tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveTokenCache(tokens) {
  try {
    // Always save to global storage for Railway deployment
    global.outlookTokens = tokens;
    cachedTokens = tokens;
    console.error('[DEBUG] Tokens saved to global storage (Railway mode)');
    
    // Also try to save to file if possible
    try {
      const tokenPath = config.AUTH_CONFIG.tokenStorePath;
      console.error(`Saving tokens to: ${tokenPath}`);
      
      fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
      console.error('Tokens also saved to file');
    } catch (fileError) {
      console.error('Failed to save tokens to file (using global storage only):', fileError.message);
    }
    
    return true;
  } catch (error) {
    console.error('Error saving token cache:', error);
    return false;
  }
}

/**
 * Refreshes the access token using the refresh token
 * @param {string} refreshToken - The refresh token
 * @returns {Promise<object|null>} - The new tokens or null if refresh failed
 */
async function refreshAccessToken(refreshToken) {
  try {
    console.error('[DEBUG] Attempting to refresh access token...');
    
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

    return new Promise((resolve, reject) => {
      const req = https.request(options, (res) => {
        let data = '';
        
        res.on('data', (chunk) => {
          data += chunk;
        });
        
        res.on('end', () => {
          try {
            const response = JSON.parse(data);
            
            if (res.statusCode === 200 && response.access_token) {
              console.error('[DEBUG] Token refresh successful');
              
              // Calculate expiration time
              const expiresIn = response.expires_in || 3600;
              const expiresAt = Date.now() + (expiresIn * 1000);
              
              const newTokens = {
                ...response,
                expires_at: expiresAt
              };
              
              // Save the new tokens
              saveTokenCache(newTokens);
              resolve(newTokens);
            } else {
              console.error('[DEBUG] Token refresh failed:', response);
              resolve(null);
            }
          } catch (error) {
            console.error('[DEBUG] Error parsing refresh response:', error);
            resolve(null);
          }
        });
      });
      
      req.on('error', (error) => {
        console.error('[DEBUG] Network error during token refresh:', error);
        resolve(null);
      });
      
      req.write(tokenData.toString());
      req.end();
    });
  } catch (error) {
    console.error('[DEBUG] Error in refreshAccessToken:', error);
    return null;
  }
}

/**
 * Gets the current access token, loading from cache if necessary
 * @returns {string|null} - The access token or null if not available
 */
function getAccessToken() {
  if (cachedTokens && cachedTokens.access_token) {
    // Check if token is close to expiring (within 5 minutes)
    const now = Date.now();
    const expiresAt = cachedTokens.expires_at || 0;
    const fiveMinutes = 5 * 60 * 1000;
    
    if (now + fiveMinutes > expiresAt && cachedTokens.refresh_token) {
      console.error('[DEBUG] Token expires soon, proactively refreshing...');
      // Don't wait for the async refresh, return current token but trigger refresh
      refreshAccessToken(cachedTokens.refresh_token);
    }
    
    return cachedTokens.access_token;
  }
  
  const tokens = loadTokenCache();
  return tokens ? tokens.access_token : null;
}

/**
 * Gets the current access token with automatic refresh if needed
 * @returns {Promise<string|null>} - The access token or null if not available
 */
async function getAccessTokenAsync() {
  const tokens = loadTokenCache();
  if (!tokens) {
    return null;
  }
  
  // Check if token is expired or close to expiring
  const now = Date.now();
  const expiresAt = tokens.expires_at || 0;
  const fiveMinutes = 5 * 60 * 1000;
  
  if (now + fiveMinutes > expiresAt && tokens.refresh_token) {
    console.error('[DEBUG] Token expired or expires soon, refreshing...');
    const newTokens = await refreshAccessToken(tokens.refresh_token);
    return newTokens ? newTokens.access_token : null;
  }
  
  return tokens.access_token;
}

/**
 * Creates a test access token for use in test mode
 * @returns {object} - The test tokens
 */
function createTestTokens() {
  const testTokens = {
    access_token: "test_access_token_" + Date.now(),
    refresh_token: "test_refresh_token_" + Date.now(),
    expires_at: Date.now() + (3600 * 1000) // 1 hour
  };
  
  saveTokenCache(testTokens);
  return testTokens;
}

module.exports = {
  loadTokenCache,
  saveTokenCache,
  getAccessToken,
  getAccessTokenAsync,
  refreshAccessToken,
  createTestTokens
};
