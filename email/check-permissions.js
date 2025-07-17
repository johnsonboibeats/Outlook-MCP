/**
 * Check mailbox permissions functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Check mailbox permissions handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCheckPermissions(args) {
  const targetMailbox = args.mailbox || 'info@paperbarkculturaltours.com.au';
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    let resultText = `Checking permissions for mailbox: ${targetMailbox}\n\n`;
    
    // 1. Check if we can access the target mailbox directly
    try {
      const directAccess = await callGraphAPI(
        accessToken,
        'GET',
        `users/${targetMailbox}/mailFolders/inbox?$select=id,displayName`
      );
      
      if (directAccess) {
        resultText += '✅ Direct access to mailbox: SUCCESS\n';
        resultText += `   Inbox ID: ${directAccess.id}\n`;
      }
    } catch (error) {
      resultText += '❌ Direct access to mailbox: FAILED\n';
      resultText += `   Error: ${error.message}\n`;
    }
    
    // 2. Check current user's delegated permissions
    try {
      const me = await callGraphAPI(accessToken, 'GET', 'me?$select=displayName,mail,userPrincipalName');
      resultText += `\nCurrent user: ${me.displayName} (${me.mail})\n`;
    } catch (error) {
      resultText += `\nCurrent user info: Error - ${error.message}\n`;
    }
    
    // 3. Try to get mailbox folders to test read permissions
    try {
      const folders = await callGraphAPI(
        accessToken,
        'GET',
        `users/${targetMailbox}/mailFolders?$select=id,displayName&$top=5`
      );
      
      if (folders.value && folders.value.length > 0) {
        resultText += '\n✅ Can read mailbox folders:\n';
        folders.value.forEach(folder => {
          resultText += `   - ${folder.displayName} (${folder.id})\n`;
        });
      }
    } catch (error) {
      resultText += '\n❌ Cannot read mailbox folders\n';
      resultText += `   Error: ${error.message}\n`;
    }
    
    // 4. Try to get messages to test message access
    try {
      const messages = await callGraphAPI(
        accessToken,
        'GET',
        `users/${targetMailbox}/messages?$select=id,subject&$top=1`
      );
      
      if (messages.value) {
        resultText += '\n✅ Can read messages: SUCCESS\n';
        resultText += `   Found ${messages.value.length} accessible messages\n`;
      }
    } catch (error) {
      resultText += '\n❌ Cannot read messages\n';
      resultText += `   Error: ${error.message}\n`;
    }
    
    // 5. Check send permissions by testing send-as capability (dry run)
    try {
      // We can't do a real dry run, but we can check if the endpoint is accessible
      const sendCheck = await callGraphAPI(
        accessToken,
        'GET',
        `users/${targetMailbox}/mailFolders/sentitems?$select=id,displayName`
      );
      
      if (sendCheck) {
        resultText += '\n✅ Send-as permissions: Likely available (can access sent items)\n';
      }
    } catch (error) {
      resultText += '\n❌ Send-as permissions: Likely NOT available\n';
      resultText += `   Error: ${error.message}\n`;
    }
    
    // 6. Recommendations
    resultText += '\n--- RECOMMENDATIONS ---\n';
    resultText += 'If you see ❌ errors above, you need to:\n';
    resultText += '1. Contact your IT administrator\n';
    resultText += '2. Request "Full Access" permissions to the shared mailbox\n';
    resultText += '3. Request "Send As" or "Send on Behalf" permissions\n';
    resultText += '4. Verify the mailbox address is correct\n';
    resultText += '5. Ensure the shared mailbox is properly configured\n';
    
    return {
      content: [{ 
        type: "text", 
        text: resultText
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error checking permissions: ${error.message}`
      }]
    };
  }
}

module.exports = handleCheckPermissions;