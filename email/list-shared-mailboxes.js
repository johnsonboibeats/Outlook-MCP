/**
 * List shared mailboxes functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List shared mailboxes handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListSharedMailboxes(args) {
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    let resultText = 'Available mailboxes and shared mailboxes:\n\n';
    
    // Get current user info
    const me = await callGraphAPI(accessToken, 'GET', 'me?$select=displayName,mail,userPrincipalName');
    resultText += `Personal mailbox: ${me.displayName} (${me.mail})\n\n`;
    
    // Test known shared mailboxes by trying to access them
    const knownSharedMailboxes = [
      'info@paperbarkculturaltours.com.au',
      'admin@paperbarkculturaltours.com.au',
      'contact@paperbarkculturaltours.com.au',
      'support@paperbarkculturaltours.com.au'
    ];
    
    const accessibleSharedMailboxes = [];
    
    for (const mailbox of knownSharedMailboxes) {
      try {
        // Try to access the mailbox by getting its inbox
        await callGraphAPI(
          accessToken,
          'GET',
          `users/${mailbox}/mailFolders/inbox?$select=id,displayName`
        );
        accessibleSharedMailboxes.push(mailbox);
      } catch (error) {
        // If we can't access it, skip it silently
        console.error(`Cannot access ${mailbox}:`, error.message);
      }
    }
    
    // Also try to discover shared mailboxes through different methods
    try {
      // Method 1: Try to get delegated mailboxes (if available)
      const delegatedResponse = await callGraphAPI(
        accessToken,
        'GET',
        'me/mailboxSettings/delegatedMailboxes'
      );
      
      if (delegatedResponse.value && delegatedResponse.value.length > 0) {
        delegatedResponse.value.forEach(delegated => {
          if (!accessibleSharedMailboxes.includes(delegated.emailAddress)) {
            accessibleSharedMailboxes.push(delegated.emailAddress);
          }
        });
      }
    } catch (error) {
      // This API might not be available, continue
      console.error('Could not get delegated mailboxes:', error.message);
    }
    
    // Method 2: Try to get organization users (limited approach)
    try {
      const orgResponse = await callGraphAPI(
        accessToken,
        'GET',
        'users?$filter=accountEnabled eq true&$select=displayName,mail&$top=20'
      );
      
      if (orgResponse.value) {
        for (const user of orgResponse.value) {
          if (user.mail && user.mail !== me.mail) {
            // Test if we can access this user's mailbox
            try {
              await callGraphAPI(
                accessToken,
                'GET',
                `users/${user.mail}/mailFolders/inbox?$select=id`
              );
              if (!accessibleSharedMailboxes.includes(user.mail)) {
                accessibleSharedMailboxes.push(`${user.mail} (${user.displayName})`);
              }
            } catch (error) {
              // Can't access this mailbox, skip
            }
          }
        }
      }
    } catch (error) {
      console.error('Could not query organization users:', error.message);
    }
    
    if (accessibleSharedMailboxes.length > 0) {
      resultText += 'Accessible shared mailboxes:\n';
      accessibleSharedMailboxes.forEach((mailbox, index) => {
        resultText += `${index + 1}. ${mailbox}\n`;
      });
      resultText += '\nTo use a shared mailbox, add the "sharedMailbox" parameter to your email commands.\n';
      resultText += 'Example: list-emails with sharedMailbox: "info@paperbarkculturaltours.com.au"\n';
    } else {
      resultText += 'No additional shared mailboxes found or accessible.\n';
      resultText += 'You may need to:\n';
      resultText += '1. Contact your IT administrator to grant access\n';
      resultText += '2. Verify the shared mailbox address\n';
      resultText += '3. Check if you have delegate permissions\n';
    }
    
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
        text: `Error listing shared mailboxes: ${error.message}`
      }]
    };
  }
}

module.exports = handleListSharedMailboxes;