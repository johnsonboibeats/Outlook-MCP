/**
 * List emails functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEmails(args) {
  const folder = args.folder || "inbox";
  const count = Math.min(args.count || 10, config.PAGINATION?.maxPageSize || 100);
  const sharedMailbox = args.sharedMailbox; // New parameter for shared mailbox
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Build API endpoint - use shared mailbox if specified
    let baseEndpoint = sharedMailbox ? `users/${sharedMailbox}` : 'me';
    let endpoint = `${baseEndpoint}/messages`;
    
    if (folder.toLowerCase() !== 'inbox') {
      // Get folder ID first if not inbox
      const folderResponse = await callGraphAPI(
        accessToken, 
        'GET', 
        `${baseEndpoint}/mailFolders?$filter=displayName eq '${folder}'`
      );
      
      if (folderResponse.value && folderResponse.value.length > 0) {
        endpoint = `${baseEndpoint}/mailFolders/${folderResponse.value[0].id}/messages`;
      }
    }
    
    // Add query parameters
    const queryParams = {
      $top: count,
      $orderby: 'receivedDateTime desc',
      $select: config.FIELDS.email.list
    };
    
    // Make API call
    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);
    
    if (!response.value || response.value.length === 0) {
      const mailboxText = sharedMailbox ? ` in shared mailbox ${sharedMailbox}` : '';
      return {
        content: [{ 
          type: "text", 
          text: `No emails found in ${folder}${mailboxText}.`
        }]
      };
    }
    
    // Format results
    const emailList = response.value.map((email, index) => {
      const sender = email.from ? email.from.emailAddress : { name: 'Unknown', address: 'unknown' };
      const date = new Date(email.receivedDateTime).toLocaleString();
      const readStatus = email.isRead ? '' : '[UNREAD] ';
      
      return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
    }).join("\n");
    
    const mailboxText = sharedMailbox ? ` in shared mailbox ${sharedMailbox}` : '';
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} emails in ${folder}${mailboxText}:\n\n${emailList}`
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
        text: `Error listing emails: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEmails;
