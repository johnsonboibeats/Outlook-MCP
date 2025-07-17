/**
 * Gets information about email attachments
 */
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Gets attachment information from an email
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetAttachments(args) {
  const { emailId, sharedMailbox } = args;
  
  if (!emailId) {
    return {
      content: [{ 
        type: "text", 
        text: "Email ID is required to get attachments."
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    // Build the endpoint - handle shared mailbox
    let endpoint = `me/messages/${emailId}/attachments`;
    if (sharedMailbox) {
      endpoint = `users/${sharedMailbox}/messages/${emailId}/attachments`;
    }
    
    // Get attachments from the email
    const response = await callGraphAPI(accessToken, 'GET', endpoint);
    const attachments = response.value || [];
    
    if (attachments.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "No attachments found in this email."
        }]
      };
    }
    
    // Download each attachment
    const downloadedAttachments = [];
    
    for (const attachment of attachments) {
      try {
        // Get the full attachment content
        const attachmentEndpoint = `${endpoint}/${attachment.id}`;
        const fullAttachment = await callGraphAPI(accessToken, 'GET', attachmentEndpoint);
        
        const sizeKB = Math.round(attachment.size / 1024);
        downloadedAttachments.push({
          name: attachment.name,
          contentType: attachment.contentType,
          size: sizeKB,
          id: attachment.id,
          content: fullAttachment.contentBytes // Base64 encoded content
        });
      } catch (error) {
        console.error(`Failed to download attachment ${attachment.name}:`, error);
        downloadedAttachments.push({
          name: attachment.name,
          contentType: attachment.contentType,
          size: Math.round(attachment.size / 1024),
          id: attachment.id,
          error: `Failed to download: ${error.message}`
        });
      }
    }
    
    // Format results
    const attachmentInfo = downloadedAttachments.map(attachment => {
      if (attachment.error) {
        return `📎 ${attachment.name}\n   Type: ${attachment.contentType}\n   Size: ${attachment.size} KB\n   Status: ❌ ${attachment.error}`;
      } else {
        const contentPreview = attachment.content ? `✅ Downloaded (${attachment.content.length} base64 chars)` : '❌ No content';
        return `📎 ${attachment.name}\n   Type: ${attachment.contentType}\n   Size: ${attachment.size} KB\n   Status: ${contentPreview}\n   Content: ${attachment.content || 'N/A'}`;
      }
    });
    
    return {
      content: [{ 
        type: "text", 
        text: `Downloaded ${attachments.length} attachment(s):\n\n${attachmentInfo.join('\n\n')}`
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
        text: `Error getting attachments: ${error.message}`
      }]
    };
  }
}

module.exports = handleGetAttachments;