/**
 * Creates a draft reply to an email
 */
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Creates a draft reply to an email
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateReplyDraft(args) {
  const { emailId, body, replyAll = false, sharedMailbox } = args;
  
  if (!emailId) {
    return {
      content: [{ 
        type: "text", 
        text: "Email ID is required to create a reply draft."
      }]
    };
  }
  
  if (!body) {
    return {
      content: [{ 
        type: "text", 
        text: "Reply body content is required."
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    // Build the endpoint - handle shared mailbox
    let endpoint = `me/messages/${emailId}`;
    if (sharedMailbox) {
      endpoint = `users/${sharedMailbox}/messages/${emailId}`;
    }
    
    // First get the original email to extract reply information
    const originalEmail = await callGraphAPI(accessToken, 'GET', endpoint);
    
    // Determine reply endpoint
    let replyEndpoint = `me/messages/${emailId}/`;
    if (sharedMailbox) {
      replyEndpoint = `users/${sharedMailbox}/messages/${emailId}/`;
    }
    
    if (replyAll) {
      replyEndpoint += 'createReplyAll';
    } else {
      replyEndpoint += 'createReply';
    }
    
    // Create the reply draft with the specified body
    const replyDraft = await callGraphAPI(
      accessToken,
      'POST',
      replyEndpoint,
      {
        message: {
          body: {
            contentType: "Text",
            content: body
          }
        }
      }
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Draft reply created successfully!\n\nDraft ID: ${replyDraft.id}\nSubject: ${replyDraft.subject}\nTo: ${replyDraft.toRecipients?.map(r => r.emailAddress.address).join(', ')}\n\nYou can edit this draft further or send it when ready.`
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
        text: `Error creating reply draft: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateReplyDraft;