/**
 * Creates a new draft email
 */
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Creates a new draft email
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateDraft(args) {
  const { 
    to, 
    cc, 
    bcc, 
    subject, 
    body, 
    importance = 'normal',
    sharedMailbox 
  } = args;
  
  if (!to) {
    return {
      content: [{ 
        type: "text", 
        text: "At least one recipient (to) is required to create a draft."
      }]
    };
  }
  
  if (!subject) {
    return {
      content: [{ 
        type: "text", 
        text: "Subject is required to create a draft."
      }]
    };
  }
  
  if (!body) {
    return {
      content: [{ 
        type: "text", 
        text: "Body content is required to create a draft."
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    // Parse recipients
    const toRecipients = to.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    }));
    
    const ccRecipients = cc ? cc.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    })) : [];
    
    const bccRecipients = bcc ? bcc.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    })) : [];
    
    // Build the draft message
    const draftMessage = {
      subject: subject,
      body: {
        contentType: "Text",
        content: body
      },
      toRecipients: toRecipients,
      importance: importance
    };
    
    // Add CC and BCC if provided
    if (ccRecipients.length > 0) {
      draftMessage.ccRecipients = ccRecipients;
    }
    
    if (bccRecipients.length > 0) {
      draftMessage.bccRecipients = bccRecipients;
    }
    
    // Build the endpoint - handle shared mailbox
    let endpoint = 'me/messages';
    if (sharedMailbox) {
      endpoint = `users/${sharedMailbox}/messages`;
    }
    
    // Create the draft
    const draft = await callGraphAPI(
      accessToken,
      'POST',
      endpoint,
      draftMessage
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Draft email created successfully!\n\nDraft ID: ${draft.id}\nSubject: ${draft.subject}\nTo: ${draft.toRecipients?.map(r => r.emailAddress.address).join(', ')}\n${draft.ccRecipients?.length ? `CC: ${draft.ccRecipients.map(r => r.emailAddress.address).join(', ')}\n` : ''}${draft.bccRecipients?.length ? `BCC: ${draft.bccRecipients.map(r => r.emailAddress.address).join(', ')}\n` : ''}\nYou can edit this draft further or send it when ready.`
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
        text: `Error creating draft: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateDraft;