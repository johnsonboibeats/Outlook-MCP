/**
 * Read email functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Read email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleReadEmail(args) {
  const emailId = args.id;
  const sharedMailbox = args.sharedMailbox; // Add support for shared mailbox
  
  if (!emailId) {
    return {
      content: [{ 
        type: "text", 
        text: "Email ID is required."
      }]
    };
  }
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Build API endpoint - use shared mailbox if specified
    const baseEndpoint = sharedMailbox ? `users/${sharedMailbox}` : 'me';
    const endpoint = `${baseEndpoint}/messages/${encodeURIComponent(emailId)}`;
    const queryParams = {
      $select: config.FIELDS.email.detail
    };
    
    try {
      const email = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);
      
      if (!email) {
        const mailboxText = sharedMailbox ? ` in shared mailbox ${sharedMailbox}` : '';
        return {
          content: [
            {
              type: "text",
              text: `Email with ID ${emailId} not found${mailboxText}.`
            }
          ]
        };
      }
      
      // Format sender, recipients, etc.
      const sender = email.from ? `${email.from.emailAddress.name} (${email.from.emailAddress.address})` : 'Unknown';
      const to = email.toRecipients ? email.toRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const cc = email.ccRecipients && email.ccRecipients.length > 0 ? email.ccRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const bcc = email.bccRecipients && email.bccRecipients.length > 0 ? email.bccRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const date = new Date(email.receivedDateTime).toLocaleString('en-AU', { 
        timeZone: 'Australia/Brisbane',
        year: 'numeric',
        month: '2-digit', 
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true
      });
      
      // Extract body content
      let body = '';
      if (email.body) {
        body = email.body.contentType === 'html' ? 
          // Simple HTML-to-text conversion for HTML bodies
          email.body.content.replace(/<[^>]*>/g, '') : 
          email.body.content;
      } else {
        body = email.bodyPreview || 'No content';
      }
      
      // Format the email
      const formattedEmail = `From: ${sender}
To: ${to}
${cc !== 'None' ? `CC: ${cc}\n` : ''}${bcc !== 'None' ? `BCC: ${bcc}\n` : ''}Subject: ${email.subject}
Date: ${date}
Importance: ${email.importance || 'normal'}
Has Attachments: ${email.hasAttachments ? 'Yes' : 'No'}

${body}`;
      
      return {
        content: [
          {
            type: "text",
            text: formattedEmail
          }
        ]
      };
    } catch (error) {
      console.error(`Error reading email: ${error.message}`);
      
      // Improved error handling with more specific messages
      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return {
          content: [
            {
              type: "text",
              text: `The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID.`
            }
          ]
        };
      } else {
        return {
          content: [
            {
              type: "text",
              text: `Failed to read email: ${error.message}`
            }
          ]
        };
      }
    }
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
        text: `Error accessing email: ${error.message}`
      }]
    };
  }
}

module.exports = handleReadEmail;
