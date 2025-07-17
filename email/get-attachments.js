/**
 * Gets information about email attachments and saves them to disk
 */
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const fs = require('fs');
const path = require('path');
const os = require('os');

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
    
    // Use user's Downloads folder
    const downloadsDir = path.join(os.homedir(), 'Downloads');
    if (!fs.existsSync(downloadsDir)) {
      fs.mkdirSync(downloadsDir, { recursive: true });
    }
    
    // Download each attachment
    const downloadedAttachments = [];
    
    for (const attachment of attachments) {
      try {
        // Skip signature images and inline images based on reliable indicators
        const fileName = attachment.name || `attachment_${attachment.id}`;
        const isInlineImage = (attachment.isInline ||          // Microsoft flag for inline content
                              attachment.contentId ||          // Has content ID (embedded in HTML)
                              attachment.contentDisposition === 'inline') && // Disposition indicates inline
                              attachment.contentType?.startsWith('image/'); // Only apply to images, not documents
        
        if (isInlineImage) {
          downloadedAttachments.push({
            name: fileName,
            contentType: attachment.contentType,
            size: Math.round(attachment.size / 1024),
            id: attachment.id,
            status: 'skipped (signature/inline image)'
          });
          continue;
        }
        
        // Get the full attachment content
        const attachmentEndpoint = `${endpoint}/${attachment.id}`;
        const fullAttachment = await callGraphAPI(accessToken, 'GET', attachmentEndpoint);
        
        if (fullAttachment.contentBytes) {
          // Save file to disk
          const filePath = path.join(downloadsDir, fileName);
          const fileBuffer = Buffer.from(fullAttachment.contentBytes, 'base64');
          
          fs.writeFileSync(filePath, fileBuffer);
          
          const sizeKB = Math.round(attachment.size / 1024);
          downloadedAttachments.push({
            name: fileName,
            contentType: attachment.contentType,
            size: sizeKB,
            id: attachment.id,
            filePath: filePath,
            status: 'saved'
          });
        } else {
          downloadedAttachments.push({
            name: attachment.name,
            contentType: attachment.contentType,
            size: Math.round(attachment.size / 1024),
            id: attachment.id,
            error: 'No content received'
          });
        }
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
      } else if (attachment.status === 'skipped (signature/inline image)') {
        return `📎 ${attachment.name}\n   Type: ${attachment.contentType}\n   Size: ${attachment.size} KB\n   Status: ⏭️ Skipped (signature/inline image)`;
      } else {
        return `📎 ${attachment.name}\n   Type: ${attachment.contentType}\n   Size: ${attachment.size} KB\n   Status: ✅ Saved to disk\n   Path: ${attachment.filePath}`;
      }
    });
    
    return {
      content: [{ 
        type: "text", 
        text: `Downloaded and saved ${attachments.length} attachment(s) to ~/Downloads:\n\n${attachmentInfo.join('\n\n')}`
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