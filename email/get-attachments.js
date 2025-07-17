/**
 * Gets information about email attachments and saves them to OneDrive
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
    
    // We'll upload directly to /Attachments/ folder - OneDrive will create it if needed
    
    // Download each attachment
    const downloadedAttachments = [];
    
    for (const attachment of attachments) {
      try {
        // Skip only signature images and inline images, but allow legitimate image attachments
        const fileName = attachment.name || `attachment_${attachment.id}`;
        const isSignatureImage = attachment.contentType?.startsWith('image/') && (
                                 attachment.isInline ||          // Microsoft flag for inline content
                                 attachment.contentId ||         // Has content ID (embedded in HTML)
                                 attachment.contentDisposition === 'inline' || // Disposition indicates inline
                                 (attachment.size && attachment.size < 5000)   // Very small images (likely signatures)
                                );
        
        if (isSignatureImage) {
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
          // Upload file to OneDrive Attachments folder
          const fileBuffer = Buffer.from(fullAttachment.contentBytes, 'base64');
          
          try {
            console.log(`Uploading ${fileName} (${fileBuffer.length} bytes) to OneDrive...`);
            const uploadResponse = await callGraphAPI(
              accessToken, 
              'PUT', 
              `me/drive/root:/Attachments/${fileName}:/content`,
              fileBuffer,
              null,
              { 'Content-Type': 'application/octet-stream' }
            );
            
            console.log(`Upload response for ${fileName}:`, JSON.stringify(uploadResponse, null, 2));
            
            const sizeKB = Math.round(attachment.size / 1024);
            downloadedAttachments.push({
              name: fileName,
              contentType: attachment.contentType,
              size: sizeKB,
              id: attachment.id,
              oneDriveUrl: uploadResponse?.webUrl || `https://onedrive.live.com/`,
              status: 'saved'
            });
          } catch (uploadError) {
            console.error(`Failed to upload ${fileName} to OneDrive:`, uploadError);
            console.error(`Upload error stack:`, uploadError.stack);
            downloadedAttachments.push({
              name: fileName,
              contentType: attachment.contentType,
              size: Math.round(attachment.size / 1024),
              id: attachment.id,
              error: `Upload failed: ${uploadError.message}`
            });
          }
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
        return `📎 ${attachment.name}\n   Type: ${attachment.contentType}\n   Size: ${attachment.size} KB\n   Status: ✅ Saved to OneDrive\n   URL: ${attachment.oneDriveUrl}`;
      }
    });
    
    return {
      content: [{ 
        type: "text", 
        text: `Downloaded and saved ${attachments.length} attachment(s) to OneDrive/Attachments:\n\n${attachmentInfo.join('\n\n')}`
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