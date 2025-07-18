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
    
    // Find or create Attachments folder using a more robust approach
    let attachmentsFolderId = null;
    try {
      // Try to access the Attachments folder directly first
      try {
        const attachmentsFolder = await callGraphAPI(accessToken, 'GET', 'me/drive/root:/Attachments');
        attachmentsFolderId = attachmentsFolder.id;
        console.log('Found Attachments folder directly:', attachmentsFolderId);
      } catch (directError) {
        console.log('Direct access failed, trying to list root children...');
        // If direct access fails, list all children and find it
        const rootItems = await callGraphAPI(accessToken, 'GET', 'me/drive/root/children');
        const attachmentsFolder = rootItems.value?.find(item => 
          item.name === 'Attachments' && item.folder !== undefined
        );
        
        if (attachmentsFolder) {
          attachmentsFolderId = attachmentsFolder.id;
          console.log('Found Attachments folder in children:', attachmentsFolderId);
        } else {
          console.log('Attachments folder not found, creating it...');
          // Create the folder
          const newFolder = await callGraphAPI(accessToken, 'POST', 'me/drive/root/children', {
            name: 'Attachments',
            folder: {},
            '@microsoft.graph.conflictBehavior': 'rename'
          });
          attachmentsFolderId = newFolder.id;
          console.log('Created new Attachments folder:', attachmentsFolderId);
        }
      }
    } catch (error) {
      console.error('Error setting up Attachments folder:', error);
      throw new Error(`Failed to access Attachments folder: ${error.message}`);
    }
    
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
            // Upload to Attachments folder
            const uploadPath = `me/drive/items/${attachmentsFolderId}:/${fileName}:/content`;
              
            const uploadResponse = await callGraphAPI(
              accessToken, 
              'PUT', 
              uploadPath,
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