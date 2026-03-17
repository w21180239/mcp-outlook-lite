import { debug, warn } from '../../utils/logger.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { saveBase64File } from '../../utils/fileOutput.js';
import { safeStringify } from '../../utils/jsonUtils.js';
import { decodeContent as decodeFileContent, formatFileSize } from '../common/fileTypeUtils.js';

function applyDecodedContent(target: Record<string, any>, decodedContent: Record<string, any>) {
  target.content = decodedContent.content;
  target.decodedContentType = decodedContent.type;
  target.encoding = decodedContent.encoding;
  target.contentIncluded = true;
  if (decodedContent.contentBytes) target.contentBytes = decodedContent.contentBytes;
  if (decodedContent.note) target.note = decodedContent.note;
  if (decodedContent.error) target.decodingError = decodedContent.error;
}

// Download attachment
export async function downloadAttachmentTool(authManager: any, args: Record<string, any>) {
  const { messageId, attachmentId, includeContent = false, decodeContent = true } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!attachmentId) {
    return createValidationError('attachmentId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    debug(`Debug: Downloading attachment ${attachmentId} from message ${messageId}`);

    // First, get attachment metadata and type
    const attachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`, {
      select: 'id,name,contentType,size,isInline,lastModifiedDateTime,@odata.type'
    });

    debug(`Debug: Attachment type: ${attachment['@odata.type']}, size: ${attachment.size}, contentType: "${attachment.contentType}"`);

    const attachmentInfo: Record<string, any> = {
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size || 0,
      sizeFormatted: formatFileSize(attachment.size),
      isInline: attachment.isInline || false,
      lastModifiedDateTime: attachment.lastModifiedDateTime,
      attachmentType: attachment['@odata.type']
    };

    if (includeContent) {
      try {
        debug('Debug: Attempting to download attachment content...');
        
        // Try different approaches based on attachment type
        if (attachment['@odata.type'] === '#microsoft.graph.fileAttachment') {
          // Standard file attachment - request with contentBytes
          const fullAttachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`, {
            select: 'id,name,contentType,size,isInline,lastModifiedDateTime,contentBytes,@odata.type'
          });
          
          if (fullAttachment.contentBytes) {
            if (decodeContent) {
              // Decode the Base64 content appropriately
              const decodedContent = await decodeFileContent(
                fullAttachment.contentBytes, 
                attachment.contentType, 
                attachment.name
              );
              
              applyDecodedContent(attachmentInfo, decodedContent);

              debug(`Debug: Successfully downloaded and decoded content (type: ${decodedContent.type}, size: ${decodedContent.size} bytes)`);
            } else {
              // Return raw Base64 content
              attachmentInfo.contentBytes = fullAttachment.contentBytes;
              attachmentInfo.contentIncluded = true;
              attachmentInfo.encoding = 'base64';
              attachmentInfo.note = 'Raw Base64 content (set decodeContent: true to decode)';
              debug(`Debug: Successfully downloaded raw content (${fullAttachment.contentBytes.length} Base64 characters)`);
            }
          } else {
            attachmentInfo.contentIncluded = false;
            attachmentInfo.contentError = 'No content bytes returned from API';
          }
          
        } else if (attachment['@odata.type'] === '#microsoft.graph.itemAttachment') {
          // Item attachment (embedded message/calendar item)
          const fullAttachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`, {
            expand: 'item'
          });
          
          if (fullAttachment.item) {
            attachmentInfo.itemContent = fullAttachment.item;
            attachmentInfo.contentIncluded = true;
            attachmentInfo.encoding = 'json';
            debug('Debug: Successfully downloaded item attachment content');
          } else {
            attachmentInfo.contentIncluded = false;
            attachmentInfo.contentError = 'No item content available for item attachment';
          }
          
        } else if (attachment['@odata.type'] === '#microsoft.graph.referenceAttachment') {
          // Reference attachment (link to SharePoint/OneDrive)
          const fullAttachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`);
          
          attachmentInfo.sourceUrl = fullAttachment.sourceUrl;
          attachmentInfo.providerType = fullAttachment.providerType;
          attachmentInfo.thumbnailUrl = fullAttachment.thumbnailUrl;
          attachmentInfo.previewUrl = fullAttachment.previewUrl;
          attachmentInfo.permission = fullAttachment.permission;
          attachmentInfo.isFolder = fullAttachment.isFolder;
          attachmentInfo.contentIncluded = false;
          attachmentInfo.contentError = 'Reference attachment - use sourceUrl to access the linked resource';
          debug('Debug: Reference attachment processed, sourceUrl:', fullAttachment.sourceUrl);
          
        } else {
          // Unknown attachment type - try the standard approach
          debug('Debug: Unknown attachment type, trying standard contentBytes approach');
          const fullAttachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`);
          
          if (fullAttachment.contentBytes) {
            if (decodeContent) {
              // Decode the Base64 content for unknown types too
              const decodedContent = await decodeFileContent(
                fullAttachment.contentBytes, 
                attachment.contentType, 
                attachment.name
              );
              
              applyDecodedContent(attachmentInfo, decodedContent);
            } else {
              // Return raw Base64 content
              attachmentInfo.contentBytes = fullAttachment.contentBytes;
              attachmentInfo.contentIncluded = true;
              attachmentInfo.encoding = 'base64';
            }
          } else {
            attachmentInfo.contentIncluded = false;
            attachmentInfo.contentError = `Unsupported attachment type: ${attachment['@odata.type']}`;
            // Include the full response for debugging
            attachmentInfo.debugInfo = {
              availableFields: Object.keys(fullAttachment),
              odataType: attachment['@odata.type']
            };
          }
        }
        
      } catch (contentError) {
        debug('Debug: Error downloading attachment content:', contentError);
        attachmentInfo.contentIncluded = false;
        attachmentInfo.contentError = `Failed to download content: ${contentError.message}`;
        attachmentInfo.errorDetails = {
          statusCode: contentError.statusCode,
          code: contentError.code
        };
      }
    } else {
      attachmentInfo.contentIncluded = false;
      attachmentInfo.contentError = 'Content download not requested (set includeContent: true to download)';
    }

    // Handle large content by saving to file if needed
    const responseText = safeStringify(attachmentInfo, 2);
    const maxMcpResponseSize = 1048576; // 1MB MCP limit
    
    if (responseText.length > maxMcpResponseSize && attachmentInfo.contentBytes) {
      warn(`Response size (${formatFileSize(responseText.length)}) exceeds MCP limit, saving to file...`);
      
      // Save the Base64 content to file
      const fileResult = await saveBase64File(
        attachmentInfo.contentBytes, 
        attachmentInfo.name, 
        attachmentInfo.contentType
      );
      
      if (fileResult.success) {
        // Replace contentBytes with file info
        const fileResponseInfo: Record<string, any> = {
          ...attachmentInfo,
          contentSavedToFile: true,
          fileOutput: fileResult,
          note: `Attachment content saved to file: ${fileResult.filePath}. Use the file path to access the full content.`,
          usage: {
            filePath: 'Use fileOutput.filePath to access the saved file',
            originalContent: 'Large content automatically saved due to MCP 1MB limit',
            decoding: attachmentInfo.encoding === 'parsed' ? 'Content was parsed before saving' : 'Raw file saved as downloaded'
          }
        };
        
        // Remove the large contentBytes from response
        delete fileResponseInfo.contentBytes;
        // Keep parsed content if it's small enough
        if (attachmentInfo.content && typeof attachmentInfo.content === 'object') {
          const contentSize = safeStringify(attachmentInfo.content).length;
          if (contentSize > maxMcpResponseSize / 2) { // If parsed content is also large
            delete fileResponseInfo.content;
            fileResponseInfo.parsedContentTruncated = true;
            fileResponseInfo.note += ' Parsed content also truncated due to size.';
          }
        }
        
        return {
          content: [
            {
              type: 'text',
              text: safeStringify(fileResponseInfo, 2),
            },
          ],
        };
      } else {
        // Fall back to truncation if file saving failed
        const largeContentInfo: Record<string, any> = {
          ...attachmentInfo,
          contentTruncated: true,
          contentSize: responseText.length,
          contentSizeFormatted: formatFileSize(responseText.length),
          mcpLimitExceeded: true,
          fileSaveError: fileResult.error,
          note: `Response size (${formatFileSize(responseText.length)}) exceeds MCP limit. File save failed: ${fileResult.error}`,
          alternatives: {
            suggestion: 'Use decodeContent: true to get parsed text content instead of raw Base64',
            rawAccess: 'Content is available but too large to return in MCP response'
          }
        };
        
        delete largeContentInfo.contentBytes;
        delete largeContentInfo.content;
        
        return {
          content: [
            {
              type: 'text',
              text: safeStringify(largeContentInfo, 2),
            },
          ],
        };
      }
    }

    return {
      content: [
        {
          type: 'text',
          text: responseText,
        },
      ],
    };
  } catch (error) {
    debug('Debug: Error in downloadAttachmentTool:', error);
    return convertErrorToToolError(error, 'Failed to download attachment');
  }
}
