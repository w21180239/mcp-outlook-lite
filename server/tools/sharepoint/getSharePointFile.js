/**
 * SharePoint File Access Tool
 * 
 * Fetches files from SharePoint using the same authenticated session as Outlook.
 * Handles various SharePoint URL formats and sharing links.
 */

import { debug } from '../../utils/logger.js';
import { convertErrorToToolError, createValidationError, createToolError } from '../../utils/mcpErrorResponse.js';
import { handleLargeContent, saveBase64File } from '../../utils/fileOutput.js';
import { safeStringify, createSafeResponse } from '../../utils/jsonUtils.js';
import { decodeContent } from '../common/fileTypeUtils.js';

/**
 * Enhanced SharePoint URL parser with comprehensive pattern matching
 * @param {string} sharePointUrl - The SharePoint URL from the email
 * @returns {object} Parsed URL components
 */
function parseSharePointUrl(sharePointUrl) {
  try {
    debug(`Debug: Parsing SharePoint URL: ${sharePointUrl}`);
    const url = new URL(sharePointUrl);
    const hostname = url.hostname.toLowerCase();
    const pathname = url.pathname;
    const searchParams = Object.fromEntries(url.searchParams);

    debug(`Debug: URL components - hostname: ${hostname}, pathname: ${pathname}`);
    debug(`Debug: Search params:`, searchParams);

    // Check if it's a SharePoint domain
    if (!hostname.includes('sharepoint.com')) {
      throw new Error(`Not a SharePoint URL: ${hostname}`);
    }

    // Pattern 1: SharePoint sharing links with format /:x:/r/personal/ or /:w:/r/sites/
    const sharingLinkPattern = /^\/(:[a-z]:)?\/([gr])\/(.+)$/i;
    const sharingMatch = pathname.match(sharingLinkPattern);

    if (sharingMatch) {
      const [, docType, accessType, resourcePath] = sharingMatch;
      debug(`Debug: Detected sharing link - docType: ${docType}, accessType: ${accessType}, resourcePath: ${resourcePath}`);

      return {
        type: 'sharing_link',
        hostname,
        originalUrl: sharePointUrl,
        docType: docType || ':x:', // Default to Excel if not specified
        accessType, // 'r' for read, 'g' for guest
        resourcePath,
        searchParams,
        isPersonal: resourcePath.startsWith('personal/'),
        isSite: resourcePath.startsWith('sites/')
      };
    }

    // Pattern 2: Direct OneDrive for Business URLs
    if (pathname.includes('/personal/')) {
      const personalMatch = pathname.match(/\/personal\/([^\/]+)/);
      if (personalMatch) {
        debug(`Debug: Detected OneDrive personal folder: ${personalMatch[1]}`);
        return {
          type: 'onedrive_personal',
          hostname,
          userFolder: personalMatch[1],
          fullPath: pathname,
          searchParams
        };
      }
    }

    // Pattern 3: Team site URLs
    if (pathname.includes('/sites/')) {
      const siteMatch = pathname.match(/\/sites\/([^\/]+)/);
      if (siteMatch) {
        debug(`Debug: Detected team site: ${siteMatch[1]}`);
        return {
          type: 'team_site',
          hostname,
          siteName: siteMatch[1],
          fullPath: pathname,
          searchParams
        };
      }
    }

    // Pattern 4: Check for any sharing parameters
    const hasShareParams = searchParams.d || searchParams.e || searchParams.share || searchParams.guestaccess;
    if (hasShareParams) {
      debug(`Debug: Detected sharing parameters`);
      return {
        type: 'sharing_with_params',
        hostname,
        fullPath: pathname,
        searchParams,
        hasShareParams: true
      };
    }

    // Fallback: Generic SharePoint URL
    debug(`Debug: Falling back to generic SharePoint URL`);
    return {
      type: 'generic_sharepoint',
      hostname,
      fullPath: pathname,
      searchParams
    };

  } catch (error) {
    debug(`Debug: URL parsing failed: ${error.message}`);
    throw new Error(`Invalid SharePoint URL: ${error.message}`);
  }
}

/**
 * Enhanced sharing link resolver with multiple strategies
 * @param {object} graphClient - Authenticated Graph API client
 * @param {string} sharingUrl - SharePoint sharing URL
 * @param {object} urlInfo - Parsed URL information
 * @returns {object} Sharing information
 */
async function resolveSharedFile(graphClient, sharingUrl, urlInfo) {
  const strategies = [
    // Strategy 1: Direct Graph API shares endpoint
    async () => {
      debug(`Debug: Trying Graph API shares endpoint`);
      const encodedUrl = Buffer.from(sharingUrl).toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=/g, '');

      return await graphClient.makeRequest(`/shares/u!${encodedUrl}/driveItem`, {
        select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,@microsoft.graph.downloadUrl,parentReference'
      });
    },

    // Strategy 2: Try to extract drive and item info from URL structure
    async () => {
      debug(`Debug: Trying URL structure parsing`);
      if ((urlInfo.type === 'sharing_link' || urlInfo.type === 'sharing_with_params') && urlInfo.searchParams.d) {
        // The 'd' parameter in SharePoint URLs is typically a shortened reference
        // Let's try different approaches to extract useful information
        const dParam = decodeURIComponent(urlInfo.searchParams.d);
        debug(`Debug: Found 'd' parameter: ${dParam}`);

        // Strategy 2a: Try to use the 'd' parameter as a sharing token with Graph API
        try {
          // Some SharePoint URLs use the 'd' parameter as a sharing token
          const sharingTokenUrl = `${urlInfo.originalUrl.split('?')[0]}?d=${encodeURIComponent(dParam)}`;
          const encodedSharingUrl = Buffer.from(sharingTokenUrl).toString('base64')
            .replace(/\+/g, '-')
            .replace(/\//g, '_')
            .replace(/=/g, '');

          debug(`Debug: Trying 'd' parameter as sharing token`);
          const result = await graphClient.makeRequest(`/shares/u!${encodedSharingUrl}/driveItem`, {
            select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,@microsoft.graph.downloadUrl,parentReference'
          });

          if (result && result.id) {
            debug(`Debug: Successfully resolved using 'd' parameter as sharing token`);
            return result;
          }
        } catch (sharingTokenError) {
          debug(`Debug: 'd' parameter as sharing token failed: ${sharingTokenError.message}`);
        }

        // Strategy 2b: Try to extract file ID patterns from 'd' parameter
        const fileIdPatterns = [
          /([A-Z0-9]{20,})/gi,  // Long alphanumeric strings
          /([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/gi, // GUIDs
          /([A-Z0-9_-]{15,})/gi // SharePoint-style IDs
        ];

        for (const pattern of fileIdPatterns) {
          const matches = dParam.match(pattern);
          if (matches) {
            for (const potentialFileId of matches) {
              debug(`Debug: Trying potential file ID: ${potentialFileId}`);

              // Try different drive contexts
              const driveContexts = ['me', 'root'];

              for (const driveContext of driveContexts) {
                try {
                  const result = await graphClient.makeRequest(`/drives/${driveContext}/items/${potentialFileId}`, {
                    select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,@microsoft.graph.downloadUrl,parentReference'
                  });
                  debug(`Debug: Successfully found file with ID ${potentialFileId} in drive context: ${driveContext}`);
                  return result;
                } catch (driveError) {
                  debug(`Debug: Drive context ${driveContext} with ID ${potentialFileId} failed: ${driveError.message}`);
                }
              }
            }
          }
        }
      }
      throw new Error('Could not extract usable file information from URL parameters');
    },

    // Strategy 3: Try SharePoint REST API endpoint construction
    async () => {
      debug(`Debug: Trying SharePoint REST API approach`);
      if (urlInfo.type === 'sharing_link' || urlInfo.type === 'sharing_with_params') {
        // Extract site collection and construct direct API call
        const siteUrl = `https://${urlInfo.hostname}`;

        // Try to get the site information first
        const siteResponse = await graphClient.makeRequest(`/sites/${urlInfo.hostname}:/`, {
          select: 'id,displayName,webUrl'
        });

        if (siteResponse && siteResponse.id) {
          debug(`Debug: Found site ID: ${siteResponse.id}`);
          // This would require additional parsing to get to the specific file
          // For now, return site info as fallback
          return {
            id: siteResponse.id,
            name: 'Site Root',
            size: 0,
            isFolder: true,
            webUrl: siteResponse.webUrl,
            note: 'Resolved to site root - specific file resolution needs additional implementation'
          };
        }
      }
      throw new Error('SharePoint REST API approach not applicable');
    }
  ];

  let lastError = null;

  // Try each strategy in sequence
  for (const [index, strategy] of strategies.entries()) {
    try {
      debug(`Debug: Attempting resolution strategy ${index + 1}`);
      const result = await strategy();
      if (result && result.id) {
        debug(`Debug: Strategy ${index + 1} succeeded`);
        return result;
      }
    } catch (error) {
      debug(`Debug: Strategy ${index + 1} failed: ${error.message}`);
      lastError = error;
    }
  }

  // All strategies failed
  throw new Error(`Failed to resolve shared file after trying ${strategies.length} strategies. Last error: ${lastError?.message || 'Unknown error'}`);
}

/**
 * Get file content from SharePoint using Graph API
 * @param {object} graphClient - Authenticated Graph API client
 * @param {string} driveId - Drive ID (site, user, etc.)
 * @param {string} itemId - File item ID
 * @param {boolean} downloadContent - Whether to download file content
 * @returns {object} File information and optionally content
 */
async function getFileFromDrive(graphClient, driveId, itemId, downloadContent = false) {
  try {
    // Get file metadata
    const fileInfo = await graphClient.makeRequest(`/drives/${driveId}/items/${itemId}`, {
      select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,@microsoft.graph.downloadUrl'
    });

    const result = {
      id: fileInfo.id,
      name: fileInfo.name,
      size: fileInfo.size,
      createdDateTime: fileInfo.createdDateTime,
      lastModifiedDateTime: fileInfo.lastModifiedDateTime,
      webUrl: fileInfo.webUrl,
      isFolder: !!fileInfo.folder,
      mimeType: fileInfo.file?.mimeType,
      downloadUrl: fileInfo['@microsoft.graph.downloadUrl']
    };

    // Download content if requested and file is not too large
    if (downloadContent && !fileInfo.folder) {
      const maxSize = 50 * 1024 * 1024; // 50MB limit

      if (fileInfo.size > maxSize) {
        result.contentError = `File too large to download inline (${Math.round(fileInfo.size / 1024 / 1024)}MB > 50MB). Use downloadUrl for direct download.`;
      } else {
        try {
          const contentResponse = await fetch(fileInfo['@microsoft.graph.downloadUrl']);
          if (contentResponse.ok) {
            const contentBuffer = await contentResponse.arrayBuffer();
            const contentBytes = Buffer.from(contentBuffer).toString('base64');
            const contentType = contentResponse.headers.get('content-type') || fileInfo.file?.mimeType;

            // Decode content intelligently based on type
            const decodedContent = await decodeContent(contentBytes, contentType, fileInfo.name);

            // Add decoded content info
            result.content = decodedContent.content;
            result.decodedContentType = decodedContent.type;
            result.encoding = decodedContent.encoding;
            result.contentType = contentType;
            result.contentSize = decodedContent.size;
            result.sizeFormatted = decodedContent.sizeFormatted;

            // Keep raw Base64 for binary files or when needed
            if (decodedContent.contentBytes) {
              result.contentBytes = decodedContent.contentBytes;
            }

            // Add any additional info
            if (decodedContent.note) {
              result.note = decodedContent.note;
            }

            if (decodedContent.error) {
              result.decodingError = decodedContent.error;
            }
          }
        } catch (downloadError) {
          result.contentError = `Failed to download content: ${downloadError.message}`;
        }
      }
    }

    return result;
  } catch (error) {
    console.error('Error getting file from drive:', error);
    throw new Error(`Failed to get file: ${error.message}`);
  }
}

/**
 * Main tool function to get SharePoint file
 * @param {object} authManager - Outlook authentication manager (same session)
 * @param {object} args - Tool arguments
 * @returns {object} MCP tool response
 */
export async function getSharePointFileTool(authManager, args) {
  try {
    // Input validation
    if (!args.sharePointUrl && !args.fileId) {
      return createValidationError('sharePointUrl or fileId', 'Either SharePoint URL or file ID is required');
    }

    // Ensure authentication
    const graphClient = await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    let fileResult;

    if (args.sharePointUrl) {
      console.error(`Fetching SharePoint file from URL: ${args.sharePointUrl}`);

      // Parse the SharePoint URL
      const urlInfo = parseSharePointUrl(args.sharePointUrl);
      console.error('Parsed URL info:', safeStringify(urlInfo, 2));

      // Standard SharePoint sharing URLs from emails should be processed directly
      // They follow the format: /:x:/r/personal/ or /:w:/r/sites/ with d= and e= parameters
      if (urlInfo.type === 'sharing_link' ||
        urlInfo.type === 'sharing_with_params' ||
        urlInfo.hasShareParams ||
        // Check for standard sharing parameters from email links
        (urlInfo.searchParams && (urlInfo.searchParams.d || urlInfo.searchParams.e))) {

        console.error('Debug: Detected standard SharePoint sharing URL from email, attempting resolution');
        try {
          fileResult = await resolveSharedFile(graphApiClient, args.sharePointUrl, urlInfo);
          debug(`Debug: Successfully resolved file: ${fileResult.name}`);

          // Handle content download if requested
          if (args.downloadContent && !fileResult.folder) {
            console.error('Debug: Content download requested for resolved sharing link');
            const maxSize = 50 * 1024 * 1024; // 50MB limit

            if (fileResult.size > maxSize) {
              fileResult.contentError = `File too large to download inline (${Math.round(fileResult.size / 1024 / 1024)}MB > 50MB). Use downloadUrl for direct download.`;
            } else {
              try {
                const downloadUrl = fileResult['@microsoft.graph.downloadUrl'];
                if (!downloadUrl) {
                  console.error('Debug: No download URL available, trying to fetch it');
                  // Try to get download URL using the file ID and parent reference
                  if (fileResult.id && fileResult.parentReference?.driveId) {
                    const freshFileInfo = await graphApiClient.makeRequest(`/drives/${fileResult.parentReference.driveId}/items/${fileResult.id}`, {
                      select: '@microsoft.graph.downloadUrl'
                    });
                    fileResult['@microsoft.graph.downloadUrl'] = freshFileInfo['@microsoft.graph.downloadUrl'];
                  } else {
                    throw new Error('No download URL available and insufficient metadata to fetch it');
                  }
                }

                const actualDownloadUrl = fileResult['@microsoft.graph.downloadUrl'];
                debug(`Debug: Using download URL: ${actualDownloadUrl ? 'URL available' : 'URL missing'}`);

                if (actualDownloadUrl) {
                  const contentResponse = await fetch(actualDownloadUrl);
                  if (contentResponse.ok) {
                    const contentBuffer = await contentResponse.arrayBuffer();
                    const contentBytes = Buffer.from(contentBuffer).toString('base64');
                    const contentType = contentResponse.headers.get('content-type') || fileResult.mimeType;

                    // Decode content intelligently based on type
                    const decodedContent = await decodeContent(contentBytes, contentType, fileResult.name);

                    // Add decoded content info
                    fileResult.content = decodedContent.content;
                    fileResult.decodedContentType = decodedContent.type;
                    fileResult.encoding = decodedContent.encoding;
                    fileResult.contentType = contentType;
                    fileResult.contentSize = decodedContent.size;
                    fileResult.sizeFormatted = decodedContent.sizeFormatted;

                    // Keep raw Base64 for binary files or when needed
                    if (decodedContent.contentBytes) {
                      fileResult.contentBytes = decodedContent.contentBytes;
                    }

                    // Add any additional info
                    if (decodedContent.note) {
                      fileResult.note = decodedContent.note;
                    }

                    if (decodedContent.error) {
                      fileResult.decodingError = decodedContent.error;
                    }

                    debug(`Debug: Successfully downloaded and decoded content (type: ${decodedContent.type}, size: ${decodedContent.size} bytes)`);
                  } else {
                    fileResult.contentError = `Failed to download content: HTTP ${contentResponse.status} ${contentResponse.statusText}`;
                  }
                } else {
                  fileResult.contentError = 'No download URL available for content download';
                }
              } catch (downloadError) {
                debug(`Debug: Content download failed: ${downloadError.message}`);
                fileResult.contentError = `Failed to download content: ${downloadError.message}`;
              }
            }
          }
        } catch (resolveError) {
          debug(`Debug: Resolution failed: ${resolveError.message}`);
          return createToolError(
            `Failed to resolve SharePoint sharing link: ${resolveError.message}`,
            'RESOLUTION_FAILED',
            {
              originalUrl: args.sharePointUrl,
              urlType: urlInfo.type,
              parsedInfo: urlInfo,
              resolutionAttempted: true,
              detailedError: resolveError.message,
              troubleshooting: {
                checkPermissions: 'Ensure you have access to the shared file',
                checkUrl: 'Verify the sharing link is valid and not expired',
                tryFileId: 'If you have the file ID, try using it directly'
              }
            }
          );
        }

      } else if (urlInfo.type === 'onedrive_personal' || urlInfo.type === 'team_site') {
        // For direct site URLs, try to provide more helpful guidance
        return createToolError(
          `Direct ${urlInfo.type} URLs require file-specific sharing links. Please use a sharing link to the specific file.`,
          'DIRECT_SITE_URL_UNSUPPORTED',
          {
            suggestion: 'Right-click the file in SharePoint/OneDrive and select "Copy link" to get a sharing link',
            urlType: urlInfo.type,
            parsedInfo: urlInfo,
            supportedFormats: [
              'https://company.sharepoint.com/:w:/r/sites/...',
              'https://company.sharepoint.com/:x:/g/personal/...',
              'https://company-my.sharepoint.com/:b:/personal/...'
            ]
          }
        );

      } else {
        // Generic SharePoint URL
        return createToolError(
          `Unsupported SharePoint URL format. Please use a file sharing link.`,
          'UNSUPPORTED_URL_FORMAT',
          {
            suggestion: 'Generate a sharing link from SharePoint by right-clicking the file and selecting "Copy link"',
            urlType: urlInfo.type,
            parsedInfo: urlInfo,
            supportedFormats: [
              'File sharing links: https://company.sharepoint.com/:w:/r/...',
              'OneDrive sharing links: https://company-my.sharepoint.com/:x:/g/...'
            ]
          }
        );
      }
    } else if (args.fileId) {
      // Direct file access using Graph API
      const driveId = args.driveId || 'me'; // Default to user's OneDrive
      fileResult = await getFileFromDrive(graphApiClient, driveId, args.fileId, args.downloadContent);
    }

    const response = {
      success: true,
      file: fileResult,
      message: `Successfully retrieved ${fileResult.isFolder ? 'folder' : 'file'}: ${fileResult.name}`,
      usage: {
        downloadUrl: 'Use the downloadUrl for direct file download',
        content: fileResult.content ? 'File content included' : 'Content not downloaded (use downloadContent: true to include)',
        webUrl: 'Use webUrl to view file in SharePoint/OneDrive'
      }
    };

    // Handle large content by saving to file if necessary
    const finalResponse = await handleLargeContent(response, ['file.contentBytes', 'file.content'], {
      filenameSuffix: fileResult.name ? `_${fileResult.name}` : '_sharepoint_file',
      contextInfo: {
        toolName: 'sharepoint_get_file',
        fileName: fileResult.name,
        fileSize: fileResult.size,
        originalUrl: args.sharePointUrl || 'Direct file access'
      }
    });

    if (finalResponse.savedToFile) {
      // Add helpful context when content was saved to file
      finalResponse.file = {
        ...finalResponse.file,
        contentAccessInfo: {
          savedToFile: true,
          reason: 'File content exceeded MCP response size limit (1MB)',
          alternatives: {
            localFile: 'Content saved to local file (see savedFiles)',
            downloadUrl: finalResponse.file.downloadUrl || 'Use downloadUrl for direct download',
            webUrl: finalResponse.file.webUrl || 'Use webUrl to view in SharePoint/OneDrive'
          }
        }
      };
    }

    return createSafeResponse(finalResponse);

  } catch (error) {
    console.error('SharePoint file access error:', error);

    if (error.isError) {
      return error; // Already an MCP error
    }

    return convertErrorToToolError(error, 'SharePoint file access failed');
  }
}

/**
 * Tool to resolve SharePoint sharing links without downloading
 * @param {object} authManager - Authentication manager
 * @param {object} args - Tool arguments
 * @returns {object} MCP tool response
 */
export async function resolveSharePointLinkTool(authManager, args) {
  try {
    if (!args.sharePointUrl) {
      return createValidationError('sharePointUrl', 'SharePoint URL is required');
    }

    const graphApiClient = authManager.getGraphApiClient();

    // Parse the URL first
    const urlInfo = parseSharePointUrl(args.sharePointUrl);
    console.error('Resolve SharePoint link - Parsed URL info:', safeStringify(urlInfo, 2));

    // Resolve the sharing link to get metadata only
    const fileInfo = await resolveSharedFile(graphApiClient, args.sharePointUrl, urlInfo);

    const result = {
      id: fileInfo.id,
      name: fileInfo.name,
      size: fileInfo.size,
      type: fileInfo.folder ? 'folder' : 'file',
      mimeType: fileInfo.file?.mimeType,
      createdDateTime: fileInfo.createdDateTime,
      lastModifiedDateTime: fileInfo.lastModifiedDateTime,
      webUrl: fileInfo.webUrl,
      downloadUrl: fileInfo['@microsoft.graph.downloadUrl'],
      sharing: {
        originalUrl: args.sharePointUrl,
        resolved: true,
        accessible: true
      }
    };

    // Add permissions info if requested
    if (args.includePermissions) {
      try {
        const permissions = await graphApiClient.makeRequest(`/drives/items/${fileInfo.id}/permissions`);
        result.sharing.permissions = permissions.value || [];
      } catch (permError) {
        result.sharing.permissionsError = 'Could not retrieve permissions information';
      }
    }

    return createSafeResponse({
      success: true,
      file: result,
      message: `Successfully resolved ${result.type}: ${result.name}`,
      usage: {
        downloadUrl: 'Use downloadUrl for direct download without re-authentication',
        webUrl: 'Use webUrl to view in browser',
        fileId: 'Use id with outlook_get_sharepoint_file for content download'
      }
    });

  } catch (error) {
    console.error('SharePoint link resolution error:', error);
    return convertErrorToToolError(error, 'SharePoint link resolution failed');
  }
}

/**
 * Tool to list files in a SharePoint site or folder
 * @param {object} authManager - Authentication manager
 * @param {object} args - Tool arguments
 * @returns {object} MCP tool response
 */
export async function listSharePointFilesTool(authManager, args) {
  try {
    const graphApiClient = authManager.getGraphApiClient();

    let listPath;
    if (args.siteId && args.driveId) {
      listPath = `/sites/${args.siteId}/drives/${args.driveId}/root/children`;
    } else if (args.driveId) {
      listPath = `/drives/${args.driveId}/root/children`;
    } else {
      listPath = '/me/drive/root/children'; // Default to user's OneDrive
    }

    if (args.folderId) {
      listPath = `/drives/${args.driveId || 'me'}/items/${args.folderId}/children`;
    }

    const response = await graphApiClient.makeRequest(listPath, {
      select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder',
      top: args.limit || 50,
      orderby: args.orderBy || 'name'
    });

    const files = (response.value || []).map(item => ({
      id: item.id,
      name: item.name,
      size: item.size,
      type: item.folder ? 'folder' : 'file',
      mimeType: item.file?.mimeType,
      createdDateTime: item.createdDateTime,
      lastModifiedDateTime: item.lastModifiedDateTime,
      webUrl: item.webUrl
    }));

    return createSafeResponse({
      success: true,
      files: files,
      count: files.length,
      message: `Found ${files.length} items`
    });

  } catch (error) {
    console.error('SharePoint list files error:', error);
    return convertErrorToToolError(error, 'SharePoint file listing failed');
  }
}
