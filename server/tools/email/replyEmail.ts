import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

// Reply to an email
export async function replyToEmailTool(authManager: any, args: Record<string, any>) {
  const { messageId, body, bodyType = 'text', comment = '', preserveUserStyling = true } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!body && !comment) {
    return createValidationError('body/comment', 'Either body or comment is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const replyPayload: Record<string, any> = {};

    // Use body or comment as the reply message text
    const replyText = body || comment;
    if (replyText) {
      if (preserveUserStyling) {
        const styledBody = await applyUserStyling(graphApiClient, replyText, bodyType);
        replyPayload.message = {
          body: {
            contentType: styledBody.type === 'html' ? 'HTML' : 'Text',
            content: styledBody.content,
          },
        };
      } else {
        replyPayload.message = {
          body: {
            contentType: bodyType === 'html' ? 'HTML' : 'Text',
            content: replyText,
          },
        };
      }
    }

    const draft = await graphApiClient.postWithRetry(`/me/messages/${messageId}/createReply`, replyPayload);

    return {
      content: [
        {
          type: 'text',
          text: `Reply draft created successfully (NOT sent). Draft ID: ${draft.id || 'N/A'}. Open Outlook to review and send.`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to reply to email');
  }
}

// Reply all to an email
export async function replyAllTool(authManager: any, args: Record<string, any>) {
  const { messageId, body, bodyType = 'text', comment = '', preserveUserStyling = true } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!body && !comment) {
    return createValidationError('body/comment', 'Either body or comment is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const replyPayload: Record<string, any> = {};

    // Use body or comment as the reply message text
    const replyText = body || comment;
    if (replyText) {
      if (preserveUserStyling) {
        const styledBody = await applyUserStyling(graphApiClient, replyText, bodyType);
        replyPayload.message = {
          body: {
            contentType: styledBody.type === 'html' ? 'HTML' : 'Text',
            content: styledBody.content,
          },
        };
      } else {
        replyPayload.message = {
          body: {
            contentType: bodyType === 'html' ? 'HTML' : 'Text',
            content: replyText,
          },
        };
      }
    }

    const draft = await graphApiClient.postWithRetry(`/me/messages/${messageId}/createReplyAll`, replyPayload);

    return {
      content: [
        {
          type: 'text',
          text: `Reply-all draft created successfully (NOT sent). Draft ID: ${draft.id || 'N/A'}. Open Outlook to review and send.`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to reply all to email');
  }
}