import type { ToolHandler } from '../types.js';
import * as tools from './index.js';

/**
 * Registry mapping MCP tool names to handler functions.
 */
const toolRegistry: Record<string, ToolHandler> = {
  // Email tools
  'outlook_list_emails': tools.listEmailsTool,
  'outlook_send_email': tools.sendEmailTool,
  'outlook_get_email': tools.getEmailTool,
  'outlook_search_emails': tools.searchEmailsTool,
  'outlook_create_draft': tools.createDraftTool,
  'outlook_reply_to_email': tools.replyToEmailTool,
  'outlook_reply_all': tools.replyAllTool,
  'outlook_forward_email': tools.forwardEmailTool,
  'outlook_delete_email': tools.deleteEmailTool,
  'outlook_move_email': tools.moveEmailTool,
  'outlook_mark_as_read': tools.markAsReadTool,
  'outlook_flag_email': tools.flagEmailTool,
  'outlook_categorize_email': tools.categorizeEmailTool,
  'outlook_archive_email': tools.archiveEmailTool,
  'outlook_batch_process_emails': tools.batchProcessEmailsTool,
  // Calendar tools
  'outlook_list_events': tools.listEventsTool,
  'outlook_create_event': tools.createEventTool,
  'outlook_get_event': tools.getEventTool,
  'outlook_update_event': tools.updateEventTool,
  'outlook_delete_event': tools.deleteEventTool,
  'outlook_respond_to_invite': tools.respondToInviteTool,
  'outlook_validate_event_datetimes': tools.validateEventDateTimesTool,
  'outlook_create_recurring_event': tools.createRecurringEventTool,
  'outlook_find_meeting_times': tools.findMeetingTimesTool,
  'outlook_check_availability': tools.checkAvailabilityTool,
  'outlook_schedule_online_meeting': tools.scheduleOnlineMeetingTool,
  'outlook_list_calendars': tools.listCalendarsTool,
  'outlook_get_calendar_view': tools.getCalendarViewTool,
  'outlook_get_busy_times': tools.getBusyTimesTool,
  'outlook_build_recurrence_pattern': tools.buildRecurrencePatternTool,
  'outlook_create_recurrence_helper': tools.createRecurrenceHelperTool,
  'outlook_check_calendar_permissions': tools.checkCalendarPermissionsTool,
  // Folder tools
  'outlook_list_folders': tools.listFoldersTool,
  'outlook_create_folder': tools.createFolderTool,
  'outlook_rename_folder': tools.renameFolderTool,
  'outlook_get_folder_stats': tools.getFolderStatsTool,
  // Attachment tools
  'outlook_list_attachments': tools.listAttachmentsTool,
  'outlook_download_attachment': tools.downloadAttachmentTool,
  'outlook_add_attachment': tools.addAttachmentTool,
  'outlook_scan_attachments': tools.scanAttachmentsTool,
  // SharePoint tools
  'outlook_get_sharepoint_file': tools.getSharePointFileTool,
  'outlook_list_sharepoint_files': tools.listSharePointFilesTool,
  'outlook_resolve_sharepoint_link': tools.resolveSharePointLinkTool,
  // Rules tools
  'outlook_list_rules': tools.listRulesTool,
  'outlook_create_rule': tools.createRuleTool,
  'outlook_delete_rule': tools.deleteRuleTool,
};

/**
 * Get the handler function for a tool name.
 */
export function getToolHandler(toolName: string): ToolHandler | null {
  return toolRegistry[toolName] || null;
}

/**
 * Get all registered tool names.
 */
export function getRegisteredToolNames(): string[] {
  return Object.keys(toolRegistry);
}
