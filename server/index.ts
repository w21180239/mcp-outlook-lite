#!/usr/bin/env node

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
  console.error('Stack trace:', (reason as any).stack || 'No stack trace available');
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  console.error('Stack trace:', error.stack);
  if ((error as any).code === 'MODULE_NOT_FOUND' || error.name === 'SyntaxError') {
    process.exit(1);
  }
});

(async function initializeServer() {
  try {
    await import('dotenv/config');

    const { Server } = await import('@modelcontextprotocol/sdk/server/index.js');
    const { StdioServerTransport } = await import('@modelcontextprotocol/sdk/server/stdio.js');
    const {
      ListToolsRequestSchema,
      CallToolRequestSchema,
      InitializeRequestSchema,
      InitializedNotificationSchema,
      ListPromptsRequestSchema,
      GetPromptRequestSchema
    } = await import('@modelcontextprotocol/sdk/types.js');

    const { OutlookAuthManager } = await import('./auth/auth.js');
    const { createProtocolError, ErrorCodes, convertErrorToToolError } = await import('./utils/mcpErrorResponse.js');
    const { allToolSchemas } = await import('./schemas/toolSchemas.js');
    const { getToolHandler, getRegisteredToolNames } = await import('./tools/dispatcher.js');
    const { promptList, getPrompt } = await import('./prompts/index.js');
    const { debug } = await import('./utils/logger.js');

    if (!process.env.AZURE_CLIENT_ID || !process.env.AZURE_TENANT_ID) {
      console.error('Error: AZURE_CLIENT_ID and AZURE_TENANT_ID environment variables are required.');
      process.exit(1);
    }

    const server = new Server(
      { name: 'outlook-mcp', version: '2.0.0' },
      { capabilities: { tools: {}, prompts: {} } }
    );

    const authManager = new OutlookAuthManager(
      process.env.AZURE_CLIENT_ID,
      process.env.AZURE_TENANT_ID
    );

    server.setRequestHandler(InitializeRequestSchema, async () => ({
      protocolVersion: '2025-06-18',
      capabilities: { tools: {}, prompts: {} },
      serverInfo: { name: 'outlook-mcp', version: '2.0.0' },
    }));

    server.setNotificationHandler(InitializedNotificationSchema, async () => {
      debug('Debug: Client initialized');
    });

    server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: allToolSchemas,
    }));

    server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      debug(`DEBUG Tool Dispatch: Called tool '${name}' with args:`, JSON.stringify(args, null, 2));

      const handler = getToolHandler(name);
      if (!handler) {
        return createProtocolError(
          ErrorCodes.METHOD_NOT_FOUND,
          `Unknown tool: ${name}`,
          { availableTools: getRegisteredToolNames() }
        );
      }

      try {
        return await handler(authManager, args);
      } catch (error: unknown) {
        console.error('Unexpected error in tool handler:', error);
        const err = error as Record<string, unknown>;
        if (err.content && err.isError !== undefined) {
          return error;
        }
        return convertErrorToToolError(error, 'Tool execution failed');
      }
    });

    server.setRequestHandler(ListPromptsRequestSchema, async () => ({
      prompts: promptList,
    }));

    server.setRequestHandler(GetPromptRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      try {
        return await getPrompt(name, args);
      } catch (error) {
        console.error(`Error getting prompt ${name}:`, error);
        throw convertErrorToToolError(error, `Failed to get prompt ${name}`);
      }
    });

    console.error('Starting Outlook MCP Server...');
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error('Outlook MCP server is ready and connected');

  } catch (error) {
    console.error('Server error:', error);
    process.exit(1);
  }
})();
