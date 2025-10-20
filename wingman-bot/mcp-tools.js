const { Client } = require('@modelcontextprotocol/sdk/client/index.js');
const { StdioClientTransport } = require('@modelcontextprotocol/sdk/client/stdio.js');

/**
 * MCPSessionManager - Manages per-user MCP sessions with automatic cleanup
 *
 * CRITICAL: Uses StdioClientTransport to let the MCP SDK spawn the process
 * Do NOT manually spawn the process - let the transport handle it
 *
 * Features:
 * - Per-user session isolation
 * - Automatic session cleanup after 15 minutes of inactivity
 * - Dynamic tool discovery from MCP server
 * - Credential injection for authenticated OData access
 */
class MCPSessionManager {
    constructor(mcpCommand, mcpArgs) {
        this.mcpCommand = mcpCommand;
        this.mcpArgs = mcpArgs || [];
        this.sessions = new Map(); // userId -> session object
        this.sessionTimeouts = new Map(); // userId -> timeout handle
        this.SESSION_TIMEOUT_MS = 15 * 60 * 1000; // 15 minutes

        console.log(`[MCPSessionManager] Initialized with command: ${mcpCommand}`);
        console.log(`[MCPSessionManager] Args: ${JSON.stringify(mcpArgs)}`);
    }

    /**
     * Get or create an MCP session for a user
     * @param {string} userId - User identifier
     * @param {Object} credentials - {username, password} for OData service
     * @returns {Promise<Object>} Session object with client and tools
     */
    async getSession(userId, credentials = null) {
        try {
            // Check if session already exists
            if (this.sessions.has(userId)) {
                console.log(`[MCPSessionManager] Reusing existing session for user: ${userId}`);
                this._refreshSessionTimeout(userId);
                return this.sessions.get(userId);
            }

            console.log(`[MCPSessionManager] Creating new session for user: ${userId}`);

            // Create new MCP client
            const client = new Client(
                {
                    name: 'wingman-teams-bot',
                    version: '1.0.0'
                },
                {
                    capabilities: {
                        tools: {}
                    }
                }
            );

            // Prepare environment variables for credentials if provided
            const env = { ...process.env };
            if (credentials) {
                env.ODATA_USERNAME = credentials.username;
                env.ODATA_PASSWORD = credentials.password;
                console.log(`[MCPSessionManager] Injecting credentials for user: ${userId}`);
            }

            // Create transport - let the SDK spawn the process
            const transport = new StdioClientTransport({
                command: this.mcpCommand,
                args: this.mcpArgs,
                env: env
            });

            // Connect to MCP server
            console.log(`[MCPSessionManager] Connecting to MCP server...`);
            await client.connect(transport);
            console.log(`[MCPSessionManager] Connected successfully`);

            // Discover available tools
            const toolsResponse = await client.listTools();
            const tools = toolsResponse.tools || [];

            console.log(`[MCPSessionManager] Discovered ${tools.length} tools:`, tools.map(t => t.name));

            // Create session object
            const session = {
                client,
                transport,
                tools,
                createdAt: new Date(),
                lastUsedAt: new Date(),
                userId
            };

            // Store session
            this.sessions.set(userId, session);

            // Set up auto-cleanup timeout
            this._refreshSessionTimeout(userId);

            return session;

        } catch (error) {
            console.error(`[MCPSessionManager] Error creating session for user ${userId}:`, error);
            throw new Error(`Failed to create MCP session: ${error.message}`);
        }
    }

    /**
     * Execute a tool call via MCP
     * @param {string} userId - User identifier
     * @param {string} toolName - Name of the tool to call
     * @param {Object} args - Tool arguments
     * @returns {Promise<Object>} Tool execution result
     */
    async executeTool(userId, toolName, args) {
        try {
            console.log(`[MCPSessionManager] Executing tool: ${toolName} for user: ${userId}`);
            console.log(`[MCPSessionManager] Tool args:`, JSON.stringify(args, null, 2));

            const session = this.sessions.get(userId);
            if (!session) {
                throw new Error(`No active session for user: ${userId}`);
            }

            // Update last used timestamp
            session.lastUsedAt = new Date();
            this._refreshSessionTimeout(userId);

            // Call the tool via MCP
            const result = await session.client.callTool({
                name: toolName,
                arguments: args
            });

            console.log(`[MCPSessionManager] Tool execution successful:`, toolName);

            return result;

        } catch (error) {
            console.error(`[MCPSessionManager] Tool execution error:`, error);
            throw new Error(`Tool execution failed: ${error.message}`);
        }
    }

    /**
     * Get available tools for a user's session
     * @param {string} userId - User identifier
     * @returns {Array} Array of tool definitions
     */
    getAvailableTools(userId) {
        const session = this.sessions.get(userId);
        if (!session) {
            return [];
        }
        return session.tools;
    }

    /**
     * Close and cleanup a user's session
     * @param {string} userId - User identifier
     * @returns {Promise<void>}
     */
    async closeSession(userId) {
        try {
            console.log(`[MCPSessionManager] Closing session for user: ${userId}`);

            const session = this.sessions.get(userId);
            if (!session) {
                console.log(`[MCPSessionManager] No session to close for user: ${userId}`);
                return;
            }

            // Clear timeout
            const timeout = this.sessionTimeouts.get(userId);
            if (timeout) {
                clearTimeout(timeout);
                this.sessionTimeouts.delete(userId);
            }

            // Close MCP client connection
            try {
                await session.client.close();
            } catch (error) {
                console.error(`[MCPSessionManager] Error closing client:`, error);
            }

            // Remove session
            this.sessions.delete(userId);

            console.log(`[MCPSessionManager] Session closed for user: ${userId}`);

        } catch (error) {
            console.error(`[MCPSessionManager] Error closing session:`, error);
        }
    }

    /**
     * Refresh the session timeout for a user
     * Auto-closes the session after SESSION_TIMEOUT_MS of inactivity
     * @param {string} userId - User identifier
     * @private
     */
    _refreshSessionTimeout(userId) {
        // Clear existing timeout
        const existingTimeout = this.sessionTimeouts.get(userId);
        if (existingTimeout) {
            clearTimeout(existingTimeout);
        }

        // Set new timeout
        const timeout = setTimeout(async () => {
            console.log(`[MCPSessionManager] Session timeout reached for user: ${userId}`);
            await this.closeSession(userId);
        }, this.SESSION_TIMEOUT_MS);

        this.sessionTimeouts.set(userId, timeout);
    }

    /**
     * Check if a user has an active session
     * @param {string} userId - User identifier
     * @returns {boolean}
     */
    hasSession(userId) {
        return this.sessions.has(userId);
    }

    /**
     * Get session statistics
     * @returns {Object} Statistics about active sessions
     */
    getStats() {
        const stats = {
            activeSessions: this.sessions.size,
            sessions: []
        };

        for (const [userId, session] of this.sessions.entries()) {
            stats.sessions.push({
                userId,
                createdAt: session.createdAt,
                lastUsedAt: session.lastUsedAt,
                toolCount: session.tools.length,
                tools: session.tools.map(t => t.name)
            });
        }

        return stats;
    }

    /**
     * Close all sessions and cleanup
     * @returns {Promise<void>}
     */
    async closeAll() {
        console.log(`[MCPSessionManager] Closing all sessions...`);

        const userIds = Array.from(this.sessions.keys());
        for (const userId of userIds) {
            await this.closeSession(userId);
        }

        console.log(`[MCPSessionManager] All sessions closed`);
    }

    /**
     * Convert MCP tools to OpenAI function calling format
     * @param {string} userId - User identifier
     * @returns {Array} Array of OpenAI function definitions
     */
    getToolsForOpenAI(userId) {
        const session = this.sessions.get(userId);
        if (!session || !session.tools) {
            return [];
        }

        return session.tools.map(tool => ({
            type: 'function',
            function: {
                name: tool.name,
                description: tool.description || `Execute ${tool.name}`,
                parameters: tool.inputSchema || {
                    type: 'object',
                    properties: {},
                    required: []
                }
            }
        }));
    }
}

module.exports = MCPSessionManager;
