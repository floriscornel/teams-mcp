require('dotenv').config();
const restify = require('restify');
const {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
    MemoryStorage,
    TurnContext
} = require('botbuilder');
const {
    Application,
    ActionPlanner,
    OpenAIModel,
    TurnState
} = require('@microsoft/teams-ai');

const CredentialStore = require('./credential-store');
const MCPSessionManager = require('./mcp-tools');

// =============================================================================
// CONFIGURATION
// =============================================================================

const config = {
    botId: process.env.MicrosoftAppId || '',
    botPassword: process.env.MicrosoftAppPassword || '',
    openAIKey: process.env.OPENAI_API_KEY,
    openAIModel: process.env.OPENAI_MODEL || 'gpt-4-turbo-2024-04-09',
    mcpCommand: process.env.MCP_SERVER_COMMAND,
    mcpArgs: process.env.MCP_SERVER_ARGS ? process.env.MCP_SERVER_ARGS.split(' ') : [],
    port: process.env.PORT || 3978,
    encryptionKey: process.env.ENCRYPTION_KEY
};

// Validate critical configuration
if (!config.openAIKey) {
    throw new Error('OPENAI_API_KEY is required in .env file');
}
if (!config.mcpCommand) {
    throw new Error('MCP_SERVER_COMMAND is required in .env file');
}
if (!config.encryptionKey || config.encryptionKey.length < 32) {
    throw new Error('ENCRYPTION_KEY must be at least 32 characters');
}

console.log('[Wingman] Configuration loaded');
console.log(`[Wingman] Bot ID: ${config.botId || '(not set - local dev mode)'}`);
console.log(`[Wingman] OpenAI Model: ${config.openAIModel}`);
console.log(`[Wingman] MCP Command: ${config.mcpCommand}`);
console.log(`[Wingman] Server Port: ${config.port}`);

// =============================================================================
// INITIALIZE COMPONENTS
// =============================================================================

// Credential store for encrypted credential storage
const credentialStore = new CredentialStore(config.encryptionKey, './credentials.json');

// MCP session manager for OData access
const mcpManager = new MCPSessionManager(config.mcpCommand, config.mcpArgs);

// Create Restify server
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

// =============================================================================
// HEALTH CHECK ENDPOINT
// =============================================================================

server.get('/health', (req, res, next) => {
    res.send(200, {
        status: 'ok',
        service: 'wingman-teams-ai',
        timestamp: new Date().toISOString(),
        sessions: mcpManager.getStats().activeSessions
    });
    next(); // CRITICAL: Must call next() for Restify
});

// =============================================================================
// BOT FRAMEWORK SETUP
// =============================================================================

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
    MicrosoftAppId: config.botId,
    MicrosoftAppPassword: config.botPassword
});

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Error handling
adapter.onTurnError = async (context, error) => {
    console.error(`[Bot Error]`, error);
    await context.sendActivity('Sorry, something went wrong. Please try again.');
};

// =============================================================================
// SYSTEM PROMPT - CRITICAL FOR MULTI-STEP ANALYSIS
// =============================================================================

const SYSTEM_PROMPT = `You are Wingman, an intelligent business intelligence AI assistant with direct access to SAP/OData services.

**Your Core Capabilities:**
You can query OData services to access real business data including Customers, Orders, Products, Employees, and more. You have tools to filter, expand, and analyze this data.

**Available OData Entities:**
- Customers: Customer information (CustomerID, CompanyName, ContactName, Country, City, etc.)
- Orders: Order records (OrderID, CustomerID, EmployeeID, OrderDate, ShipCountry, etc.)
- Order_Details: Line items (OrderID, ProductID, UnitPrice, Quantity, Discount)
- Products: Product catalog (ProductID, ProductName, UnitPrice, UnitsInStock, CategoryID, etc.)
- Employees: Employee records (EmployeeID, FirstName, LastName, Title, etc.)
- Categories: Product categories (CategoryID, CategoryName, Description)
- Suppliers: Supplier information (SupplierID, CompanyName, Country, etc.)

**CRITICAL - Multi-Step Analysis:**
You can and MUST perform complex multi-step analysis by making MULTIPLE tool calls in sequence. Don't limit yourself - make AS MANY tool calls as needed to complete the analysis thoroughly.

**Example: ABC Customer Segmentation**
When asked to perform ABC customer segmentation, follow these steps:

Step 1: Fetch orders with details
   - Call filter_Orders with: {"$expand": "Order_Details,Customer", "$top": 1000}
   - This gets orders with line items and customer info

Step 2: Fetch additional data if needed
   - If you need more complete data, call filter_Order_Details separately
   - Use $top to control result size, make multiple calls if needed

Step 3: Calculate revenue per customer
   - Process the data to calculate: Revenue = UnitPrice × Quantity × (1 - Discount)
   - Group by customer and sum total revenue
   - YOU must do this calculation - don't ask the user

Step 4: Classify customers into ABC segments
   - Sort customers by revenue (descending)
   - Calculate cumulative revenue percentage
   - Classify:
     * A customers: Top 20% that generate ~80% of revenue
     * B customers: Next 30% that generate ~15% of revenue
     * C customers: Bottom 50% that generate ~5% of revenue

Step 5: Present results
   - Show customers in each segment
   - Include revenue totals and percentages
   - Format as a clear table

**Other Analysis Patterns:**

**Revenue Analysis:**
1. Fetch Order_Details with $expand for Orders and Products
2. Calculate: Revenue = UnitPrice × Quantity × (1 - Discount)
3. Group by time period, product, or customer as needed
4. Present trends and insights

**Customer Analysis:**
1. Fetch Customers with $expand for Orders
2. Calculate metrics (order count, average order value, etc.)
3. Segment by country, city, or other dimensions
4. Identify patterns and top customers

**Product Performance:**
1. Fetch Products with $expand for Order_Details
2. Calculate sales volume and revenue
3. Identify best/worst performers
4. Analyze by category or supplier

**OData Query Syntax Examples:**

Basic filter:
{"$filter": "Country eq 'Germany'"}

Expand related entities:
{"$expand": "Order_Details,Customer"}

Combine multiple operations:
{
  "$filter": "OrderDate ge 1997-01-01",
  "$expand": "Order_Details,Customer",
  "$top": 100,
  "$orderby": "OrderDate desc"
}

Select specific fields:
{"$select": "CustomerID,CompanyName,Country", "$top": 20}

**IMPORTANT RULES:**

1. **Make Multiple Tool Calls:** Don't limit yourself to one or two calls. Complex analysis requires 3-7+ calls. Make them!

2. **Fetch Complete Data:** If you get partial data, make MORE calls to get what you need. Use $top and $skip for pagination if needed.

3. **Do The Math:** YOU must perform all calculations. Don't ask the user to calculate revenue, percentages, or classifications. Process the data yourself.

4. **Show Your Work:** Explain what you're doing: "I fetched X records...", "Calculating revenue for Y customers...", "Classifying into ABC segments..."

5. **Format Results Clearly:** Use tables, bullet points, and clear headers. Include totals and percentages.

6. **Handle Errors Gracefully:** If a query fails, try a simpler query or explain what went wrong.

7. **Be Thorough:** For complex questions, break them down into steps and execute each step systematically.

**Response Format:**
- Use markdown for formatting
- Create tables for structured data
- Use bullet points for lists
- Include summary statistics
- Explain your analysis process

Remember: You are powerful! Make multiple tool calls, process data thoroughly, and provide complete business intelligence insights.`;

// =============================================================================
// TEAMS AI APPLICATION
// =============================================================================

// Storage for conversation state
const storage = new MemoryStorage();

// Create AI model
const model = new OpenAIModel({
    apiKey: config.openAIKey,
    defaultModel: config.openAIModel,
    useSystemMessages: true,
    logRequests: true
});

// Create application with AI capabilities
const app = new Application({
    storage,
    ai: {
        planner: new ActionPlanner({
            model,
            prompts: {
                default: {
                    text: SYSTEM_PROMPT,
                    config: {
                        schema: 1,
                        type: 'completion',
                        completion: {
                            model: config.openAIModel,
                            temperature: 0.3, // Low temperature for consistent analysis
                            max_tokens: 4000,
                            top_p: 1.0,
                            presence_penalty: 0,
                            frequency_penalty: 0
                        }
                    }
                }
            },
            defaultPrompt: 'default'
        }),
        // CRITICAL: Set high max iterations for multi-step analysis
        max_iterations: 10,
        max_time: 300000 // 5 minutes
    }
});

// =============================================================================
// ADAPTIVE CARDS
// =============================================================================

/**
 * Create the initial welcome card with setup button
 */
function createWelcomeCard() {
    return {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
            {
                type: 'TextBlock',
                text: '👋 Welcome to Wingman',
                size: 'Large',
                weight: 'Bolder'
            },
            {
                type: 'TextBlock',
                text: 'Your AI-powered business intelligence assistant',
                wrap: true,
                spacing: 'Small'
            },
            {
                type: 'TextBlock',
                text: 'I can help you analyze OData/SAP data using natural language. Ask me questions like:',
                wrap: true,
                spacing: 'Medium'
            },
            {
                type: 'TextBlock',
                text: '• "Show customers from Germany"\n• "Do ABC customer segmentation"\n• "What are our top products by revenue?"\n• "Analyze orders from the last quarter"',
                wrap: true,
                spacing: 'Small'
            },
            {
                type: 'TextBlock',
                text: '**First, set up your OData credentials:**',
                wrap: true,
                weight: 'Bolder',
                spacing: 'Medium'
            }
        ],
        actions: [
            {
                type: 'Action.Submit',
                title: '🔧 Setup Demo',
                data: { action: 'setup' }
            }
        ]
    };
}

/**
 * Create the setup card for credential input
 */
function createSetupCard() {
    return {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
            {
                type: 'TextBlock',
                text: '🔧 OData Service Setup',
                size: 'Large',
                weight: 'Bolder'
            },
            {
                type: 'TextBlock',
                text: 'Enter your OData service credentials. These will be encrypted and stored securely.',
                wrap: true,
                spacing: 'Small'
            },
            {
                type: 'TextBlock',
                text: '**For Demo (Northwind OData):** Leave blank or use any values',
                wrap: true,
                spacing: 'Medium',
                color: 'Accent'
            },
            {
                type: 'Input.Text',
                id: 'username',
                label: 'Username',
                placeholder: 'Enter username',
                value: 'demo'
            },
            {
                type: 'Input.Text',
                id: 'password',
                label: 'Password',
                placeholder: 'Enter password',
                style: 'password',
                value: 'demo'
            }
        ],
        actions: [
            {
                type: 'Action.Submit',
                title: '💾 Save Credentials',
                data: { action: 'save_credentials' }
            }
        ]
    };
}

/**
 * Create success confirmation card
 */
function createSuccessCard() {
    return {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
            {
                type: 'TextBlock',
                text: '✅ Setup Complete!',
                size: 'Large',
                weight: 'Bolder',
                color: 'Good'
            },
            {
                type: 'TextBlock',
                text: 'Your credentials have been encrypted and saved. You can now ask me business intelligence questions!',
                wrap: true,
                spacing: 'Small'
            },
            {
                type: 'TextBlock',
                text: '**Try asking:**',
                wrap: true,
                weight: 'Bolder',
                spacing: 'Medium'
            },
            {
                type: 'TextBlock',
                text: '• "List all customers from Germany"\n• "Perform ABC customer segmentation"\n• "Show top 10 products by sales"\n• "Analyze revenue by country"',
                wrap: true,
                spacing: 'Small'
            }
        ]
    };
}

// =============================================================================
// MESSAGE HANDLERS
// =============================================================================

/**
 * Single unified message handler for both text messages and card submissions
 */
app.activity('message', async (context, state) => {
    try {
        const userId = context.activity.from.id;
        const userName = context.activity.from.name || 'User';

        console.log(`[Message] From: ${userName} (${userId})`);

        // Handle card submissions
        if (context.activity.value?.action) {
            const action = context.activity.value.action;

            if (action === 'setup') {
                // Show setup card
                await context.sendActivity({
                    attachments: [{
                        contentType: 'application/vnd.microsoft.card.adaptive',
                        content: createSetupCard()
                    }]
                });
                return;
            }

            if (action === 'save_credentials') {
                // Save credentials
                const username = context.activity.value.username;
                const password = context.activity.value.password;

                if (!username || !password) {
                    await context.sendActivity('❌ Please provide both username and password.');
                    return;
                }

                // Store encrypted credentials
                credentialStore.storeCredentials(userId, username, password);

                // Send success card
                await context.sendActivity({
                    attachments: [{
                        contentType: 'application/vnd.microsoft.card.adaptive',
                        content: createSuccessCard()
                    }]
                });
                return;
            }
        }

        // Handle text messages
        if (context.activity.text) {
            const message = context.activity.text.trim();

            // Check for setup commands
            if (message.toLowerCase().includes('setup') || message.toLowerCase().includes('configure')) {
                await context.sendActivity({
                    attachments: [{
                        contentType: 'application/vnd.microsoft.card.adaptive',
                        content: createSetupCard()
                    }]
                });
                return;
            }

            // Check for welcome/help commands
            if (message.toLowerCase() === 'help' || message.toLowerCase() === 'start') {
                await context.sendActivity({
                    attachments: [{
                        contentType: 'application/vnd.microsoft.card.adaptive',
                        content: createWelcomeCard()
                    }]
                });
                return;
            }

            // Check if user has credentials
            if (!credentialStore.hasCredentials(userId)) {
                await context.sendActivity('👋 Welcome! Please set up your OData credentials first.');
                await context.sendActivity({
                    attachments: [{
                        contentType: 'application/vnd.microsoft.card.adaptive',
                        content: createWelcomeCard()
                    }]
                });
                return;
            }

            // Get credentials and create MCP session
            const credentials = credentialStore.getCredentials(userId);
            const session = await mcpManager.getSession(userId, credentials);

            // Register MCP tools dynamically with the AI
            const tools = mcpManager.getToolsForOpenAI(userId);

            if (tools.length === 0) {
                await context.sendActivity('⚠️ No OData tools available. Please check your MCP server configuration.');
                return;
            }

            console.log(`[Message] Available tools for ${userName}: ${tools.length}`);

            // Send typing indicator
            await context.sendActivity({ type: 'typing' });

            // Process the message with AI
            // The Teams AI library will handle the tool calling loop
            try {
                // This is handled automatically by the Teams AI library
                // The app will process the message through the AI planner
                // which will make tool calls as needed
                return true; // Continue to AI processing

            } catch (error) {
                console.error('[Message] AI processing error:', error);
                await context.sendActivity('❌ Error processing your request. Please try again.');
                return false;
            }
        }

    } catch (error) {
        console.error('[Message Handler Error]:', error);
        await context.sendActivity('Sorry, I encountered an error. Please try again.');
    }
});

/**
 * Register MCP tools as AI actions
 * This allows the AI to call OData queries
 */
app.ai.action('filter_Orders', async (context, state, parameters) => {
    const userId = context.activity.from.id;
    console.log('[AI Action] filter_Orders called with:', parameters);

    try {
        const result = await mcpManager.executeTool(userId, 'filter_Orders', parameters);
        return JSON.stringify(result);
    } catch (error) {
        console.error('[AI Action] Error:', error);
        return JSON.stringify({ error: error.message });
    }
});

app.ai.action('filter_Customers', async (context, state, parameters) => {
    const userId = context.activity.from.id;
    console.log('[AI Action] filter_Customers called with:', parameters);

    try {
        const result = await mcpManager.executeTool(userId, 'filter_Customers', parameters);
        return JSON.stringify(result);
    } catch (error) {
        console.error('[AI Action] Error:', error);
        return JSON.stringify({ error: error.message });
    }
});

app.ai.action('filter_Order_Details', async (context, state, parameters) => {
    const userId = context.activity.from.id;
    console.log('[AI Action] filter_Order_Details called with:', parameters);

    try {
        const result = await mcpManager.executeTool(userId, 'filter_Order_Details', parameters);
        return JSON.stringify(result);
    } catch (error) {
        console.error('[AI Action] Error:', error);
        return JSON.stringify({ error: error.message });
    }
});

app.ai.action('filter_Products', async (context, state, parameters) => {
    const userId = context.activity.from.id;
    console.log('[AI Action] filter_Products called with:', parameters);

    try {
        const result = await mcpManager.executeTool(userId, 'filter_Products', parameters);
        return JSON.stringify(result);
    } catch (error) {
        console.error('[AI Action] Error:', error);
        return JSON.stringify({ error: error.message });
    }
});

// Add more tool actions as discovered from MCP server
// These are common Northwind OData entities

app.ai.action('filter_Employees', async (context, state, parameters) => {
    const userId = context.activity.from.id;
    try {
        const result = await mcpManager.executeTool(userId, 'filter_Employees', parameters);
        return JSON.stringify(result);
    } catch (error) {
        return JSON.stringify({ error: error.message });
    }
});

app.ai.action('filter_Categories', async (context, state, parameters) => {
    const userId = context.activity.from.id;
    try {
        const result = await mcpManager.executeTool(userId, 'filter_Categories', parameters);
        return JSON.stringify(result);
    } catch (error) {
        return JSON.stringify({ error: error.message });
    }
});

app.ai.action('filter_Suppliers', async (context, state, parameters) => {
    const userId = context.activity.from.id;
    try {
        const result = await mcpManager.executeTool(userId, 'filter_Suppliers', parameters);
        return JSON.stringify(result);
    } catch (error) {
        return JSON.stringify({ error: error.message });
    }
});

// =============================================================================
// CONVERSATION UPDATE HANDLER (Welcome Message)
// =============================================================================

app.activity('conversationUpdate', async (context, state) => {
    if (context.activity.membersAdded) {
        for (const member of context.activity.membersAdded) {
            if (member.id !== context.activity.recipient.id) {
                // Send welcome card to new user
                await context.sendActivity({
                    attachments: [{
                        contentType: 'application/vnd.microsoft.card.adaptive',
                        content: createWelcomeCard()
                    }]
                });
            }
        }
    }
});

// =============================================================================
// START SERVER
// =============================================================================

// Handle bot messages
server.post('/api/messages', async (req, res) => {
    await adapter.process(req, res, async (context) => {
        await app.run(context);
    });
});

// Start listening
server.listen(config.port, () => {
    console.log(`\n\n✅ Wingman Teams AI Bot is running!`);
    console.log(`🌐 Server: http://localhost:${config.port}`);
    console.log(`💚 Health: http://localhost:${config.port}/health`);
    console.log(`🤖 Messages: http://localhost:${config.port}/api/messages`);
    console.log(`\n📝 Next steps:`);
    console.log(`   1. Test health endpoint: curl http://localhost:${config.port}/health`);
    console.log(`   2. Connect Bot Framework Emulator to: http://localhost:${config.port}/api/messages`);
    console.log(`   3. Start chatting with Wingman!\n`);
});

// Graceful shutdown
process.on('SIGINT', async () => {
    console.log('\n[Wingman] Shutting down gracefully...');
    await mcpManager.closeAll();
    process.exit(0);
});

process.on('SIGTERM', async () => {
    console.log('\n[Wingman] Shutting down gracefully...');
    await mcpManager.closeAll();
    process.exit(0);
});
