# Wingman Teams AI Bot

**AI-powered business intelligence assistant for Microsoft Teams with SAP/OData integration via MCP (Model Context Protocol)**

Wingman allows users to query OData services using natural language. Ask questions like "Show customers from Germany" or "Do ABC customer segmentation" and get AI-powered insights from real business data.

## 🏗️ Architecture

- **Bot Framework**: @microsoft/teams-ai (Microsoft Teams AI Library)
- **AI Engine**: OpenAI GPT-4 with function calling
- **Data Bridge**: MCP protocol (@modelcontextprotocol/sdk) to odata-mcp.exe (Go binary)
- **Security**: AES-256-CBC encrypted credential storage
- **Server**: Restify on port 3978
- **Node.js**: 18+ (tested with Node.js 22)

## ✨ Features

- 🤖 **Natural Language Queries**: Ask business questions in plain English
- 🔄 **Multi-Step Analysis**: AI automatically makes multiple tool calls for complex analysis
- 📊 **ABC Segmentation**: Built-in customer segmentation capabilities
- 🔐 **Secure Credentials**: AES-256 encrypted storage with IV
- 👥 **Multi-User Support**: Isolated MCP sessions per user
- ⚡ **Session Management**: Auto-cleanup after 15 minutes of inactivity
- 💬 **Adaptive Cards**: Interactive UI for setup and configuration

## 📋 Prerequisites

1. **Node.js 18+** (tested with Node.js 22)
2. **OpenAI API Key** (GPT-4 access)
3. **OData MCP Server** (e.g., odata-mcp.exe)
4. **Bot Framework Emulator** (for local testing) or Teams app registration

## 🚀 Quick Start

### 1. Install Dependencies

```bash
cd wingman-bot
npm install
```

### 2. Configure Environment

Copy `.env.example` to `.env` and fill in your values:

```bash
cp .env.example .env
```

Edit `.env`:

```bash
# Microsoft Bot Framework (leave empty for local dev)
MicrosoftAppId=
MicrosoftAppPassword=

# OpenAI Configuration (REQUIRED)
OPENAI_API_KEY=sk-your-actual-key-here
OPENAI_MODEL=gpt-4-turbo-2024-04-09

# MCP Server Configuration (REQUIRED)
# Windows example:
MCP_SERVER_COMMAND=C:\\path\\to\\odata-mcp.exe
# Linux/Mac example:
# MCP_SERVER_COMMAND=/path/to/odata-mcp

# MCP Arguments - Demo Northwind OData service
MCP_SERVER_ARGS=--transport stdio --service https://services.odata.org/V2/Northwind/Northwind.svc/

# Server Configuration
PORT=3978

# Encryption Key (REQUIRED - Generate a secure random string)
ENCRYPTION_KEY=your-secure-32-character-minimum-encryption-key-here
```

### 3. Start the Bot

```bash
npm start
```

You should see:

```
✅ Wingman Teams AI Bot is running!
🌐 Server: http://localhost:3978
💚 Health: http://localhost:3978/health
🤖 Messages: http://localhost:3978/api/messages
```

### 4. Test Health Endpoint

```bash
curl http://localhost:3978/health
```

Expected response:
```json
{
  "status": "ok",
  "service": "wingman-teams-ai",
  "timestamp": "2024-01-15T10:30:00.000Z",
  "sessions": 0
}
```

### 5. Connect Bot Framework Emulator

1. Download [Bot Framework Emulator](https://github.com/Microsoft/BotFramework-Emulator/releases)
2. Open the emulator
3. Click "Open Bot"
4. Enter Bot URL: `http://localhost:3978/api/messages`
5. Leave App ID and Password empty (for local dev)
6. Click "Connect"

### 6. Start Chatting!

In the emulator, you'll see a welcome card. Click "Setup Demo" and enter credentials (for Northwind demo, use any values like "demo"/"demo").

Then try these queries:

- "List all customers from Germany"
- "Do ABC customer segmentation"
- "Show top 10 products by revenue"
- "Analyze orders from France"
- "What are the best selling categories?"

## 📁 Project Structure

```
wingman-bot/
├── index.js              # Main bot application
├── mcp-tools.js          # MCP session manager
├── credential-store.js   # Encrypted credential storage
├── package.json          # Dependencies
├── .env.example          # Environment template
├── .env                  # Your configuration (not in git)
├── credentials.json      # Encrypted credentials (auto-created)
└── README.md            # This file
```

## 🔧 Configuration Details

### Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `MicrosoftAppId` | No (local) / Yes (prod) | Bot Framework App ID |
| `MicrosoftAppPassword` | No (local) / Yes (prod) | Bot Framework Password |
| `OPENAI_API_KEY` | **Yes** | Your OpenAI API key |
| `OPENAI_MODEL` | No | OpenAI model (default: gpt-4-turbo-2024-04-09) |
| `MCP_SERVER_COMMAND` | **Yes** | Path to odata-mcp executable |
| `MCP_SERVER_ARGS` | **Yes** | Arguments for MCP server |
| `PORT` | No | Server port (default: 3978) |
| `ENCRYPTION_KEY` | **Yes** | Min 32 chars for AES-256 encryption |

### MCP Server Setup

The bot requires an OData MCP server. Example using the public Northwind demo:

```bash
# MCP_SERVER_ARGS format:
--transport stdio --service https://services.odata.org/V2/Northwind/Northwind.svc/
```

For production, replace with your SAP/OData endpoint:

```bash
MCP_SERVER_ARGS=--transport stdio --service https://your-sap-server.com/odata/v4/service/ --username ${ODATA_USERNAME} --password ${ODATA_PASSWORD}
```

## 🎯 Usage Examples

### Simple Query
```
User: "Show me customers from Germany"

Wingman: [Fetches and displays German customers in a table]
```

### Complex Analysis
```
User: "Do ABC customer segmentation"

Wingman:
Step 1: Fetching orders with customer details...
Step 2: Calculating revenue per customer...
Step 3: Classifying into ABC segments...

**ABC Customer Segmentation Results**

**A Customers (Top 20% - 80% of revenue):**
| Customer | Revenue | % of Total |
|----------|---------|------------|
| QUICK    | $110,277 | 15.2% |
| SAVEA    | $115,673 | 16.0% |
...

**B Customers (Next 30% - 15% of revenue):**
...

**C Customers (Bottom 50% - 5% of revenue):**
...
```

### Multi-Entity Query
```
User: "Show orders from France with product details"

Wingman: [Fetches Orders with $expand=Order_Details,Products and filters by France]
```

## 🛡️ Security Features

### Encrypted Credential Storage

Credentials are encrypted using **AES-256-CBC** with:
- Random 16-byte IV per encryption
- SHA-256 derived encryption key
- Format: `iv:encrypted_data` (hex encoded)

Example:
```javascript
const credentialStore = new CredentialStore(encryptionKey);
credentialStore.storeCredentials(userId, 'username', 'password');
const creds = credentialStore.getCredentials(userId);
```

### Per-User Session Isolation

Each user gets their own MCP session with:
- Isolated credential injection
- Automatic cleanup after 15 minutes
- No cross-user data leakage

## 🔍 Advanced Features

### Multi-Step Analysis

The bot is configured with `maxIterations: 10` to allow complex multi-step analysis:

```javascript
ai: {
    max_iterations: 10,  // Allow up to 10 tool calls
    max_time: 300000     // 5 minute timeout
}
```

The AI automatically:
1. Breaks down complex queries
2. Makes multiple tool calls
3. Processes and aggregates data
4. Performs calculations
5. Presents formatted results

### Available OData Entities

When connected to Northwind demo:
- Customers
- Orders
- Order_Details
- Products
- Employees
- Categories
- Suppliers
- Shippers
- Regions
- Territories

### Supported OData Queries

The MCP server provides tools for each entity with support for:
- `$filter`: Filter results (e.g., `Country eq 'Germany'`)
- `$expand`: Include related entities (e.g., `Order_Details,Customer`)
- `$select`: Choose specific fields
- `$top`: Limit results
- `$skip`: Pagination
- `$orderby`: Sort results

## 🧪 Testing Scenarios

### 1. Setup Flow
1. Start conversation → Welcome card appears
2. Click "Setup Demo" → Setup card with username/password
3. Submit credentials → Success message
4. Ask query → AI processes with OData access

### 2. Simple Queries
- "List customers from Germany"
- "Show all products"
- "Get orders from 1997"

### 3. Complex Analysis
- "Do ABC customer segmentation"
- "Analyze revenue by country"
- "Top 10 products by sales volume"
- "Customer order frequency analysis"

### 4. Multi-Entity Queries
- "Show orders with product details"
- "List customers and their order count"
- "Products by category with sales data"

## 🐛 Troubleshooting

### Bot won't start

**Check configuration:**
```bash
# Ensure all required env vars are set
cat .env | grep OPENAI_API_KEY
cat .env | grep MCP_SERVER_COMMAND
cat .env | grep ENCRYPTION_KEY
```

**Test health endpoint:**
```bash
curl http://localhost:3978/health
```

### MCP Connection Failed

**Check MCP server path:**
```bash
# Windows
dir C:\path\to\odata-mcp.exe

# Linux/Mac
ls -la /path/to/odata-mcp
```

**Test MCP server manually:**
```bash
./odata-mcp.exe --transport stdio --service https://services.odata.org/V2/Northwind/Northwind.svc/
```

### No Tools Available

**Check logs:**
```bash
# Look for:
[MCPSessionManager] Discovered X tools: [tool names]
```

**Verify MCP server supports tool discovery:**
- MCP server must implement `list_tools` protocol
- Server must be accessible and responding

### Credentials Not Saving

**Check encryption key length:**
```bash
# Must be at least 32 characters
echo -n $ENCRYPTION_KEY | wc -c
```

**Check file permissions:**
```bash
# Ensure credentials.json is writable
ls -la credentials.json
```

### AI Not Making Multiple Tool Calls

**Check system prompt:**
- Ensure SYSTEM_PROMPT includes multi-step instructions
- Verify `maxIterations: 10` is set
- Check OpenAI model supports function calling

## 📊 Monitoring

### Session Statistics

Access session stats via health endpoint:
```bash
curl http://localhost:3978/health | jq
```

### Logs

The bot provides detailed logging:
```
[Wingman] Configuration loaded
[CredentialStore] Initialized with store at: ./credentials.json
[MCPSessionManager] Initialized with command: ./odata-mcp.exe
[MCPSessionManager] Creating new session for user: user123
[MCPSessionManager] Discovered 7 tools: [filter_Customers, filter_Orders, ...]
[AI Action] filter_Orders called with: {"$expand":"Order_Details,Customer"}
```

## 🚀 Production Deployment

### Azure Bot Service

1. Register bot at https://dev.botframework.com/
2. Get App ID and Password
3. Update .env with credentials
4. Deploy to Azure App Service or Container Instance
5. Configure messaging endpoint in Bot Framework

### Environment Setup

```bash
# Production .env
MicrosoftAppId=your-app-id
MicrosoftAppPassword=your-app-password
OPENAI_API_KEY=sk-production-key
MCP_SERVER_COMMAND=/usr/local/bin/odata-mcp
MCP_SERVER_ARGS=--transport stdio --service https://production-odata-service.com
ENCRYPTION_KEY=production-secure-random-key-at-least-32-chars
PORT=3978
```

### Security Checklist

- ✅ Use strong ENCRYPTION_KEY (generate with crypto.randomBytes)
- ✅ Store credentials.json securely (consider Azure Key Vault)
- ✅ Enable HTTPS for production endpoints
- ✅ Implement rate limiting
- ✅ Monitor and log access
- ✅ Rotate encryption keys periodically
- ✅ Use Azure AD authentication for OData services

## 📚 Additional Resources

- [Microsoft Teams AI Library](https://github.com/microsoft/teams-ai)
- [Bot Framework Documentation](https://docs.microsoft.com/en-us/azure/bot-service/)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [OpenAI Function Calling](https://platform.openai.com/docs/guides/function-calling)
- [OData Protocol](https://www.odata.org/)

## 🤝 Contributing

This is an MVP implementation. Areas for enhancement:

- [ ] Support for more OData operations (POST, PATCH, DELETE)
- [ ] Advanced query builder UI
- [ ] Result caching and pagination
- [ ] Export results to Excel/CSV
- [ ] Scheduled reports and alerts
- [ ] Multi-language support
- [ ] Voice interaction support

## 📝 License

MIT License - See LICENSE file for details

## 🆘 Support

For issues or questions:

1. Check the troubleshooting section above
2. Review logs for error details
3. Test MCP server independently
4. Verify OpenAI API access
5. Check Bot Framework Emulator connection

---

**Built with ❤️ using Microsoft Teams AI Library and Model Context Protocol**
