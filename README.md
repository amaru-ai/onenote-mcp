# OneNote MCP Server

A Model Context Protocol (MCP) server implementation that enables AI language models like Claude and other LLMs to interact with Microsoft OneNote.

> This project is based on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks, with modifications to simplify authentication and improve usability.

## What Does This Do?

This server allows AI assistants to:
- Access your OneNote notebooks, sections, and pages
- Create new pages in your notebooks
- Search through your notes
- Read complete note content, including HTML formatting and text
- Analyze and summarize your notes directly

All of this happens directly through the AI interface without you having to switch contexts.

## Using with AI Assistants

### Setup for Cursor

1. Clone this repository and follow the installation steps below
2. Start the MCP server: `npm start`
3. Register the server in Cursor:
   - Open Cursor preferences (Cmd+, on Mac or Ctrl+, on Windows)
   - Go to the "MCP" tab
   - Add a new MCP server with these settings:
     - Name: `onenote` 
     - Command: `node`
     - Args: `["/path/to/your/onenote-mcp.mjs"]` (use absolute path)

   Here's the complete JSON configuration example:
   ```json
   {
     "mcpServers": {
       "onenote": {
         "command": "node",
         "args": ["/absolute/path/to/your/onenote-mcp.mjs"],
         "env": {}
       }
     }
   }
   ```
   
4. Restart Cursor
5. In Cursor, you can now interact with your OneNote data using natural language:

```
Can you show me my OneNote notebooks?
Create a new page in my first notebook with a summary of this conversation
Find notes related to "project planning" in my OneNote
```

The first time you ask about OneNote, the AI will guide you through the authentication process.

### Setup for Claude Desktop (or other MCP-compatible assistants)

1. Clone this repository and follow the installation steps below
2. Start the MCP server: `npm start`
3. In the Claude Desktop settings, add the OneNote MCP server:
   - Name: `onenote`
   - Command: `node`
   - Args: `["/path/to/your/onenote-mcp.mjs"]` (use absolute path)
   
   JSON configuration example:
   ```json
   {
     "mcpServers": {
       "onenote": {
         "command": "node",
         "args": ["/absolute/path/to/your/onenote-mcp.mjs"],
         "env": {}
       }
     }
   }
   ```
   
4. You can now ask Claude to interact with your OneNote data

### Configuration

- **`ONENOTE_REQUEST_TIMEOUT_MS`** – Timeout in milliseconds for outgoing requests (e.g. fetching page content). Default is 60000 (60 seconds). Increase if you hit timeouts on slow or large pages. Example in Cursor MCP config: `"env": { "ONENOTE_REQUEST_TIMEOUT_MS": "120000" }`.

## Features

- Authentication with Microsoft OneNote using device code flow (no Azure setup needed)
- List all notebooks, sections, and pages
- Create new pages with HTML content
- Read complete page content, including HTML formatting
- Extract text content for AI analysis and summaries
- Summarize content of all pages in a single operation
- Read full content of all pages in a readable format
- Search across your notes

## Installation

### Prerequisites

- Node.js 16 or higher (install from [nodejs.org](https://nodejs.org/))
- An active Microsoft account with access to OneNote
- Git (install from [git-scm.com](https://git-scm.com/))

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/onenote-mcp.git
cd onenote-mcp
```

### Step 2: Download the TypeScript SDK

This project requires the MCP TypeScript SDK, which needs to be downloaded separately:

```bash
git clone https://github.com/modelcontextprotocol/typescript-sdk.git
cd typescript-sdk
npm install
npm run build
cd ..
```

### Step 3: Install Project Dependencies

```bash
npm install
```

### Step 4: Start the MCP Server

```bash
npm start
```

This will start the MCP server, and you'll see a message:
```
Server started successfully.
Use the "authenticate" tool to start the authentication flow,
or use "saveAccessToken" if you already have a token.
```

### Step 5: Authenticate Through Your AI Assistant

Once the server is running, you can authenticate directly through your AI assistant:

1. In Cursor, Anthropic's Claude Desktop, or any MCP-compatible assistant, ask to authenticate with OneNote:
   ```
   Can you authenticate with my OneNote account?
   ```

2. The AI will trigger the authentication flow and provide you with:
   - A URL (typically microsoft.com/devicelogin)
   - A code to enter

3. Go to the URL, enter the code, and sign in with your Microsoft account

4. After successful authentication, you can start using OneNote with your AI assistant

## Available MCP Tools

Once authenticated, the following tools are available for AI assistants to use:

| Tool Name | Description |
|-----------|-------------|
| `authenticate` | Start the Microsoft authentication flow |
| `listNotebooks` | Get a list of all your OneNote notebooks |
| `getNotebook` | Get details of a specific notebook |
| `listSections` | List all sections in a notebook |
| `listPages` | List all pages in a section |
| `getPage` | Get the complete content of a specific page, including HTML formatting |
| `createPage` | Create a new page with HTML content |
| `searchPages` | Search for pages across your notebooks |

## Example Interactions

Here are some examples of how you can interact with the OneNote MCP through your AI assistant:

```
User: Can you show me my OneNote notebooks?
AI: (uses listNotebooks) I found 3 notebooks: "Work", "Personal", and "Projects"

User: What sections are in my Projects notebook?
AI: (uses listSections) Your Projects notebook has the following sections: "Active Projects", "Ideas", and "Completed"

User: Create a new page in Projects with today's date as the title
AI: (uses createPage) I've created a new page titled "2025-04-12" in your Projects notebook

User: Find all my notes about machine learning
AI: (uses searchPages) I found 5 pages with content related to machine learning...

User: Can you read and summarize my notes on the "Project Requirements" page?
AI: (uses getPage) Based on your "Project Requirements" page, here's a summary: The project requires Python 3.8+, integration with AWS services, and completion by Q3. Key deliverables include a web dashboard, API, and documentation...

User: Extract all the action items from my "Team Meeting" notes
AI: (uses getPage) Here are all the action items from your "Team Meeting" notes:
1. John to complete API documentation by Friday
2. Sarah to schedule design review meeting
3. Team to finalize Q3 roadmap by end of month

User: Summarize content of all my OneNote pages
AI: (runs get-all-page-contents.js) Here's a summary of all your pages:
- Questions: Contains strategic business questions about competitor analysis
- 2025-04-12: Discussion about monetization strategy for bank transfers
- Role Specification: Details about the Chief Payments Officer position
...

User: I want to read through all my OneNote pages so I can ask questions about them
AI: (runs read-all-pages.js) I've retrieved the full content of all your pages in a readable format. Now you can ask me specific questions about any of the content.
```

## Advanced: Direct Script Usage

For testing or development purposes, you can also use the provided scripts directly:

```bash
# Authenticate with Microsoft
npm run auth

# List your notebooks
npm run list-notebooks

# List sections in your first notebook
npm run list-sections

# List pages in the first section
npm run list-pages

# Create a new page
npm run create-page

# Summarize content of all pages
node get-all-page-contents.js

# Read full content of all pages
node read-all-pages.js
```

## Troubleshooting

### Authentication Issues

- If authentication fails, make sure you're using a modern browser without tracking prevention
- Try clearing browser cookies and cache
- If you get "expired_token" errors, restart the authentication process

### Server Won't Start

- Verify Node.js is installed (version 16+): `node --version`
- Make sure all dependencies are installed: `npm install`
- Check that the TypeScript SDK was built correctly

### AI Can't Connect to the Server

- Ensure the MCP server is running (`npm start`)
- Check your AI assistant's settings to make sure it's configured to use MCP
- For Cursor, make sure it's the latest version that supports MCP

## Security Notes

- Authentication tokens are stored locally in `.access-token.txt`
- Tokens grant access to your OneNote data, so keep them secure
- Tokens expire after some time, requiring re-authentication
- No Azure setup or API keys are required

## Credits

This project builds upon the [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks, with a focus on simplifying the authentication process and improving the user experience with AI assistants.

## License

This project is licensed under the MIT License - see the LICENSE file for details
