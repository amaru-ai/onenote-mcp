#!/usr/bin/env node

import { McpServer } from './typescript-sdk/dist/esm/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { StdioServerTransport } from './typescript-sdk/dist/esm/server/stdio.js';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';
import { DeviceCodeCredential } from '@azure/identity';
import fetch from 'node-fetch';

// Load environment variables
dotenv.config();

// Get the current file's directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

// Request timeout in ms (set ONENOTE_REQUEST_TIMEOUT_MS in env to override; default 60s)
const REQUEST_TIMEOUT_MS = process.env.ONENOTE_REQUEST_TIMEOUT_MS
  ? parseInt(process.env.ONENOTE_REQUEST_TIMEOUT_MS, 10)
  : 60_000;

// Create the MCP server
const server = new McpServer(
  {
    name: "onenote",
    version: "1.0.0",
    description: "OneNote MCP Server"
  },
  {
    capabilities: {
      tools: {
        listChanged: true
      }
    }
  }
);

// Try to read the stored access token
let accessToken = null;
try {
  if (fs.existsSync(tokenFilePath)) {
    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
    try {
      // Try to parse as JSON first (new format)
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
    } catch (parseError) {
      // Fall back to using the raw token (old format)
      accessToken = tokenData;
    }
  }
} catch (error) {
  console.error('Error reading access token file:', error.message);
}

// Alternatively, check if token is in environment variables
if (!accessToken && process.env.GRAPH_ACCESS_TOKEN) {
  accessToken = process.env.GRAPH_ACCESS_TOKEN;
}

let graphClient = null;

// Client ID for Microsoft Graph API access
const clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Microsoft Graph Explorer client ID
const scopes = ['Notes.Read.All', 'Notes.ReadWrite.All', 'User.Read'];

// Function to ensure Graph client is created
async function ensureGraphClient() {
  if (!graphClient) {
    // Read token from file if it exists
    try {
      if (fs.existsSync(tokenFilePath)) {
        const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
        try {
          // Try to parse as JSON first (new format)
          const parsedToken = JSON.parse(tokenData);
          accessToken = parsedToken.token;
        } catch (parseError) {
          // Fall back to using the raw token (old format)
          accessToken = tokenData;
        }
      }
    } catch (error) {
      console.error("Error reading token file:", error);
    }

    if (!accessToken) {
      throw new Error("Access token not found. Please save access token first.");
    }

    // Create Microsoft Graph client with extended timeout
    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
      defaultTimeout: 180000 // 3 minutes (180 seconds) instead of default 30 seconds
    });
  }
  return graphClient;
}

// Create graph client with device code auth or access token
async function createGraphClient() {
  if (accessToken) {
    // Use access token if available
    graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          return accessToken;
        }
      }
    });
    return { type: 'token', client: graphClient };
  } else {
    // Use device code flow
    const credential = new DeviceCodeCredential({
      clientId: clientId,
      userPromptCallback: (info) => {
        // This will be shown to the user with the URL and code
        console.error('\n' + info.message);
      }
    });

    try {
      // Get an access token using device code flow
      const tokenResponse = await credential.getToken(scopes);

      // Save the token for future use
      accessToken = tokenResponse.token;
      fs.writeFileSync(tokenFilePath, JSON.stringify({ token: accessToken }));

      // Initialize Graph client with the token
      graphClient = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            return accessToken;
          }
        }
      });

      return { type: 'device_code', client: graphClient };
    } catch (error) {
      console.error('Authentication error:', error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
}

// Tool for starting authentication flow
server.tool(
  "authenticate",
  "Start the authentication flow with Microsoft Graph",
  async () => {
    try {
      const result = await createGraphClient();
      if (result.type === 'device_code') {
        return {
          content: [
            {
              type: "text",
              text: "Authentication started. Please check the console for the URL and code."
            }
          ]
        };
      } else {
        return {
          content: [
            {
              type: "text",
              text: "Already authenticated with an access token."
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error in authentication:", error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
);

// Tool for saving an access token provided by the user
server.tool(
  "saveAccessToken",
  "Save a Microsoft Graph access token for later use",
  async (params) => {
    try {
      // Save the token for future use
      accessToken = params.random_string;
      const tokenData = JSON.stringify({ token: accessToken });
      fs.writeFileSync(tokenFilePath, tokenData);
      await createGraphClient();
      return {
        content: [
          {
            type: "text",
            text: "Access token saved successfully"
          }
        ]
      };
    } catch (error) {
      console.error("Error saving access token:", error);
      throw new Error(`Failed to save access token: ${error.message}`);
    }
  }
);

// Tool for listing all notebooks
server.tool(
  "listNotebooks",
  "List all OneNote notebooks",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api("/me/onenote/notebooks").get();
      // Return content as an array of text items
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing notebooks:", error);
      throw new Error(`Failed to list notebooks: ${error.message}`);
    }
  }
);

// Tool for getting notebook details
server.tool(
  "getNotebook",
  "Get details of a specific notebook",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/notebooks`).get();
      const notebook = response.value.find(n => n.displayName && n.displayName.toLowerCase().includes("lewis's notebook"));
      if (!notebook) {
        throw new Error('No notebook with "lewis" in the display name found');
      }
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(notebook)
          }
        ]
      };
    } catch (error) {
      console.error("Error getting notebook:", error);
      throw new Error(`Failed to get notebook: ${error.message}`);
    }
  }
);

// Tool for listing sections in a notebook
server.tool(
  "listSections",
  "List all sections in a notebook. If notebookId is provided, lists sections in that notebook. Otherwise, lists all sections.",
  {
    notebookId: {
      type: "string",
      description: "The ID of the notebook to list sections from. If not provided, lists all sections across all notebooks."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient();

      let response;
      if (params.notebookId) {
        // List sections in specific notebook
        console.error(`Fetching sections in notebook: ${params.notebookId}`);
        response = await graphClient.api(`/me/onenote/notebooks/${params.notebookId}/sections`).get();
      } else {
        // List all sections
        console.error("Fetching all sections");
        response = await graphClient.api(`/me/onenote/sections`).get();
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing sections:", error);
      throw new Error(`Failed to list sections: ${error.message}`);
    }
  }
);

// Tool for listing pages in a section
server.tool(
  "listPages",
  "List all pages in a section. Supports pagination: use fetchAll to get every page, or use top + the returned nextLink to get the next page.",
  {
    sectionId: {
      type: "string",
      description: "The ID of the section to list pages from. If not provided, uses the first section found."
    },
    top: {
      type: "number",
      description: "Page size (e.g. 20). Default is API default (~20). Use with nextLink for manual pagination."
    },
    nextLink: {
      type: "string",
      description: "The @odata.nextLink from a previous listPages response to fetch the next page of results."
    },
    fetchAll: {
      type: "boolean",
      description: "If true, follow all pages and return every page in one response. Default false."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient();

      let sectionId = params.sectionId;
      const top = params.top;
      const nextLink = params.nextLink;
      const fetchAll = params.fetchAll === true;

      // If nextLink provided, request that page directly (no sectionId needed)
      if (nextLink) {
        const response = await graphClient.api(nextLink).get();
        const value = response.value || [];
        const hasMore = !!response["@odata.nextLink"];
        const out = { value, nextLink: hasMore ? response["@odata.nextLink"] : undefined };
        return {
          content: [{ type: "text", text: JSON.stringify(out) }]
        };
      }

      // If no section ID provided, get the first section
      if (!sectionId) {
        const sectionsResponse = await graphClient.api(`/me/onenote/sections`).get();

        if (sectionsResponse.value.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({ value: [], nextLink: undefined })
              }
            ]
          };
        }

        sectionId = sectionsResponse.value[0].id;
      }

      let url = `/me/onenote/sections/${sectionId}/pages`;
      if (top != null && top > 0) {
        url += (url.includes("?") ? "&" : "?") + `$top=${Math.min(Math.floor(top), 999)}`;
      }

      let response = await graphClient.api(url).get();
      let allValues = [...(response.value || [])];

      if (fetchAll) {
        while (response["@odata.nextLink"]) {
          response = await graphClient.api(response["@odata.nextLink"]).get();
          allValues = allValues.concat(response.value || []);
        }
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ value: allValues, nextLink: undefined })
            }
          ]
        };
      }

      const hasMore = !!response["@odata.nextLink"];
      const out = {
        value: allValues,
        nextLink: hasMore ? response["@odata.nextLink"] : undefined
      };
      return {
        content: [{ type: "text", text: JSON.stringify(out) }]
      };
    } catch (error) {
      console.error("Error listing pages:", error);
      throw new Error(`Failed to list pages: ${error.message}`);
    }
  }
);

// Helper function to check if page should be skipped
function shouldSkipPage(page) {
  const title = (page.title || '').toLowerCase();

  // Check if title contains "private" or "(old)"
  if (title.includes('private') || title.includes('(old)')) {
    return { skip: true, reason: `Title contains excluded keyword: "${page.title}"` };
  }

  // Check if last modified date is older than 2022/1/1
  if (page.lastModifiedDateTime) {
    const lastModified = new Date(page.lastModifiedDateTime);
    const cutoffDate = new Date('2022-01-01');
    if (lastModified < cutoffDate) {
      return { skip: true, reason: `Last modified date (${lastModified.toISOString()}) is older than 2022/1/1` };
    }
  }

  return { skip: false };
}

// Tool for getting the content of a page
server.tool(
  "getPage",
  "Get the content of a page by page ID. Skips pages with 'private' or '(old)' in title, or pages last modified before 2022/1/1.",
  {
    pageId: {
      type: "string",
      description: "The ID of the page to retrieve content from"
    }
  },
  async (params) => {
    try {
      console.error("GetPage called with params:", params);
      await ensureGraphClient();

      const pageId = params.pageId;
      if (!pageId) {
        throw new Error("Page ID is required");
      }

      console.error("Fetching page with ID:", pageId);

      // First get page metadata to verify it exists and get the title
      const page = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error("Target page found:", page.title);
      console.error("Page ID:", page.id);

      // Check if page should be skipped
      const skipCheck = shouldSkipPage(page);
      if (skipCheck.skip) {
        console.error(`Skipping page: ${skipCheck.reason}`);
        return {
          content: [
            {
              type: "text",
              text: `Page skipped: ${skipCheck.reason}`
            }
          ]
        };
      }

      // Fetch the content using direct HTTP request (with configurable timeout)
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      console.error("Fetching content from:", url);

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        signal: controller.signal
      });
      clearTimeout(timeoutId);

      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status} ${response.statusText}`);
      }

      const content = await response.text();
      console.error(`Content received! Length: ${content.length} characters`);

      // Return the raw HTML content
      return {
        content: [
          {
            type: "text",
            text: content
          }
        ]
      };
    } catch (error) {
      console.error("Error in getPage:", error);
      throw new Error(`Failed to get page: ${error.message}`);
    }
  }
);

// Tool for getting page content using Graph client (alternative method)
server.tool(
  "getPageContent",
  "Get the content of a page using the Graph API client method. Skips pages with 'private' or '(old)' in title, or pages last modified before 2022/1/1.",
  {
    pageId: {
      type: "string",
      description: "The ID of the page to retrieve content from"
    }
  },
  async (params) => {
    try {
      console.error("GetPageContent called with params:", params);
      await ensureGraphClient();

      const pageId = params.pageId;
      if (!pageId) {
        throw new Error("Page ID is required");
      }

      console.error("Fetching page with ID:", pageId);

      // Get page metadata
      const page = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error("Found page:", page.title);

      // Check if page should be skipped
      const skipCheck = shouldSkipPage(page);
      if (skipCheck.skip) {
        console.error(`Skipping page: ${skipCheck.reason}`);
        return {
          content: [
            {
              type: "text",
              text: `Page skipped: ${skipCheck.reason}`
            }
          ]
        };
      }

      // Fetch the content using the /content endpoint with Accept header
      console.error("Fetching page content...");
      const content = await graphClient.api(`/me/onenote/pages/${pageId}/content`)
        .header('Accept', 'text/html')
        .get();

      const contentString = typeof content === 'string' ? content : JSON.stringify(content, null, 2);
      console.error(`Content received! Length: ${contentString.length} characters`);

      // Return the content
      return {
        content: [
          {
            type: "text",
            text: contentString
          }
        ]
      };
    } catch (error) {
      console.error("Error in getPageContent:", error);
      throw new Error(`Failed to get page content: ${error.message}`);
    }
  }
);

// Tool for downloading page content to a file
server.tool(
  "downloadFile",
  "Download page content to a specified file path. Skips pages with 'private' or '(old)' in title, or pages last modified before 2022/1/1.",
  {
    pageId: {
      type: "string",
      description: "The ID of the page to download"
    },
    filePath: {
      type: "string",
      description: "The absolute file path where the page content should be saved"
    }
  },
  async (params) => {
    try {
      console.error("DownloadFile called with params:", params);
      await ensureGraphClient();

      const pageId = params.pageId;
      const filePath = params.filePath;

      if (!pageId) {
        throw new Error("Page ID is required");
      }
      if (!filePath) {
        throw new Error("File path is required");
      }

      console.error("Fetching page with ID:", pageId);

      // First get page metadata to verify it exists and get the title
      const page = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error("Target page found:", page.title);
      console.error("Page ID:", page.id);

      // Check if page should be skipped
      const skipCheck = shouldSkipPage(page);
      if (skipCheck.skip) {
        console.error(`Skipping page: ${skipCheck.reason}`);
        return {
          content: [
            {
              type: "text",
              text: `Page skipped: ${skipCheck.reason}`
            }
          ]
        };
      }

      // Fetch the content using direct HTTP request (with configurable timeout)
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      console.error("Fetching content from:", url);

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        signal: controller.signal
      });
      clearTimeout(timeoutId);

      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status} ${response.statusText}`);
      }

      const content = await response.text();
      console.error(`Content received! Length: ${content.length} characters`);

      // Save to file
      fs.writeFileSync(filePath, content, 'utf8');
      console.error(`Content saved to: ${filePath}`);

      // Return success message
      return {
        content: [
          {
            type: "text",
            text: `Page "${page.title}" successfully downloaded to ${filePath} (${content.length} characters)`
          }
        ]
      };
    } catch (error) {
      console.error("Error in downloadFile:", error);
      throw new Error(`Failed to download file: ${error.message}`);
    }
  }
);

// Tool for creating a new page in a section
server.tool(
  "createPage",
  "Create a new page in a section with specified title and content",
  {
    sectionId: {
      type: "string",
      description: "The ID of the section to create the page in. If not provided, uses the first section found."
    },
    title: {
      type: "string",
      description: "The title of the new page. Defaults to 'New Page'."
    },
    content: {
      type: "string",
      description: "The HTML body content of the page. If not provided, creates a simple default page."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient();

      let sectionId = params.sectionId;

      // If no section ID provided, use the first section
      if (!sectionId) {
        const sectionsResponse = await graphClient.api(`/me/onenote/sections`).get();
        if (sectionsResponse.value.length === 0) {
          throw new Error("No sections found");
        }
        sectionId = sectionsResponse.value[0].id;
        console.error(`Using first section: ${sectionsResponse.value[0].displayName}`);
      }

      const title = params.title || "New Page";
      const now = new Date();
      const formattedDate = now.toISOString().split('T')[0];
      const formattedTime = now.toLocaleTimeString();

      // Use provided content or create default content
      const bodyContent = params.content || `
        <h1>${title}</h1>
        <p>This page was created via the Microsoft Graph API at ${formattedTime} on ${formattedDate}.</p>
      `;

      // Create HTML content
      const htmlContent = `
        <!DOCTYPE html>
        <html>
          <head>
            <title>${title}</title>
          </head>
          <body>
            ${bodyContent}
          </body>
        </html>
      `;

      console.error(`Creating page "${title}" in section ${sectionId}`);

      const response = await graphClient
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header("Content-Type", "application/xhtml+xml")
        .post(htmlContent);

      console.error(`Page created successfully: ${response.title}`);

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response)
          }
        ]
      };
    } catch (error) {
      console.error("Error creating page:", error);
      throw new Error(`Failed to create page: ${error.message}`);
    }
  }
);

// Tool for searching pages
server.tool(
  "searchPages",
  "Search for pages by title across all sections or within a specific section. Fetches all pages automatically.",
  {
    searchTerm: {
      type: "string",
      description: "The search term to filter pages by title. If not provided, returns all pages."
    },
    sectionId: {
      type: "string",
      description: "The ID of the section to search within. If not provided, searches all sections."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient();

      const apiUrl = params.sectionId
        ? `/me/onenote/sections/${params.sectionId}/pages`
        : '/me/onenote/pages';

      console.error(`Searching pages${params.sectionId ? ` in section ${params.sectionId}` : ' across all sections'}`);

      // Get all pages by following pagination links
      let response = await graphClient.api(apiUrl).get();
      let pages = response.value || [];

      // Follow all nextLinks to get all pages
      while (response['@odata.nextLink']) {
        console.error(`Fetching next page of results...`);
        response = await graphClient.api(response['@odata.nextLink']).get();
        pages = pages.concat(response.value || []);
      }

      console.error(`Retrieved ${pages.length} total pages`);

      // If search term is provided, filter the results
      if (params.searchTerm && params.searchTerm.trim().length > 0) {
        const searchTerm = params.searchTerm.toLowerCase();
        const filteredPages = pages.filter(page => {
          // Search in title
          return page.title && page.title.toLowerCase().includes(searchTerm);
        });

        console.error(`Filtered to ${filteredPages.length} pages matching "${params.searchTerm}"`);

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(filteredPages)
            }
          ]
        };
      } else {
        // Return all pages if no search term
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(pages)
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error searching pages:", error);
      throw new Error(`Failed to search pages: ${error.message}`);
    }
  }
);

// Connect to stdio and start server
async function main() {
  try {
    // Connect to standard I/O
    const transport = new StdioServerTransport();
    await server.connect(transport);

    console.error('Server started successfully.');
    console.error('Use the "authenticate" tool to start the authentication flow,');
    console.error('or use "saveAccessToken" if you already have a token.');

    // Keep the process alive
    process.on('SIGINT', () => {
      process.exit(0);
    });
  } catch (error) {
    console.error('Error starting server:', error);
    process.exit(1);
  }
}

main();