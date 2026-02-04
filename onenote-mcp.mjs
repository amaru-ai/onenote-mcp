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

    // Create Microsoft Graph client
    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
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
  "List all sections in a notebook",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/sections`).get();
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

// Tool for getting the content of a page
server.tool(
  "getPage",
  "Get the content of a page",
  async (params) => {
    try {
      console.error("GetPage called with params:", params);
      await ensureGraphClient();

      // First, list all pages to find the one we want
      const pagesResponse = await graphClient.api('/me/onenote/pages').get();
      console.error("Got", pagesResponse.value.length, "pages");

      let targetPage;

      // If a page ID is provided, use it to find the page
      if (params.random_string && params.random_string.length > 0) {
        const pageId = params.random_string;
        console.error("Looking for page with ID:", pageId);

        // Look for exact match first
        targetPage = pagesResponse.value.find(p => p.id === pageId);

        // If no exact match, try matching by title
        if (!targetPage) {
          console.error("No exact match, trying title search");
          targetPage = pagesResponse.value.find(p =>
            p.title && p.title.toLowerCase().includes(params.random_string.toLowerCase())
          );
        }

        // If still no match, try partial ID match
        if (!targetPage) {
          console.error("No title match, trying partial ID match");
          targetPage = pagesResponse.value.find(p =>
            p.id.includes(pageId) || pageId.includes(p.id)
          );
        }
      } else {
        // If no ID provided, use the first page
        console.error("No ID provided, using first page");
        targetPage = pagesResponse.value[0];
      }

      if (!targetPage) {
        throw new Error("Page not found");
      }

      console.error("Target page found:", targetPage.title);
      console.error("Page ID:", targetPage.id);

      try {
        const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${targetPage.id}/content`;
        console.error("Fetching content from:", url);

        // Make direct HTTP request with fetch (with configurable timeout)
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
        console.error("Error getting content:", error);

        // Return a simple error message
        return {
          content: [
            {
              type: "text",
              text: `Error retrieving page content: ${error.message}`
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error in getPage:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error in getPage: ${error.message}`
          }
        ]
      };
    }
  }
);

// Tool for creating a new page in a section
server.tool(
  "createPage",
  "Create a new page in a section",
  async (params) => {
    try {
      await ensureGraphClient();
      // Get sections first
      const sectionsResponse = await graphClient.api(`/me/onenote/sections`).get();

      if (sectionsResponse.value.length === 0) {
        throw new Error("No sections found");
      }

      // Use the first section
      const sectionId = sectionsResponse.value[0].id;

      // Create simple HTML content
      const simpleHtml = `
        <!DOCTYPE html>
        <html>
          <head>
            <title>New Page</title>
          </head>
          <body>
            <p>This is a new page created via the Microsoft Graph API</p>
          </body>
        </html>
      `;

      const response = await graphClient
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header("Content-Type", "application/xhtml+xml")
        .post(simpleHtml);

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
  "Search for pages across notebooks",
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

      let pages;

      // If a section ID is provided, get pages from that section only
      if (params.sectionId) {
        const response = await graphClient.api(`/me/onenote/sections/${params.sectionId}/pages`).get();
        pages = response.value;
      } else {
        // Get all pages across all sections
        const response = await graphClient.api(`/me/onenote/pages`).get();
        pages = response.value;
      }

      // If search term is provided, filter the results
      if (params.searchTerm && params.searchTerm.length > 0) {
        const searchTerm = params.searchTerm.toLowerCase();
        const filteredPages = pages.filter(page => {
          // Search in title
          if (page.title && page.title.toLowerCase().includes(searchTerm)) {
            return true;
          }
          return false;
        });

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