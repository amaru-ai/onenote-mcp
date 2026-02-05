import { Client } from '@microsoft/microsoft-graph-client';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

// Get current directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

async function listSections() {
  try {
    // Read the access token
    if (!fs.existsSync(tokenFilePath)) {
      console.error('Access token not found. Please authenticate first.');
      return;
    }

    const tokenData = fs.readFileSync(tokenFilePath, 'utf8').trim();
    let accessToken;

    try {
      // Try to parse as JSON first (new format: { "token": "..." })
      const parsedToken = JSON.parse(tokenData);
      accessToken = (parsedToken.token || parsedToken.access_token || '').trim();
    } catch (parseError) {
      // Fall back to using the raw token (old format: plain JWT on one line)
      accessToken = tokenData;
    }

    if (!accessToken) {
      console.error('Access token not found in file.');
      return;
    }

    // // JWT format is three base64 parts separated by dots
    // if (!/^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(accessToken)) {
    //   console.error('Invalid token format: token in .access-token.txt is not a valid JWT (expected three dot-separated parts).');
    //   console.error('Re-authenticate and save a new token, e.g. run "node authenticate.js" or use the MCP saveAccessToken tool.');
    //   return;
    // }

    // Create Microsoft Graph client
    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    // First, let's get all notebooks
    console.log('Fetching notebooks...');
    const notebooksResponse = await client.api('/me/onenote/notebooks').get();

    if (notebooksResponse.value.length === 0) {
      console.log('No notebooks found.');
      return;
    }

    // Use notebook with "lewis" in the display name
    const notebook = notebooksResponse.value.find(n => n.displayName && n.displayName.toLowerCase().includes("lewis's notebook"));
    if (!notebook) {
      console.log('No notebook with "lewis" in the display name found.');
      return;
    }
    console.log(`Using notebook: "${notebook.displayName}"`);

    // Get sections in the selected notebook
    console.log(`Fetching sections in "${notebook.displayName}" notebook...`);
    const sectionsResponse = await client.api(`/me/onenote/notebooks/${notebook.id}/sections`).get();

    console.log(`\nSections in ${notebook.displayName} Notebook:`);
    console.log('============================');

    if (sectionsResponse.value.length === 0) {
      console.log('No sections found.');
    } else {
      sectionsResponse.value.forEach((section, index) => {
        console.log(`${index + 1}. ${section.displayName} -- ${section.id}`);
      });
    }

  } catch (error) {
    console.error('Error listing sections:', error);
  }
}

// Run the function
listSections();