import { Client } from '@microsoft/microsoft-graph-client';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

// Get current directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

// Parse CLI args: --top=20, --next-link="url", --fetch-all
function parseArgs() {
  const args = process.argv.slice(2);
  const opts = { top: null, nextLink: null, fetchAll: false };
  for (const arg of args) {
    if (arg === '--fetch-all') opts.fetchAll = true;
    else if (arg.startsWith('--top=')) opts.top = parseInt(arg.slice(6), 10);
    else if (arg.startsWith('--next-link=')) opts.nextLink = arg.slice(12).replace(/^["']|["']$/g, '');
  }
  return opts;
}

async function listPages() {
  const { top, nextLink, fetchAll } = parseArgs();

  try {
    // Read the access token
    if (!fs.existsSync(tokenFilePath)) {
      console.error('Access token not found. Please authenticate first.');
      return;
    }

    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
    let accessToken;

    try {
      // Try to parse as JSON first (new format)
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
    } catch (parseError) {
      // Fall back to using the raw token (old format)
      accessToken = tokenData;
    }

    if (!accessToken) {
      console.error('Access token not found in file.');
      return;
    }

    // Create Microsoft Graph client
    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    let pagesResponse;
    let sectionDisplayName = '';

    if (nextLink) {
      // Fetch next page using the link from a previous response
      console.log('Fetching next page...');
      pagesResponse = await client.api(nextLink).get();
    } else {
      // First, get all notebooks
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

      if (sectionsResponse.value.length === 0) {
        console.log('No sections found in this notebook.');
        return;
      }

      // Use the first section (you can modify this to select a specific section)
      const section = sectionsResponse.value[0];
      sectionDisplayName = section.displayName;
      console.log(`Using section: "${section.displayName}"`);

      // Get pages in the section
      let url = `/me/onenote/sections/${section.id}/pages`;
      if (top != null && top > 0) {
        url += `?$top=${Math.min(Math.floor(top), 999)}`;
      }
      console.log(`Fetching pages in "${section.displayName}" section...`);
      pagesResponse = await client.api(url).get();
    }

    let allPages = [...(pagesResponse.value || [])];

    if (fetchAll && pagesResponse['@odata.nextLink']) {
      let response = pagesResponse;
      while (response['@odata.nextLink']) {
        response = await client.api(response['@odata.nextLink']).get();
        allPages = allPages.concat(response.value || []);
      }
    }

    // console.log(pagesResponse);
    console.log(`\nPages${sectionDisplayName ? ` in ${sectionDisplayName}` : ''}:`);
    console.log('=====================');

    if (allPages.length === 0) {
      console.log('No pages found.');
    } else {
      allPages.forEach((page, index) => {
        console.log(`${index + 1}. ${page.title} (Created: ${new Date(page.createdDateTime).toLocaleString()})`);
      });
    }

    const next = fetchAll ? null : (pagesResponse['@odata.nextLink'] || null);
    if (next) {
      console.log('\nMore results available. Run with:');
      console.log(`  node list-pages.js --next-link="${next}"`);
    }
  } catch (error) {
    console.error('Error listing pages:', error);
  }
}

// Run the function
listPages();