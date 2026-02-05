import { Client } from '@microsoft/microsoft-graph-client';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

// Get current directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

// Parse CLI args: --top=20, --next-link="url", --fetch-all, --section-id="id"
function parseArgs() {
  const args = process.argv.slice(2);
  const opts = { top: null, nextLink: null, fetchAll: false, sectionId: null };
  for (const arg of args) {
    if (arg === '--fetch-all') opts.fetchAll = true;
    else if (arg.startsWith('--top=')) opts.top = parseInt(arg.slice(6), 10);
    else if (arg.startsWith('--next-link=')) opts.nextLink = arg.slice(12).replace(/^["']|["']$/g, '');
    else if (arg.startsWith('--section-id=')) opts.sectionId = arg.slice(13).replace(/^["']|["']$/g, '');
  }
  return opts;
}

async function listPages() {
  const { top, nextLink, fetchAll, sectionId } = parseArgs();

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

    // Create Microsoft Graph client with extended timeout
    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
      defaultTimeout: 180000 // 3 minutes (180 seconds) instead of default 30 seconds
    });

    if (nextLink) {
      // Fetch next page using the link from a previous response (single-section pagination)
      console.log('Fetching next page...');
      const pagesResponse = await client.api(nextLink).get();
      let allPages = [...(pagesResponse.value || [])];

      if (fetchAll && pagesResponse['@odata.nextLink']) {
        let response = pagesResponse;
        let batchNum = 1;
        console.log(`Batch 1: Fetched ${allPages.length} pages`);

        while (response['@odata.nextLink']) {
          batchNum++;
          try {
            // Small delay to avoid rate limiting
            await new Promise(resolve => setTimeout(resolve, 100));

            response = await client.api(response['@odata.nextLink']).get();
            const newPages = response.value || [];
            allPages = allPages.concat(newPages);
            console.log(`Batch ${batchNum}: Fetched ${newPages.length} pages (total: ${allPages.length})`);
          } catch (error) {
            console.error(`Error fetching batch ${batchNum}:`, error.message);
            console.log(`Continuing with ${allPages.length} pages fetched so far...`);
            break;
          }
        }
        console.log(`\nCompleted! Fetched ${allPages.length} total pages in ${batchNum} batches`);
      }

      console.log('\nPages:');
      console.log('=====================');
      if (allPages.length === 0) {
        console.log('No pages found.');
      } else {
        allPages.forEach((page, index) => {
          console.log(`${index + 1}. ${page.title} (Created: ${new Date(page.createdDateTime).toLocaleString()}) -- ${page.id}`);
        });
      }
      const next = fetchAll ? null : (pagesResponse['@odata.nextLink'] || null);
      if (next) {
        console.log('\nMore results available. Run with:');
        console.log(`  node list-pages.js --next-link="${next}"`);
      }
      return;
    }

    // If section ID is provided, fetch pages for that section only
    if (sectionId) {
      console.log(`Fetching section information for ID: ${sectionId}...`);
      const section = await client.api(`/me/onenote/sections/${sectionId}`).get();
      console.log(`\n${section.displayName} -- ${section.id}\n`);

      let url = `/me/onenote/sections/${sectionId}/pages`;
      if (top != null && top > 0) {
        url += `?$top=${Math.min(Math.floor(top), 999)}`;
      }
      let pagesResponse = await client.api(url).get();
      let allPages = [...(pagesResponse.value || [])];

      if (fetchAll && pagesResponse['@odata.nextLink']) {
        let response = pagesResponse;
        let batchNum = 1;
        console.log(`Batch 1: Fetched ${allPages.length} pages`);

        while (response['@odata.nextLink']) {
          batchNum++;
          try {
            // Small delay to avoid rate limiting
            await new Promise(resolve => setTimeout(resolve, 100));

            response = await client.api(response['@odata.nextLink']).get();
            const newPages = response.value || [];
            allPages = allPages.concat(newPages);
            console.log(`Batch ${batchNum}: Fetched ${newPages.length} pages (total: ${allPages.length})`);
          } catch (error) {
            console.error(`Error fetching batch ${batchNum}:`, error.message);
            console.log(`Continuing with ${allPages.length} pages fetched so far...`);
            break;
          }
        }
        console.log(`\nCompleted! Fetched ${allPages.length} total pages in ${batchNum} batches`);
      }

      console.log('\nPages:');
      console.log('=====================');
      if (allPages.length === 0) {
        console.log('No pages found.');
      } else {
        allPages.forEach((page, index) => {
          console.log(`${index + 1}. ${page.title} (Created: ${new Date(page.createdDateTime).toLocaleString()}) -- ${page.id}`);
        });
      }

      const next = fetchAll ? null : (pagesResponse['@odata.nextLink'] || null);
      if (next) {
        console.log('\nMore results available. Run with:');
        console.log(`  node list-pages.js --section-id="${sectionId}" --next-link="${next}"`);
      }
      return;
    }

    // No nextLink or sectionId: get notebook and iterate all sections
    console.log('Fetching notebooks...');
    const notebooksResponse = await client.api('/me/onenote/notebooks').get();

    if (notebooksResponse.value.length === 0) {
      console.log('No notebooks found.');
      return;
    }

    const notebook = notebooksResponse.value.find(n => n.displayName && n.displayName.toLowerCase().includes("lewis's notebook"));
    if (!notebook) {
      console.log('No notebook with "lewis" in the display name found.');
      return;
    }
    console.log(`Using notebook: "${notebook.displayName}"`);

    console.log(`Fetching sections in "${notebook.displayName}" notebook...`);
    const sectionsResponse = await client.api(`/me/onenote/notebooks/${notebook.id}/sections`).get();

    if (sectionsResponse.value.length === 0) {
      console.log('No sections found in this notebook.');
      return;
    }

    const sections = sectionsResponse.value;

    for (const section of sections) {
      console.log(`\n\n${section.displayName} -- ${section.id}`);

      let url = `/me/onenote/sections/${section.id}/pages`;
      if (top != null && top > 0) {
        url += `?$top=${Math.min(Math.floor(top), 999)}`;
      }
      let pagesResponse = await client.api(url).get();
      let allPages = [...(pagesResponse.value || [])];

      if (fetchAll && pagesResponse['@odata.nextLink']) {
        let response = pagesResponse;
        let batchNum = 1;
        console.log(`  Batch 1: Fetched ${allPages.length} pages`);

        while (response['@odata.nextLink']) {
          batchNum++;
          try {
            // Small delay to avoid rate limiting
            await new Promise(resolve => setTimeout(resolve, 100));

            response = await client.api(response['@odata.nextLink']).get();
            const newPages = response.value || [];
            allPages = allPages.concat(newPages);
            console.log(`  Batch ${batchNum}: Fetched ${newPages.length} pages (total: ${allPages.length})`);
          } catch (error) {
            console.error(`  Error fetching batch ${batchNum}:`, error.message);
            console.log(`  Continuing with ${allPages.length} pages fetched so far...`);
            break;
          }
        }
        if (batchNum > 1) {
          console.log(`  Completed! Fetched ${allPages.length} total pages in ${batchNum} batches`);
        }
      }

      if (allPages.length === 0) {
        console.log('No pages found.');
      } else {
        allPages.forEach((page, index) => {
          console.log(`${index + 1}. ${page.title} (Created: ${new Date(page.createdDateTime).toLocaleString()}) -- ${page.id}`);
        });
      }

      const next = fetchAll ? null : (pagesResponse['@odata.nextLink'] || null);
      if (next) {
        console.log('\nMore results available. Run with:');
        console.log(`  node list-pages.js --next-link="${next}"`);
      }
    }
  } catch (error) {
    console.error('Error listing pages:', error);
  }
}

// Run the function
listPages();