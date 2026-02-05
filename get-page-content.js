import { Client } from '@microsoft/microsoft-graph-client';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

// Get current directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

// Get the page ID from command line
const pageId = process.argv[2];
if (!pageId) {
  console.error('Please provide a page ID as argument. Example: node get-page-content.js "1-abc123..."');
  process.exit(1);
}

// Function to check if page should be skipped
function shouldSkipPage(page) {
  const title = (page.title || '').toLowerCase();

  // Check if title contains "private" or "(old)"
  if (title.includes('private') || title.includes('(old)')) {
    console.log(`Skipping page: Title contains excluded keyword: "${page.title}"`);
    return true;
  }

  // Check if last modified date is older than 2022/1/1
  if (page.lastModifiedDateTime) {
    const lastModified = new Date(page.lastModifiedDateTime);
    const cutoffDate = new Date('2022-01-01');
    if (lastModified < cutoffDate) {
      console.log(`Skipping page: Last modified date (${lastModified.toISOString()}) is older than 2022/1/1`);
      return true;
    }
  }

  return false;
}

async function getPageContent() {
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

    // Fetch page metadata to get the title
    console.log(`Fetching page with ID: "${pageId}"...`);
    const page = await client.api(`/me/onenote/pages/${pageId}`).get();
    console.log(`Found page: "${page.title}" (ID: ${page.id})`);

    // Check if page should be skipped
    if (shouldSkipPage(page)) {
      console.log('Page will not be downloaded due to filtering rules.');
      return;
    }

    // Fetch the content using the /content endpoint
    console.log('\nFetching page content...');
    const content = await client.api(`/me/onenote/pages/${pageId}/content`)
      .header('Accept', 'text/html')
      .get();

    console.log(`Content received! Length: ${typeof content === 'string' ? content.length : JSON.stringify(content).length} characters`);

    // Save full content to file
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const safeFileName = page.title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    const outputPath = path.join(outputDir, `${safeFileName}_${pageId.substring(0, 8)}.html`);

    const contentString = typeof content === 'string' ? content : JSON.stringify(content, null, 2);
    fs.writeFileSync(outputPath, contentString, 'utf8');
    console.log(`Full content saved to: ${outputPath}`);

    // Extract text content for console snippet
    let plainText = contentString
      .replace(/<[^>]*>?/gm, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    // Log snippet to console (first 500 characters)
    const snippet = plainText.length > 500 ? plainText.substring(0, 500) + '...' : plainText;
    console.log('\n--- PAGE CONTENT SNIPPET ---\n');
    console.log(snippet);
    console.log('\n--- END OF SNIPPET ---\n');

  } catch (error) {
    console.error('Error:', error);
  }
}

getPageContent();
