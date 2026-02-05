#!/usr/bin/env node

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { Client } from '@microsoft/microsoft-graph-client';
import fetch from 'node-fetch';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Configuration
const PAGE_LIST_FILE = '/Users/lewisanderson/Library/CloudStorage/GoogleDrive-lewis@amaru.ai/My Drive/Private - Lewis - Amaru/OneNoteExport/page-list-index.txt';
const OUTPUT_DIR = '/Users/lewisanderson/Library/CloudStorage/GoogleDrive-lewis@amaru.ai/My Drive/Private - Lewis - Amaru/OneNoteExport/Pages/';
const FAILED_DOWNLOADS_FILE = '/Users/lewisanderson/Library/CloudStorage/GoogleDrive-lewis@amaru.ai/My Drive/Private - Lewis - Amaru/OneNoteExport/failed-downloads.txt';
const MAX_CONSECUTIVE_FAILURES = 10;
const REQUEST_TIMEOUT_MS = 60_000;
const FAST_MODE_DAYS = 30; // Consider file "recent" if within this many days

// Read access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');
let accessToken = null;
try {
  if (fs.existsSync(tokenFilePath)) {
    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
    try {
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
    } catch (parseError) {
      accessToken = tokenData;
    }
  }
} catch (error) {
  console.error('Error reading access token file:', error.message);
  process.exit(1);
}

if (!accessToken) {
  console.error('Access token not found. Please run authenticate first.');
  process.exit(1);
}

// Create Graph client
const graphClient = Client.init({
  authProvider: (done) => {
    done(null, accessToken);
  },
  defaultTimeout: 180000
});

// Helper function to check if page should be skipped based on title only
function shouldSkipPageByTitle(title) {
  const lowerTitle = title.toLowerCase();
  return lowerTitle.includes('(old)') || lowerTitle.includes('private');
}

// Helper function to check if page should be skipped by last modified date
function shouldSkipByDate(page) {
  if (page.lastModifiedDateTime) {
    const lastModified = new Date(page.lastModifiedDateTime);
    const cutoffDate = new Date('2022-01-01');
    return lastModified < cutoffDate;
  }
  return false;
}

// Format date as YYYYMMDD
function formatDate(dateString) {
  const date = new Date(dateString);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

// Sanitize filename to remove invalid characters
function sanitizeFilename(filename) {
  return filename.replace(/[<>:"/\\|?*]/g, '_');
}

// Check if a recent file exists for this page title (for fast mode)
function hasRecentFileByTitle(pageTitle, fastModeDays = FAST_MODE_DAYS) {
  if (!fs.existsSync(OUTPUT_DIR)) {
    return { found: false };
  }

  const sanitizedTitle = sanitizeFilename(pageTitle);
  const files = fs.readdirSync(OUTPUT_DIR);
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - fastModeDays);

  // Look for files matching pattern: {sanitizedTitle}--YYYYMMDD.html
  const pattern = new RegExp(`^${sanitizedTitle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}--\\d{8}\\.html$`);

  for (const file of files) {
    if (pattern.test(file)) {
      // Extract date from filename
      const dateMatch = file.match(/--(\d{8})\.html$/);
      if (dateMatch) {
        const dateStr = dateMatch[1];
        const year = parseInt(dateStr.substring(0, 4));
        const month = parseInt(dateStr.substring(4, 6)) - 1;
        const day = parseInt(dateStr.substring(6, 8));
        const fileDate = new Date(year, month, day);

        if (fileDate >= cutoffDate) {
          return { found: true, filename: file, date: fileDate };
        }
      }
    }
  }

  return { found: false };
}

// Download a single page
async function downloadPage(pageId, pageTitle, options = {}) {
  const { fastMode = false, retryCount = 0 } = options;

  try {
    // In fast mode, check if a recent file exists before fetching metadata
    if (fastMode) {
      const recentFile = hasRecentFileByTitle(pageTitle);
      if (recentFile.found) {
        console.log(`  ⏭️  Fast mode: Recent file exists (${recentFile.filename})`);
        return { success: true, skipped: true, reason: 'fast mode - recent file exists' };
      }
    }

    console.log(`Fetching metadata for: ${pageTitle}`);

    // Get page metadata
    const page = await graphClient.api(`/me/onenote/pages/${pageId}`).get();

    // Check if should skip by date
    if (shouldSkipByDate(page)) {
      console.log(`  ⏭️  Skipped (last modified before 2022-01-01): ${pageTitle}`);
      return { success: true, skipped: true, reason: 'old date' };
    }

    // Create filename with sanitized title and formatted date
    const formattedDate = formatDate(page.lastModifiedDateTime);
    const sanitizedTitle = sanitizeFilename(pageTitle);
    const filename = `${sanitizedTitle}--${formattedDate}.html`;
    const filePath = path.join(OUTPUT_DIR, filename);

    // Check if file already exists
    if (fs.existsSync(filePath)) {
      console.log(`  ⏭️  Already exists: ${filename}`);
      return { success: true, skipped: true, reason: 'already exists' };
    }

    // Download page content
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
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

    // Ensure output directory exists
    if (!fs.existsSync(OUTPUT_DIR)) {
      fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    }

    // Save to file
    fs.writeFileSync(filePath, content, 'utf8');
    console.log(`  ✅ Downloaded: ${filename} (${content.length} chars)`);

    return { success: true, skipped: false, filename };
  } catch (error) {
    if (retryCount < 1) {
      console.log(`  ⚠️  Error, retrying: ${error.message}`);
      return await downloadPage(pageId, pageTitle, { fastMode, retryCount: retryCount + 1 });
    } else {
      console.log(`  ❌ Failed after retry: ${error.message}`);
      return { success: false, error: error.message };
    }
  }
}

// Main function
async function main() {
  // Parse command line arguments
  const args = process.argv.slice(2);
  const fastMode = args.includes('--fast');

  if (fastMode) {
    console.log(`⚡ Fast mode enabled: Skipping metadata fetch for pages with recent files (within ${FAST_MODE_DAYS} days)`);
    console.log('');
  }

  console.log('Reading page list...');
  const fileContent = fs.readFileSync(PAGE_LIST_FILE, 'utf8');
  const lines = fileContent.split('\n');

  // Parse page entries (format: "123. Title (Created: date) -- pageId")
  const pages = [];
  for (const line of lines) {
    const match = line.match(/^\d+\.\s+(.+?)\s+\(Created:.*?\)\s+--\s+(.+)$/);
    if (match) {
      const title = match[1].trim();
      const pageId = match[2].trim();
      pages.push({ title, pageId });
    }
  }

  console.log(`Found ${pages.length} pages in list`);

  // Filter out pages with "(old)" or "private" in title
  const pagesToDownload = pages.filter(p => !shouldSkipPageByTitle(p.title));
  const skippedByTitle = pages.length - pagesToDownload.length;

  console.log(`Skipped ${skippedByTitle} pages by title (old/private)`);
  console.log(`Will attempt to download ${pagesToDownload.length} pages`);
  console.log('');

  // Download pages
  let consecutiveFailures = 0;
  let totalSuccess = 0;
  let totalSkipped = 0;
  let totalFailed = 0;
  const failedPages = [];

  for (let i = 0; i < pagesToDownload.length; i++) {
    const { title, pageId } = pagesToDownload[i];
    console.log(`[${i + 1}/${pagesToDownload.length}] ${title}`);

    const result = await downloadPage(pageId, title, { fastMode });

    if (result.success) {
      consecutiveFailures = 0;
      if (result.skipped) {
        totalSkipped++;
      } else {
        totalSuccess++;
      }
    } else {
      consecutiveFailures++;
      totalFailed++;
      failedPages.push({ title, pageId, error: result.error });

      if (consecutiveFailures >= MAX_CONSECUTIVE_FAILURES) {
        console.log('');
        console.log(`❌ Stopping due to ${MAX_CONSECUTIVE_FAILURES} consecutive failures`);
        break;
      }
    }

    // Small delay to avoid rate limiting
    await new Promise(resolve => setTimeout(resolve, 100));
  }

  // Write failed downloads to file
  if (failedPages.length > 0) {
    const failedContent = failedPages.map(p => `${p.title} -- ${p.pageId} (Error: ${p.error})`).join('\n');
    fs.writeFileSync(FAILED_DOWNLOADS_FILE, failedContent, 'utf8');
    console.log('');
    console.log(`Failed downloads logged to: ${FAILED_DOWNLOADS_FILE}`);
  }

  // Summary
  console.log('');
  console.log('=== Summary ===');
  console.log(`Total pages in list: ${pages.length}`);
  console.log(`Skipped by title (old/private): ${skippedByTitle}`);
  console.log(`Skipped by date (before 2022-01-01): ${totalSkipped}`);
  console.log(`Successfully downloaded: ${totalSuccess}`);
  console.log(`Failed: ${totalFailed}`);
}

main().catch(error => {
  console.error('Fatal error:', error);
  process.exit(1);
});
