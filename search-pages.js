#!/usr/bin/env node
// Usage: node search-pages.js [keyword]
//        node search-pages.js [keyword] --section-id=<sectionId>
//        node search-pages.js
// Uses .access-token.txt (same as MCP). No keyword = list all pages.

import { Client } from '@microsoft/microsoft-graph-client';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const tokenFilePath = path.join(__dirname, '.access-token.txt');

function parseArgs() {
  const args = process.argv.slice(2);
  const opts = { searchTerm: null, sectionId: null };
  for (const arg of args) {
    if (arg.startsWith('--section-id=')) {
      opts.sectionId = arg.slice(13).replace(/^["']|["']$/g, '');
    } else if (!arg.startsWith('--')) {
      opts.searchTerm = arg;
    }
  }
  return opts;
}

async function searchPages() {
  const { searchTerm, sectionId } = parseArgs();
  const hasSearchTerm = searchTerm && searchTerm.trim().length > 0;

  try {
    if (!fs.existsSync(tokenFilePath)) {
      console.error('Access token not found. Please authenticate first (e.g. run the MCP server and use saveAccessToken).');
      process.exit(1);
    }

    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
    let accessToken;
    try {
      const parsed = JSON.parse(tokenData);
      accessToken = parsed.token;
    } catch {
      accessToken = tokenData;
    }

    if (!accessToken) {
      console.error('Access token not found in file.');
      process.exit(1);
    }

    const client = Client.init({
      authProvider: (done) => done(null, accessToken),
    });

    const apiUrl = sectionId
      ? `/me/onenote/sections/${sectionId}/pages`
      : '/me/onenote/pages';
    let response = await client.api(apiUrl).get();
    let pages = response.value || [];

    while (response['@odata.nextLink']) {
      response = await client.api(response['@odata.nextLink']).get();
      pages = pages.concat(response.value || []);
    }

    if (hasSearchTerm) {
      const term = searchTerm.trim().toLowerCase();
      pages = pages.filter((p) => p.title && p.title.toLowerCase().includes(term));
    }

    if (hasSearchTerm) {
      console.log(`Search: "${searchTerm}"${sectionId ? ` (section ${sectionId})` : ''}`);
    } else {
      console.log('All pages' + (sectionId ? ` in section ${sectionId}` : '') + ':');
    }
    console.log('=====================');

    if (pages.length === 0) {
      console.log('No pages found.');
    } else {
      pages.forEach((page, i) => {
        const modified = page.lastModifiedDateTime
          ? new Date(page.lastModifiedDateTime).toLocaleString()
          : '—';
        console.log(`${i + 1}. ${page.title || '(untitled)'}`);
        console.log(`   id: ${page.id}  lastModified: ${modified}`);
      });
      console.log(`\nTotal: ${pages.length} page(s)`);
    }
  } catch (error) {
    console.error('Error searching pages:', error);
    process.exit(1);
  }
}

searchPages();
