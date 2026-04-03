const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const SESSION_FILE = path.join(__dirname, 'notebooklm_session.json');

const YOUTUBE_LINKS = [
  'https://www.youtube.com/watch?v=Kxc8KceOb14',  // Topic1: Evolution of HRM
  'https://www.youtube.com/watch?v=8ReX2poQyJ0',  // Topic1: Transformation Personnel->HRM
  'https://www.youtube.com/watch?v=8mwCiDKgNd4',  // Topic2: HR Strategy and Planning
  'https://www.youtube.com/watch?v=ha2ZCiWKtTU',  // Topic3: HRP Programme Implementation
  'https://www.youtube.com/watch?v=oas5n1nFHQQ',  // Topic4: Job Analysis
  'https://www.youtube.com/watch?v=uUG-Z5sg2UM',  // Topic4: Job Design
];

const PDF_DIR = path.join(__dirname, 'course_pdfs');
const PDFS = fs.readdirSync(PDF_DIR).filter(f => f.endsWith('.pdf')).map(f => path.join(PDF_DIR, f));

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 100 });

  // Load saved session if exists
  let context;
  if (fs.existsSync(SESSION_FILE)) {
    console.log('Loading saved session...');
    context = await browser.newContext({ storageState: SESSION_FILE });
  } else {
    context = await browser.newContext();
  }

  const page = await context.newPage();
  await page.goto('https://notebooklm.google.com');
  await page.waitForLoadState('networkidle');

  // Check if login needed
  if (page.url().includes('accounts.google') || page.url().includes('signin')) {
    console.log('Please log in with Google in the browser...');
    await page.waitForURL(
      url => url.toString().includes('notebooklm.google.com') && !url.toString().includes('accounts.google'),
      { timeout: 300000 }
    );
    console.log('Logged in! Saving session...');
    await context.storageState({ path: SESSION_FILE });
  }

  console.log('On NotebookLM. Creating notebook for Malaika MGMT268...');
  await page.waitForTimeout(2000);

  // Click "New notebook" button
  const newNotebookBtn = await page.$(
    'button[aria-label*="new" i], button[aria-label*="create" i], [data-test-id*="new"], mat-card button, button:has-text("New notebook")'
  );
  if (newNotebookBtn) {
    await newNotebookBtn.click();
    console.log('Clicked New Notebook');
  } else {
    // Try clicking the + or create button
    await page.click('button:has-text("New"), button:has-text("Create"), [aria-label="New notebook"]').catch(() => {});
  }

  await page.waitForTimeout(3000);
  await page.screenshot({ path: path.join(__dirname, 'notebooklm_state.png') });

  // Save current state
  await context.storageState({ path: SESSION_FILE });
  console.log('Session saved to:', SESSION_FILE);
  console.log('\nBrowser staying open for manual interaction...');
  console.log('PDFs to add:', PDFS.length);
  console.log('YouTube videos to add:', YOUTUBE_LINKS.length);

  // Keep open for 5 minutes for manual use
  await page.waitForTimeout(300000);
  await browser.close();
})();
