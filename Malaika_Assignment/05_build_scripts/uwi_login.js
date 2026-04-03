/**
 * Auto-login helper for UWI platform.
 * Usage: const { getLoggedInPage } = require('./uwi_login');
 */
require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const { chromium } = require('playwright');

const BASE_URL = process.env.UWI_BASE_URL;
const USERNAME = process.env.UWI_USERNAME;
const PASSWORD = process.env.UWI_PASSWORD;

async function getLoggedInPage() {
  const browser = await chromium.launch({ headless: false, slowMo: 100 });
  const context = await browser.newContext();
  const page = await context.newPage();

  await page.goto(`${BASE_URL}/login/index.php`);
  await page.waitForLoadState('networkidle');

  // Fill login form
  await page.fill('#usernameUserInput', USERNAME);
  await page.fill('#password', PASSWORD);
  await page.click('button[type=submit]');
  await page.waitForLoadState('networkidle');

  // Wait for redirect to course platform (SAML SSO flow)
  await page.waitForURL(url => url.toString().includes('courses.global.uwi.edu/my') || url.toString().includes('courses.global.uwi.edu/course'), { timeout: 30000 });
  const url = page.url();
  if (!url.includes('courses.global.uwi.edu')) {
    throw new Error('Login failed — check credentials in .env');
  }

  console.log('Logged in successfully as', USERNAME);
  return { browser, context, page };
}

module.exports = { getLoggedInPage, BASE_URL };
