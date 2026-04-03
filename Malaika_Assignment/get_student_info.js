require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  const page = await context.newPage();

  const BASE = process.env.UWI_BASE_URL;

  // Login
  await page.goto(`${BASE}/login/index.php`);
  await page.waitForLoadState('networkidle');
  await page.fill('#usernameUserInput', process.env.UWI_USERNAME);
  await page.fill('#password', process.env.UWI_PASSWORD);
  await page.click('button[type=submit]');
  await page.waitForURL(u => u.toString().includes('courses.global.uwi.edu/my'), { timeout: 30000 });

  // Go to profile page
  await page.goto(`${BASE}/user/profile.php`);
  await page.waitForLoadState('networkidle');
  const profileText = await page.evaluate(() => document.body.innerText);

  // Also try the edit profile for more details
  await page.goto(`${BASE}/user/edit.php`);
  await page.waitForLoadState('networkidle');

  const firstName = await page.$eval('#id_firstname', el => el.value).catch(() => '');
  const lastName  = await page.$eval('#id_lastname',  el => el.value).catch(() => '');
  const email     = await page.$eval('#id_email',     el => el.value).catch(() => '');
  const idNumber  = await page.$eval('#id_idnumber',  el => el.value).catch(() => '');

  const info = { firstName, lastName, email, idNumber, username: process.env.UWI_USERNAME };
  console.log('Student Info:', JSON.stringify(info, null, 2));
  fs.writeFileSync(path.join(__dirname, 'student_info.json'), JSON.stringify(info, null, 2));

  // Also check profile page for any additional info
  const nameOnPage = await page.$eval('.page-header-headings h1, h1.h2', el => el.innerText).catch(() => '');
  console.log('Name on page:', nameOnPage);

  await browser.close();
})();
