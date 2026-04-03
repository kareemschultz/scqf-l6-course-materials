const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE_URL = 'https://2025cpe.tle.courses.global.uwi.edu';

const URL_MODULES = [
  { id: '51781', name: 'Topic1_Evolution_of_HRM' },
  { id: '51784', name: 'Topic1_Transformation_Personnel_to_HRM' },
  { id: '51787', name: 'Topic2_HR_Strategy_and_Planning' },
  { id: '51791', name: 'Topic3_HRP_Programme_Implementation' },
  { id: '51793', name: 'Topic4_Job_Analysis' },
  { id: '51794', name: 'Topic4_Job_Design' },
];

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 200 });
  const context = await browser.newContext();
  const page = await context.newPage();

  // Check login
  await page.goto(`${BASE_URL}/mod/assign/view.php?id=51797`);
  if (page.url().includes('login')) {
    console.log('Please log in...');
    await page.waitForURL(url => !url.toString().includes('login'), { timeout: 180000 });
  }
  console.log('Logged in!\n');

  const results = {};

  for (const mod of URL_MODULES) {
    const url = `${BASE_URL}/mod/url/view.php?id=${mod.id}`;
    console.log(`\nVisiting: ${mod.name}`);
    await page.goto(url);
    await page.waitForLoadState('networkidle');

    // Look for "Go to" link or YouTube button that opens the real URL
    // Moodle URL modules typically have a "Go to [url]" link
    const goLink = await page.$eval(
      'a[href*="youtube"], a[href*="youtu.be"], .urlworkaround a, [class*="url"] a, a[target="_blank"]',
      el => el.href
    ).catch(() => null);

    if (goLink) {
      console.log(`  Found link: ${goLink}`);
      results[mod.name] = goLink;
      continue;
    }

    // Try clicking the "Visit resource" or play button
    const visitBtn = await page.$('a[href*="youtube"], button, .mod_url a, [class*="resourcecontent"] a, a[rel="noopener"]');
    if (visitBtn) {
      // Set up listener for new tab/navigation
      const [newPage] = await Promise.all([
        context.waitForEvent('page', { timeout: 5000 }).catch(() => null),
        visitBtn.click()
      ]);

      if (newPage) {
        await newPage.waitForLoadState('domcontentloaded').catch(() => {});
        const ytUrl = newPage.url();
        console.log(`  New tab opened: ${ytUrl}`);
        results[mod.name] = ytUrl;
        await newPage.close();
      } else {
        await page.waitForLoadState('networkidle').catch(() => {});
        const ytUrl = page.url();
        console.log(`  Navigated to: ${ytUrl}`);
        results[mod.name] = ytUrl;
      }
    } else {
      // Extract from iframe src and convert embed -> watch
      const iframeSrc = await page.$eval('iframe', el => el.src).catch(() => null);
      if (iframeSrc) {
        const videoId = iframeSrc.match(/embed\/([a-zA-Z0-9_-]+)/)?.[1];
        const watchUrl = videoId ? `https://www.youtube.com/watch?v=${videoId}` : iframeSrc;
        console.log(`  From iframe: ${watchUrl}`);
        results[mod.name] = watchUrl;
      } else {
        console.log(`  Could not find YouTube link`);
        results[mod.name] = url;
      }
    }
  }

  console.log('\n=== YouTube Video Links ===');
  for (const [name, link] of Object.entries(results)) {
    console.log(`${name}:\n  ${link}`);
  }

  fs.writeFileSync(
    path.join(__dirname, 'course_content', 'youtube_watch_links.json'),
    JSON.stringify(results, null, 2)
  );
  console.log('\nSaved: youtube_watch_links.json');

  await browser.close();
})();
