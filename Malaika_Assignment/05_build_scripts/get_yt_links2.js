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
  const browser = await chromium.launch({ headless: false, slowMo: 300 });
  const context = await browser.newContext();
  const page = await context.newPage();

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
    await page.waitForTimeout(1000);

    // Dump ALL links on the page
    const allLinks = await page.$$eval('a[href]', els =>
      els.map(el => ({ text: el.innerText.trim(), href: el.href }))
    );
    const ytLinks = allLinks.filter(l =>
      l.href.includes('youtube.com') || l.href.includes('youtu.be')
    );
    console.log(`  All links count: ${allLinks.length}`);
    console.log(`  YouTube links: ${JSON.stringify(ytLinks)}`);

    // Get iframe src
    const iframes = await page.$$eval('iframe', els => els.map(el => ({ src: el.src, id: el.id })));
    console.log(`  Iframes: ${JSON.stringify(iframes)}`);

    // Screenshot for inspection
    await page.screenshot({ path: path.join(__dirname, `screenshots_${mod.id}.png`) });

    // Extract video ID from embed URL and build watch URL
    let watchUrl = null;
    for (const iframe of iframes) {
      const match = iframe.src.match(/(?:embed\/|v=)([a-zA-Z0-9_-]{11})/);
      if (match) {
        watchUrl = `https://www.youtube.com/watch?v=${match[1]}`;
        break;
      }
    }

    if (!watchUrl && ytLinks.length > 0) {
      watchUrl = ytLinks[0].href;
    }

    if (watchUrl) {
      console.log(`  >>> YouTube Watch URL: ${watchUrl}`);
      results[mod.name] = watchUrl;
    } else {
      // Try clicking "Watch on YouTube" text or any button
      const ytBtn = await page.$('a[title*="YouTube"], a[aria-label*="YouTube"], [class*="youtube"], button[class*="play"]');
      if (ytBtn) {
        const [newTab] = await Promise.all([
          context.waitForEvent('page', { timeout: 5000 }).catch(() => null),
          ytBtn.click()
        ]);
        if (newTab) {
          await newTab.waitForLoadState('domcontentloaded').catch(() => {});
          watchUrl = newTab.url();
          await newTab.close();
        }
      }
      results[mod.name] = watchUrl || `No YouTube URL found - check screenshot: screenshots_${mod.id}.png`;
      console.log(`  >>> Result: ${results[mod.name]}`);
    }
  }

  console.log('\n\n=== FINAL YouTube Watch Links ===');
  for (const [name, link] of Object.entries(results)) {
    console.log(`${name}: ${link}`);
  }

  fs.writeFileSync(
    path.join(__dirname, 'course_content', 'youtube_watch_links.json'),
    JSON.stringify(results, null, 2)
  );

  await browser.close();
})();
