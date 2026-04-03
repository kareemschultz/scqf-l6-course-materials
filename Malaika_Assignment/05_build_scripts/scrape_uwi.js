const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE_URL = 'https://2025cpe.tle.courses.global.uwi.edu';
const COURSE_URL = `${BASE_URL}/mod/assign/view.php?id=51797`;
const OUTPUT_DIR = path.join(__dirname, 'course_content');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 100 });
  const context = await browser.newContext();
  const page = await context.newPage();

  console.log('Opening UWI platform — please log in...');
  await page.goto(COURSE_URL);

  // Wait for user to log in (up to 3 minutes)
  console.log('Waiting for you to log in... (up to 3 minutes)');
  await page.waitForURL(url => !url.toString().includes('login'), { timeout: 180000 });
  console.log('Logged in! Starting content extraction...');

  const results = [];

  // --- 1. Grab the assignment page ---
  await page.goto(COURSE_URL);
  await page.waitForLoadState('networkidle');
  const assignmentText = await page.evaluate(() => document.body.innerText);
  fs.writeFileSync(path.join(OUTPUT_DIR, 'assignment_brief.txt'), assignmentText);
  console.log('Saved: assignment_brief.txt');

  // --- 2. Navigate to the course home to find Topics 1-4 ---
  // Find course home link
  const courseHomeLink = await page.$('a[href*="/course/view.php"]');
  if (courseHomeLink) {
    const href = await courseHomeLink.getAttribute('href');
    await page.goto(href.startsWith('http') ? href : BASE_URL + href);
    await page.waitForLoadState('networkidle');
    const coursePageText = await page.evaluate(() => document.body.innerText);
    fs.writeFileSync(path.join(OUTPUT_DIR, 'course_home.txt'), coursePageText);
    console.log('Saved: course_home.txt');

    // --- 3. Find and visit all topic/section links ---
    const links = await page.$$eval('a[href]', anchors =>
      anchors
        .map(a => ({ text: a.innerText.trim(), href: a.href }))
        .filter(a => a.text && a.href && !a.href.includes('javascript'))
    );

    const topicLinks = links.filter(l =>
      /topic\s*[1-4]|week\s*[1-4]|lecture|slides|resource|youtube|youtu\.be|video/i.test(l.text + l.href)
    );

    console.log(`Found ${topicLinks.length} potential topic/resource links`);
    fs.writeFileSync(path.join(OUTPUT_DIR, 'all_links.json'), JSON.stringify(links, null, 2));
    fs.writeFileSync(path.join(OUTPUT_DIR, 'topic_links.json'), JSON.stringify(topicLinks, null, 2));

    // Visit each topic link and save content
    for (let i = 0; i < Math.min(topicLinks.length, 20); i++) {
      const link = topicLinks[i];
      try {
        console.log(`Visiting: ${link.text} -> ${link.href}`);
        await page.goto(link.href);
        await page.waitForLoadState('networkidle');
        const text = await page.evaluate(() => document.body.innerText);
        const filename = `topic_${i + 1}_${link.text.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 40)}.txt`;
        fs.writeFileSync(path.join(OUTPUT_DIR, filename), `URL: ${link.href}\n\n${text}`);
        console.log(`Saved: ${filename}`);
      } catch (e) {
        console.log(`Skipped ${link.href}: ${e.message}`);
      }
    }
  }

  // --- 4. Also dump all links from course for manual review ---
  console.log('\nAll done! Content saved to:', OUTPUT_DIR);
  console.log('Browser staying open for 30 seconds for manual review...');
  await page.goto(COURSE_URL);
  await new Promise(r => setTimeout(r, 30000));

  await browser.close();
  console.log('Browser closed.');
})();
