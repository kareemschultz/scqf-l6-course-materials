/**
 * Full pipeline for Malaika MGMT268 Assignment
 * 1. Auto-login to UWI
 * 2. Download ALL course resources (Topics 1-4) as PDFs
 * 3. Resolve all YouTube video links
 * 4. Save everything locally
 */
require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');

const BASE_URL = process.env.UWI_BASE_URL;
const USERNAME = process.env.UWI_USERNAME;
const PASSWORD = process.env.UWI_PASSWORD;

const PDF_DIR  = path.join(__dirname, 'course_pdfs');
const TXT_DIR  = path.join(__dirname, 'course_content');
const [PDF_DIR_OK, TXT_DIR_OK] = [PDF_DIR, TXT_DIR].map(d => { fs.mkdirSync(d, { recursive: true }); return d; });

// All resources to download (Topics 1-4)
const RESOURCES = [
  { id: '51764', name: 'Course_Overview' },
  { id: '51765', name: 'Course_Schedule' },
  { id: '51778', name: 'Getting_Started' },
  { id: '51782', name: 'Topic1_HRM_Overview_Notes' },
  { id: '51783', name: 'Topic1_HRM_Overview_Lecture' },
  { id: '51785', name: 'Topic1_Personnel_Mgmt_vs_HRM' },
  { id: '51788', name: 'Topic2_3_Strategy_HRP_Notes' },
  { id: '51789', name: 'Topic2_Strategic_Planning_Lecture' },
  { id: '51792', name: 'Topic3_HRP_Lecture' },
  { id: '51795', name: 'Topic4_Job_Analysis_Design_Lecture' },
  { id: '51796', name: 'Topic4_Job_Analysis_Design_Notes' },
  { id: '51799', name: 'Topic5_Recruitment_Selection_Notes1' },
];

// YouTube URL modules
const URL_MODULES = [
  { id: '51781', name: 'Topic1_Evolution_of_HRM' },
  { id: '51784', name: 'Topic1_Transformation_Personnel_to_HRM' },
  { id: '51787', name: 'Topic2_HR_Strategy_and_Planning' },
  { id: '51791', name: 'Topic3_HRP_Programme_Implementation' },
  { id: '51793', name: 'Topic4_Job_Analysis' },
  { id: '51794', name: 'Topic4_Job_Design' },
];

function downloadFile(url, dest, cookies) {
  return new Promise((resolve, reject) => {
    const cookieStr = cookies.map(c => `${c.name}=${c.value}`).join('; ');
    const follow = (u) => {
      const proto = u.startsWith('https') ? https : http;
      const file = fs.createWriteStream(dest);
      proto.get(u, { headers: { Cookie: cookieStr, 'User-Agent': 'Mozilla/5.0' } }, res => {
        if (res.statusCode === 301 || res.statusCode === 302) {
          file.close(); try { fs.unlinkSync(dest); } catch(_) {}
          return follow(res.headers.location);
        }
        res.pipe(file);
        file.on('finish', () => { file.close(); resolve(res.headers['content-type'] || ''); });
      }).on('error', e => { try { fs.unlinkSync(dest); } catch(_) {} reject(e); });
    };
    follow(url);
  });
}

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 200 });
  const context = await browser.newContext();
  const page = await context.newPage();

  // --- AUTO LOGIN ---
  console.log('\n=== Step 1: Auto-login ===');
  await page.goto(`${BASE_URL}/login/index.php`);
  await page.waitForLoadState('networkidle');
  await page.fill('#usernameUserInput', USERNAME);
  await page.fill('#password', PASSWORD);
  await page.click('button[type=submit]');
  await page.waitForURL(u => u.toString().includes('courses.global.uwi.edu/my'), { timeout: 30000 });
  console.log('Logged in as', USERNAME);

  const cookies = await context.cookies();

  // --- DOWNLOAD PDFs ---
  console.log('\n=== Step 2: Downloading course resources ===');
  for (const res of RESOURCES) {
    const url = `${BASE_URL}/mod/resource/view.php?id=${res.id}`;
    try {
      await page.goto(url);
      await page.waitForLoadState('networkidle');

      const dlLink = await page.$eval(
        'a[href*="pluginfile"], a[href*=".pdf"], a[href*=".pptx"], a[href*=".docx"], a[href*=".ppt"]',
        el => el.href
      ).catch(() => null);

      const finalUrl = dlLink || page.url();

      if (finalUrl.match(/\.(pdf|pptx|docx|ppt|doc)(\?|$)/i) || finalUrl.includes('pluginfile')) {
        const ext = (finalUrl.match(/\.(pdf|pptx|docx|ppt|doc)/i) || ['', 'pdf'])[1];
        const dest = path.join(PDF_DIR, `${res.name}.${ext}`);
        if (!fs.existsSync(dest)) {
          await downloadFile(finalUrl, dest, cookies);
          console.log(`  Downloaded: ${res.name}.${ext}`);
        } else {
          console.log(`  Already exists: ${res.name}.${ext}`);
        }
      } else {
        const text = await page.evaluate(() => document.body.innerText);
        fs.writeFileSync(path.join(TXT_DIR, `${res.name}.txt`), text);
        console.log(`  Saved as text: ${res.name}.txt`);
      }
    } catch(e) {
      console.log(`  Error for ${res.name}: ${e.message}`);
    }
  }

  // --- RESOLVE YOUTUBE LINKS ---
  console.log('\n=== Step 3: Resolving YouTube links ===');
  const youtubeLinks = {};
  for (const mod of URL_MODULES) {
    const url = `${BASE_URL}/mod/url/view.php?id=${mod.id}`;
    try {
      await page.goto(url);
      await page.waitForLoadState('networkidle');

      const ytLink = await page.$eval('a[href*="youtu"]', el => el.href).catch(() => null);
      const iframeSrc = await page.$eval('iframe', el => el.src).catch(() => null);
      const videoId = (iframeSrc || '').match(/embed\/([a-zA-Z0-9_-]{11})/)?.[1];
      const watchUrl = ytLink || (videoId ? `https://www.youtube.com/watch?v=${videoId}` : null);

      youtubeLinks[mod.name] = watchUrl;
      console.log(`  ${mod.name}: ${watchUrl}`);
    } catch(e) {
      console.log(`  Error for ${mod.name}: ${e.message}`);
    }
  }

  fs.writeFileSync(
    path.join(TXT_DIR, 'youtube_links.json'),
    JSON.stringify(youtubeLinks, null, 2)
  );

  // --- SUMMARY ---
  console.log('\n=== Done! ===');
  console.log('PDFs:', fs.readdirSync(PDF_DIR).length, 'files in', PDF_DIR);
  console.log('YouTube links:', Object.keys(youtubeLinks).length, 'saved to youtube_links.json');

  fs.writeFileSync(
    path.join(__dirname, 'pipeline_summary.json'),
    JSON.stringify({ pdfs: fs.readdirSync(PDF_DIR), youtubeLinks }, null, 2)
  );

  await browser.close();
})();
