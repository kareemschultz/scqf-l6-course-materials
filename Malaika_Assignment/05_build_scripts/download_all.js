const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');

const BASE_URL = 'https://2025cpe.tle.courses.global.uwi.edu';
const OUTPUT_DIR = path.join(__dirname, 'course_content');
const PDF_DIR = path.join(__dirname, 'course_pdfs');

if (!fs.existsSync(PDF_DIR)) fs.mkdirSync(PDF_DIR, { recursive: true });

// All resource files (Topics 1-4 + key ones)
const RESOURCES = [
  { id: '51782', name: 'Topic1_HRM_Overview_Notes' },
  { id: '51783', name: 'Topic1_HRM_Overview_Lecture' },
  { id: '51785', name: 'Topic1_Personnel_Mgmt_vs_HRM' },
  { id: '51788', name: 'Topic2_3_Strategy_HRP_Notes' },
  { id: '51789', name: 'Topic2_Strategic_Planning_Lecture' },
  { id: '51792', name: 'Topic3_HRP_Lecture' },
  { id: '51795', name: 'Topic4_Job_Analysis_Design_Lecture' },
  { id: '51796', name: 'Topic4_Job_Analysis_Design_Notes' },
];

// URL modules (YouTube videos)
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
    const options = {
      headers: { 'Cookie': cookieStr, 'User-Agent': 'Mozilla/5.0' }
    };
    const proto = url.startsWith('https') ? https : http;
    const file = fs.createWriteStream(dest);
    proto.get(url, options, (res) => {
      if (res.statusCode === 302 || res.statusCode === 301) {
        file.close();
        fs.unlinkSync(dest);
        return downloadFile(res.headers.location, dest, cookies).then(resolve).catch(reject);
      }
      res.pipe(file);
      file.on('finish', () => { file.close(); resolve(res.headers['content-type'] || ''); });
    }).on('error', (err) => { fs.unlinkSync(dest); reject(err); });
  });
}

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 50 });
  const context = await browser.newContext();
  const page = await context.newPage();

  // Check if already logged in by going to course
  console.log('Checking login status...');
  await page.goto(`${BASE_URL}/mod/assign/view.php?id=51797`);

  if (page.url().includes('login')) {
    console.log('Please log in...');
    await page.waitForURL(url => !url.toString().includes('login'), { timeout: 180000 });
  }
  console.log('Logged in!');

  // Get cookies for direct downloads
  const cookies = await context.cookies();

  // --- Resolve YouTube / URL module links ---
  console.log('\n=== Resolving YouTube video URLs ===');
  const videoLinks = {};
  for (const mod of URL_MODULES) {
    const url = `${BASE_URL}/mod/url/view.php?id=${mod.id}`;
    try {
      await page.goto(url);
      await page.waitForLoadState('networkidle');
      // Get the actual redirect URL or embedded link
      const finalUrl = page.url();
      const pageText = await page.evaluate(() => document.body.innerText);
      const iframeUrl = await page.$eval('iframe', el => el.src).catch(() => null);
      const linkUrl = await page.$eval('a[href*="youtube"], a[href*="youtu.be"]', el => el.href).catch(() => null);

      const youtubeUrl = iframeUrl || linkUrl || (finalUrl.includes('youtube') || finalUrl.includes('youtu.be') ? finalUrl : null);

      videoLinks[mod.name] = {
        moodleUrl: url,
        youtubeUrl: youtubeUrl,
        finalUrl: finalUrl,
        snippet: pageText.substring(0, 300)
      };
      console.log(`${mod.name}: ${youtubeUrl || finalUrl}`);
    } catch(e) {
      console.log(`Error for ${mod.name}: ${e.message}`);
    }
  }

  fs.writeFileSync(
    path.join(OUTPUT_DIR, 'youtube_video_links.json'),
    JSON.stringify(videoLinks, null, 2)
  );
  console.log('Saved: youtube_video_links.json');

  // --- Download PDF resources ---
  console.log('\n=== Downloading PDF/file resources ===');
  for (const res of RESOURCES) {
    const url = `${BASE_URL}/mod/resource/view.php?id=${res.id}`;
    try {
      await page.goto(url);
      await page.waitForLoadState('networkidle');

      // Try to find direct download link
      const downloadLink = await page.$eval(
        'a[href*="pluginfile"], a[href*=".pdf"], a[href*=".pptx"], a[href*=".docx"], a[href*=".ppt"]',
        el => el.href
      ).catch(() => null);

      const finalUrl = downloadLink || page.url();

      if (finalUrl && (finalUrl.includes('pluginfile') || finalUrl.match(/\.(pdf|pptx|docx|ppt|doc)(\?|$)/i))) {
        const ext = finalUrl.match(/\.(pdf|pptx|docx|ppt|doc)/i)?.[1] || 'pdf';
        const dest = path.join(PDF_DIR, `${res.name}.${ext}`);
        console.log(`Downloading ${res.name}.${ext}...`);
        const contentType = await downloadFile(finalUrl, dest, cookies);
        console.log(`  Saved: ${res.name}.${ext} (${contentType})`);
      } else {
        // Save page text as fallback
        const text = await page.evaluate(() => document.body.innerText);
        fs.writeFileSync(path.join(OUTPUT_DIR, `${res.name}.txt`), `URL: ${url}\nFinal: ${finalUrl}\n\n${text}`);
        console.log(`  Saved as text: ${res.name}.txt`);
      }
    } catch(e) {
      console.log(`Error for ${res.name}: ${e.message}`);
    }
  }

  console.log('\n=== All done! ===');
  console.log('PDFs saved to:', PDF_DIR);
  console.log('YouTube links saved to:', path.join(OUTPUT_DIR, 'youtube_video_links.json'));

  await browser.close();
})();
