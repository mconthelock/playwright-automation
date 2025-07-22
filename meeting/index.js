// routes/login.js
const express = require("express");
const router = express.Router();
const { chromium } = require("playwright");
const launchOptions = {
  headless: true,
  args: [
    '--no-sandbox',
    '--disable-setuid-sandbox',
    '--disable-blink-features=AutomationControlled',
    '--disable-infobars',
    '--start-maximized',
  ],
};

router.get("/", async (req, res) => {
  try {
    const browser = await chromium.launch(launchOptions);
    const context = await browser.newContext({
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/115.0.0.0 Safari/537.36',
      viewport: { width: 1280, height: 800 },
      locale: 'en-US',
      permissions: ['notifications'],
    });
    const page = await context.newPage();
    await evadeDetection(page);
  
    await page.goto('https://amecoutlook.mitsubishielevatorasia.co.th/owa/auth/logon.aspx?');
  
    await page.fill('input[name="username"]', 'kanittha@MitsubishiElevatorAsia.co.th');
    await page.fill('input[name="password"]', 'ISdell11P@ssw0rd');
    await page.click('text=Sign in');
    await page.waitForNavigation();
  
    await page.click('button[aria-label="Open the app launcher to access your Office 365 apps"]');
    await page.waitForSelector('a#O365_AppTile_ShellCalendar');
    await page.click('a#O365_AppTile_ShellCalendar');
    await page.waitForURL(/\/calendar/);
    await page.waitForTimeout(3000);
  
    await page.getByRole('button', { name: 'New' }).click();
    await page.waitForSelector('input[placeholder="Add a title for the event"]');
    await page.fill('input[placeholder="Add a title for the event"]', 'ประชุมโครงการ A');
  
    await page.click('button[aria-label^="start date"]');
    const startPopupId = await page.getAttribute('button[aria-label^="start date"]', 'aria-owns');
    await selectDateWithMonthYear(page, startPopupId, 2025, 7, 26);
    await page.fill('input[aria-label="start time"]', '09:30 AM');
  
    await page.click('button[aria-label^="end date"]');
    const endPopupId = await page.getAttribute('button[aria-label^="end date"]', 'aria-owns');
    await selectDateWithMonthYear(page, endPopupId, 2025, 7, 26);
    await page.fill('input[aria-label="end time"]', '10:30 AM');
  
    await page.click('input[aria-labelledby="MeetingCompose.LocationInputLabel"]');
    await selectRoom(page, "EP Fueangfa Room");
  
    const attendees = [
      'kanittha@MitsubishiElevatorAsia.co.th',
      'chalorms@MitsubishiElevatorAsia.co.th'
    ];
  
    for (const email of attendees) {
      const input = page.getByRole('textbox', { name: 'Add people' });
      await input.click();
      await input.fill(email);
      await page.waitForTimeout(1000);
      await page.keyboard.press('Enter');
      await page.waitForTimeout(500);
    }
  
    await page.getByRole('button', { name: 'Send' }).click();
    console.log('✅ ส่งคำเชิญเรียบร้อยแล้ว');
  
    await logoutFromOutlook(page);
    await browser.close();
    console.log('✅ ปิด browser แล้ว');

  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

async function evadeDetection(page) {
  await page.addInitScript(() => {
    Object.defineProperty(navigator, 'webdriver', {
      get: () => false,
    });

    window.chrome = {
      runtime: {},
    };

    Object.defineProperty(navigator, 'plugins', {
      get: () => [1, 2, 3],
    });

    Object.defineProperty(navigator, 'languages', {
      get: () => ['en-US', 'en'],
    });
  });
}

async function selectDateWithMonthYear(page, calendarPopupId, targetYear, targetMonth, targetDay) {
  const monthYearSelector = `#${calendarPopupId} button._dx_g`;
  const nextButtonSelector = `#${calendarPopupId} button._dx_k`;
  const prevButtonSelector = `#${calendarPopupId} button._dx_l`;

  const parseMonthYear = (text) => {
    const [monthName, yearStr] = text.split(' ');
    const monthNames = [
      'January','February','March','April','May','June',
      'July','August','September','October','November','December'
    ];
    return {
      monthNum: monthNames.indexOf(monthName) + 1,
      yearNum: parseInt(yearStr)
    };
  };

  await page.waitForSelector(monthYearSelector);
  let currentMonthYear = await page.locator(monthYearSelector).innerText();
  let { monthNum, yearNum } = parseMonthYear(currentMonthYear);

  while (yearNum !== targetYear || monthNum !== targetMonth) {
    if (yearNum > targetYear || (yearNum === targetYear && monthNum > targetMonth)) {
      await page.click(prevButtonSelector);
    } else {
      await page.click(nextButtonSelector);
    }
    await page.waitForTimeout(500);
    currentMonthYear = await page.locator(monthYearSelector).innerText();
    ({ monthNum, yearNum } = parseMonthYear(currentMonthYear));
  }

  const dayCell = page.locator(`#${calendarPopupId} abbr`, { hasText: String(targetDay) });
  await dayCell.first().click({ force: true });
}

async function selectRoom(page, roomName) {
  await page.getByRole('button', { name: 'Add room' }).click();
  
console.log("✅ กดปุ่ม Add room แล้ว");

// ✅ รอให้ popup ขึ้น
await page.waitForSelector('div[role="dialog"], div._exadr_c.scrollContainer', { timeout: 15000 })
  .catch(() => {
    throw new Error("❌ ยังไม่เห็น popup เลือกห้อง");
  });

// ✅ รอให้เลิกแสดง “Finding rooms…”
await waitUntilRoomsAreLoaded(page);

// ✅ ต่อด้วยการโหลด container ห้องและเลือกห้องได้เลย
const container = page.locator('div._exadr_c.scrollContainer');
await container.waitFor({ state: 'attached', timeout: 10000 });

const allRooms = await page.locator('span._exadr_r').allTextContents();
console.log("📝 ห้องที่เจอทั้งหมด:", allRooms.join('|'));



  await container.evaluate(el => el.scrollTop = el.scrollHeight);
  await page.waitForTimeout(2000);

  const targetRoom = page.locator('span._exadr_r', { hasText: roomName }).first();
  if ((await targetRoom.count()) === 0) throw new Error(`ไม่พบห้อง ${roomName}`);

  await targetRoom.scrollIntoViewIfNeeded();
  await targetRoom.click();


  // ดักชื่อห้องแบบยืดหยุ่น
  const roomLocator = page.locator('span', { hasText: roomName });

  try {
    await roomLocator.first().waitFor({ state: 'visible', timeout: 10000 });
    await roomLocator.first().scrollIntoViewIfNeeded();
    await roomLocator.first().click();
    console.log(`✅ คลิกห้อง "${roomName}" สำเร็จ`);
  } catch (e) {
    await page.screenshot({ path: 'room-not-found.png' });
    const html = await page.content();
    require('fs').writeFileSync('room-not-found.html', html);
    throw new Error(`❌ ไม่พบห้อง "${roomName}" หรือยังไม่โหลดทั้งหมด`);
  }
}





async function logoutFromOutlook(page) {
  try {
    await page.click('button[aria-label*="menu with submenu"]');
    await page.waitForTimeout(1000);
    const signOutButton = page.getByRole('menuitem', { name: 'Sign out' });
    await signOutButton.waitFor({ state: 'visible', timeout: 5000 });
    await signOutButton.click();
    await signOutButton.waitFor({ state: 'detached', timeout: 5000 });
  } catch (err) {
    console.error('⚠️ Logout Failed:', err.message);
  }
}

async function waitUntilRoomsAreLoaded(page) {
  console.log("⏳ กำลังรอให้ระบบโหลดห้องให้เสร็จ...");

  const findingRoomsLocator = page.locator('span._fce_7._fce_8', { hasText: 'Finding rooms...' });

  // รอจนข้อความ "Finding rooms..." หายไป หรือ element หายไปจาก DOM
  await findingRoomsLocator.waitFor({ state: 'detached', timeout: 20000 }).catch(() => {
    console.warn("⚠️ รอนานเกินไป แต่ยังเห็น 'Finding rooms...' อยู่");
  });

  console.log("✅ ห้องโหลดเสร็จแล้ว!");
}


module.exports = router;
