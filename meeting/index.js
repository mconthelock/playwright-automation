// routes/login.js
const express = require("express");
const router = express.Router();
const { chromium } = require("playwright");

router.get("/", async (req, res) => {
  try {
    const browser = await chromium.launch({ headless: false, slowMo: 100 });
    const context = await browser.newContext();
    const page = await context.newPage();
    await page.goto(
      "https://amecoutlook.mitsubishielevatorasia.co.th/owa/auth/logon.aspx?"
    );
    await page.fill(
      'input[name="username"]',
      "kanittha@MitsubishiElevatorAsia.co.th"
    );
    await page.fill('input[name="password"]', "ISdell11P@ssw0rd");
    await page.click("text=Sign in");
    await page.waitForNavigation();

    await page.click(
      'button[aria-label="Open the app launcher to access your Office 365 apps"]'
    );
    await page.waitForSelector("a#O365_AppTile_ShellCalendar", {
      timeout: 1000,
    });
    await page.click("a#O365_AppTile_ShellCalendar");
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;
