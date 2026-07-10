# Test Report — Login restricted to WPX2026 (PR #47)

**How tested:** Served the site locally on the PR branch, loaded `wpx_dashboard.html` (data fetched live from Dropbox), and exercised the login in Chrome.

## Results
- It should reject a valid roster player ID — **passed**
- It should unlock the master dashboard with WPX2026 — **passed**
- Login placeholder reads "Enter password" — **passed**

Adversarial note: I used `1001` = "Ender174", a **real** roster ID (pulled from the workbook's Roster sheet) that the OLD code accepted and would have locked the dashboard to that player. This distinguishes the change from a random invalid number (which both old and new code reject).

## Evidence

### Real roster ID (1001) is now rejected
Error "Invalid password. Please try again." shows; dashboard stays locked.

![ID rejected](/home/ubuntu/screenshots/ss_225c1f26.png)

### WPX2026 unlocks the full master dashboard
Player dropdown + full analytics render.

![Master unlocked](/home/ubuntu/screenshots/ss_0f19f67d.png)
