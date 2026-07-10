# Test Plan — Login restricted to WPX2026 (PR #47)

Page served locally at http://localhost:8099/wpx_dashboard.html on the PR branch.
Code refs: `wpx_dashboard.html` doLogin() L205-227 (player-ID branch removed), placeholder L113.
Data loads from Dropbox (verified reachable); wait for the "Data loaded" state before logging in.

## Test 1: A previously-valid roster player ID is now rejected (the key regression)
- On the login screen, type `1001` (a real roster ID = "Ender174", which the OLD code accepted and would have locked the dashboard to that player).
- Click "Enter".
- PASS if login is REJECTED: error text "Invalid password. Please try again." shows and the dashboard stays locked (login box still visible, no player data shown).
- FAIL if it unlocks the dashboard (old behavior).
- Adversarial note: a random invalid number would be rejected under BOTH old and new code — using a real ID (1001) is what distinguishes the change.

## Test 2: WPX2026 unlocks the master dashboard
- Type `WPX2026`, click "Enter".
- PASS if the login box disappears and the dashboard content loads with the player-select dropdown visible (master/full view).
- FAIL if rejected or dropdown not shown.

## Test 3: Placeholder text
- PASS if the login input placeholder reads "Enter password" (not "Enter your ID").
