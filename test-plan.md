# Test Plan — WPX→GT rebrand of training materials (PR #47)

Served locally at http://localhost:8099 (python http.server on branch devin/...-training-gt-rebrand).

## Test 1: Training index shows GT branding
- Open http://localhost:8099/wpx_training.html
- PASS if browser tab title reads "GT — Training Guides" and subtitle reads "Strategy resources for GT alliance members".
- FAIL if any "WPX" text appears in title/subtitle.

## Test 2: Guide card link works + guide page rebranded
- From training index, click the "Base Defense & Raiding" guide card.
- PASS if it navigates to WPX_Guide_Base_Defense.html (link intact), hero badge reads "GT Training Series", and footer logo reads "⚔️ GT ⚔️".
- FAIL if link 404s or any visible "WPX" branding remains.

## Test 3: Daily Duel blurb rebranded
- Open http://localhost:8099/WPX_Daily_Duel.html
- PASS if hero badge reads "GT Daily Briefing" and footer logo reads "⚔️ GT ⚔️".
- FAIL if "WPX" appears in badge/logo.

Note (not a failure): footer line "Warrior Phoenix Alliance — Training Series" intentionally left unchanged.
