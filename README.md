# GFS Dashboard - One Screener Test

Test screener: ADVANCEGFSCASHSTOCKS
https://chartink.com/screener/advancegfscashstocks

This diagnostic build uses Playwright/Chromium to read the rendered Chartink result table first, with a JSON-response fallback. It keeps all displayed result columns.

Run GitHub Actions -> Refresh GFS Dashboard - One Screener Test -> Run workflow.

Expected log:
OK | rows=... | fields=... | method=...

Do not add the other five screeners until this test successfully captures live rows.
