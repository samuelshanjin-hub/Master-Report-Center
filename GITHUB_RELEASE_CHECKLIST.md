# Master Report Center v2.0.2 — GitHub Release Checklist
## Build
- [ ] Run `build_release.bat` on Windows.
- [ ] Confirm `dist\Master Report Center\Master Report Center.exe` starts without Python installed.
- [ ] Confirm Excel COM functions work with Microsoft Excel desktop installed.
- [ ] Confirm no `.env`, local database, backend secret, test log, test source, or developer patch file exists in `dist`.
## Backend / payments
- [ ] Confirm desktop default API points to the intended Cloudflare Worker.
- [ ] Confirm Worker `/health` reports the intended production environment.
- [ ] Confirm PayPal credentials are Live before public release.
- [ ] Confirm PayPal webhook is Live and points to the production Worker.
- [ ] Confirm Personal AUD 49 / Company AUD 149.
- [ ] Confirm terms version `MRC-PURCHASE-1.0`.
## Functional release test
- [ ] Fresh install / first start.
- [ ] Trial shows 12 days.
- [ ] All four modules run during trial.
- [ ] Purchase flow opens PayPal.
- [ ] Confirmed payment activates the license.
- [ ] Active state persists after restart.
- [ ] Same-device recovery works without consuming a migration.
- [ ] New-device recovery / migration works.
- [ ] Refund request reaches backend.
- [ ] Confirmed refund changes state to REFUNDED and locks modules.
- [ ] Existing user-created report files remain untouched.
## Paths
- [ ] Trial/license/install ID under `%LOCALAPPDATA%\GoldenFir\MasterReportCenter`.
- [ ] Logs under `%LOCALAPPDATA%\GoldenFir\MasterReportCenter\logs`.
- [ ] Imported configuration under `%LOCALAPPDATA%\GoldenFir\MasterReportCenter\config`.
- [ ] User reports write only to user-selected/user-owned locations.
## GitHub assets
- [ ] Release title `Master Report Center v2.0.2`.
- [ ] Upload installer or zipped `dist\Master Report Center\` package.
- [ ] Generate SHA-256 checksum.
- [ ] Add release notes, support, website and privacy policy links.
- [ ] State Microsoft Excel desktop requirement and 12-day trial.
