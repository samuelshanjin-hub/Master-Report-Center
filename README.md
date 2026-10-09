# Master Report Center v2.0.2 — GitHub Release Candidate
**Publisher:** Golden Fir Solutions  
**Platform:** Windows 10/11  
**Application version:** 2.0.2

Master Report Center provides four reporting workflows: Split Reports, Consolidate Reports, Synchronize Reports and Lookup Data.

## Requirements
- Windows 10 or Windows 11
- Microsoft Excel desktop application is required for Excel-preserving / COM-based operations.
- Internet access is required for purchase, activation, recovery and licence-status synchronization.

## Trial and licensing
- 12-day trial
- Personal License — AUD 49 one-time
- Company License — AUD 149 one-time
- Purchase, activation, recovery and refund requests are handled through the Golden Fir licensing backend.

## Local application data
Runtime state is stored outside the application installation folder under:
`%LOCALAPPDATA%\GoldenFir\MasterReportCenter\`

## Building
Run `build_release.bat`. The release build uses PyInstaller ONEDIR and creates `dist\Master Report Center\`.

## Release warning
This package is the Version 2.0.2 release source baseline. The Cloudflare licensing backend and PayPal payment flow have been validated against the production/live environment, including purchase, activation, recovery, device migration, refund and post-refund license revocation.
