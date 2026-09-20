# ORIGIN FIBER Executive Dashboard

Google Apps Script web application for the Origin Fiber executive dashboard and the daily LINE performance report.

## Website

The dashboard is published as a static GitHub Pages website while continuing to read from the original Google Sheets database through the Apps Script API:

https://adisornchamnongchit.github.io/ORIGIN-FIBER-EXECUTIVE-DASHBOARD-Data-Base/

The browser keeps the latest successful response locally, so repeat visits can display the dashboard immediately while fresh data loads in the background.

## Project files

- `Code.gs` — dashboard data API, LINE report generator, webhook handling, and daily trigger
- `index.html` — responsive dashboard interface for Apps Script and GitHub Pages
- `appsscript.json` — Apps Script manifest and required OAuth scopes

## Daily LINE report

- Sends to the configured `ORF Report` LINE group
- Includes an overview and team performance for five operational zones
- Sends the previous day's data
- The default time-driven trigger runs daily at approximately 20:00 Asia/Bangkok

## Required Script Properties

Configure these values in **Apps Script → Project Settings → Script Properties**. Never commit their values to this repository.

- `LINE_CHANNEL_ACCESS_TOKEN`
- `LINE_REPORT_GROUP_ID`
- `LINE_REPORT_GROUP_VERIFIED_NAME`
- `LINE_WEBHOOK_KEY`

## Deployment

Deploy the Apps Script project as a web app and configure the LINE Messaging API webhook with the deployed web app URL and webhook key.
