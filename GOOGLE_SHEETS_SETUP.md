# Google Sheets Shared Setup

Use this when you want everyone opening the website to see and add to the same shared data.

1. Create a Google Sheet.
2. In the sheet, open `Extensions` > `Apps Script`.
3. Paste the contents of `Code.gs` or `google-apps-script.gs`.
4. Click `Deploy` > `New deployment`.
5. Choose type `Web app`.
6. Set `Execute as` to `Me`.
7. Set `Who has access` to `Anyone with the link`.
8. Deploy and copy the Web app URL.
9. Paste that URL into `config.js` as `SHEETS_WEB_APP_URL`.

After that, publish the website folder. Everyone using the same website link will read and add rows through the same spreadsheet. Received dates added from the website will show in Google Sheets as `dd / mm / yyyy`.

The website will show an `Open spreadsheet` link after the shared spreadsheet connects. If you want the link to appear before the first connection finishes, paste the Google Sheet URL into `config.js` as `SPREADSHEET_URL`.

If you already have rows in the sheet with a blank `ID` column, keep them. The script will add missing IDs automatically the next time the website loads or a report is generated, so those older rows can appear on the website and in the PDF summary.

The website checks the shared spreadsheet automatically while it is open. By default it refreshes every 10 seconds, controlled by `AUTO_REFRESH_SECONDS` in `config.js`.

The `PDF Summary` button uses the `Weekly Summary` tab. It updates that tab with the selected date range, adds the category chart, creates a PDF in Google Drive, and opens the PDF link. After pasting the latest script, Google may ask you to authorize Drive/PDF export permissions the first time you run it.

You can also create the PDF from inside the spreadsheet. Reload the Google Sheet after deploying the script, then use `ClaimsCo` > `Create Weekly Summary PDF`. If you want a sheet button, add a Drawing in Google Sheets and assign it this script name: `createWeeklySummaryPdfFromPrompt`.

If the page says it cannot reach the shared spreadsheet, replace the Apps Script code with the latest `google-apps-script.gs`, then deploy a new version of the Web app. The website uses a script-tag connection so it can work from a static website such as GitHub Pages.

Quick test: open your Web app URL in an incognito/private browser window. It should show JSON starting with `{"ok":true`. If it asks you to sign in, says authorization is required, or shows an HTML error page, the deployment access is still restricted.
