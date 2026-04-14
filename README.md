## Multi-Region KML Zone Exporter
A robust Google Apps Script automation designed to process large-scale geospatial data from Google Sheets and export it into KML format for Google Earth. This system is built to handle multiple regions and customer types while bypassing standard Google script execution limits through intelligent batching.
## 🚀 Features

* Batch Processing: Automatically splits large datasets into manageable chunks to prevent script timeouts.
* Resume Capability: Tracks progress in a master sheet; if a run is interrupted, it resumes exactly where it left off.
* Multi-Region Support: Loads unique configurations (folders, spreadsheets, colors) for different regions from a central dashboard.
* Automated Error Handling: Includes a retry mechanism and sends email alerts to admins if a specific region fails multiple times.
* Zone Grouping: Automatically groups customers by geographic "Zones" for organized viewing in Google Earth.

------------------------------
## 🛠️ Setup Guide## 1. Prepare the Configuration Spreadsheet
Create a new Google Sheet to act as your Config Dashboard and add a sheet named Config with the following headers:

* Region, Type, SpreadsheetId, DashboardId, FolderId, ColorColumn

## 2. Install the Script

   1. In your Config Spreadsheet, go to Extensions > Apps Script.
   2. Delete any existing code and paste the provided script.
   3. Update the CONFIG_SPREADSHEET_ID at the top of the script with your spreadsheet's ID.

## 3. Configure Services & Permissions

   1. Enable Sheets API: In the Apps Script editor, click Services (+ icon), search for "Google Sheets API," and click Add.
   2. Authorize: Run the doExportAllZone function once manually to grant the necessary permissions (Drive, Sheets, and Email).

## 4. Create a Trigger (Optional but Recommended)
To fully automate the export, set up a time-based trigger:

   1. Click the Triggers (clock icon) in the Apps Script sidebar.
   2. Click + Add Trigger.
   3. Choose doExportAllZone as the function to run.
   4. Select Time-driven and set it to run every 5–10 minutes. The script will automatically check for pending tasks and stop when finished.

------------------------------
## 📂 Configuration Mapping

| Constant | Purpose |
|---|---|
| BATCH_SIZE | Number of rows processed in one go. |
| MAX_EXECUTION_TIME_MS | Safety limit (default 5 mins) before the script pauses to save progress. |
| ADMIN_EMAIL | Receives failure alerts after 4 unsuccessful attempts. |

