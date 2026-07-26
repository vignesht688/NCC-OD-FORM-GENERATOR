# Permanent Cadet Database Storage Implementation Plan

Currently, the NCC Attendance & OD Form Automator loads default cadets from [cadets.json](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/cadets.json) and saves any new cadets in the browser's `localStorage`. This makes added cadets **temporary** (they are lost if browser cache is cleared, if opened in a different browser, or on a different machine).

To make the cadet database **permanent**, we will implement a dual-mode solution:
1. **Server Mode (Automatic Sync)**: A local Node/Express backend that reads and writes directly to [cadets.json](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/cadets.json) on disk when cadets are added.
2. **Offline Mode (Manual Backup)**: Import and Export buttons in the UI that let users download the updated `cadets.json` and replace the local file manually.

---

## User Review Required

> [!IMPORTANT]
> Running the backend server requires Node.js. If you run the backend using `npm start` (which we will configure), your cadet additions will be **instantly and permanently** saved to `cadets.json`.
>
> If you prefer not to run a local server, you will still be able to use the new **Export Database** button to download a permanent `cadets.json` file and place it in the folder manually.

---

## Proposed Changes

### 1. Backend Server Component
We will create a lightweight Express server in the project root to serve the files and read/write the JSON database.

#### [NEW] [server.js](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/server.js)
- Instantiate an Express application.
- Serve static files from the project directory.
- Define `GET /api/cadets` to read and serve [cadets.json](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/cadets.json).
- Define `POST /api/cadets` to write the updated cadet database list directly back to [cadets.json](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/cadets.json).
- Auto-open the application in the default browser on startup.

#### [MODIFY] [package.json](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/package.json)
- Add `"express"` dependency.
- Add `"start": "node server.js"` script to easily run the backend.

---

### 2. Frontend Updates

#### [MODIFY] [index.html](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/index.html)
- Add a new "Database Management" section or actions under "Add New Cadet" / "Select Cadets" tab to:
  - Download current database (`cadets.json` backup).
  - Upload/restore database file.
- Add a visual indicator showing whether the application is connected to the backend server (Sync Active) or running standalone (Offline Mode).

#### [MODIFY] [styles.css](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/styles.css)
- Style database action buttons and the connection/sync status badge.

#### [MODIFY] [app.js](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/app.js)
- Check connection to the local Express server on load.
- If connected:
  - Fetch database from `GET /api/cadets`.
  - When saving/adding a cadet, perform a POST request to `POST /api/cadets` to save directly to disk.
- If not connected (standalone `file://` or static file hosting):
  - Fallback to fetching default `cadets.json` and merging with `localStorage` (current behavior).
  - Show a banner indicating "Offline mode - changes saved to browser only".
- Implement functions for:
  - Exporting the current combined database as a downloaded `cadets.json` file.
  - Importing a JSON file to merge or replace the database.

---

## Verification Plan

### Automated/Manual verification
1. Run `npm install` to install express.
2. Launch the backend using `npm start` (or `node server.js`).
3. Verify that the server starts on `http://localhost:3000` and opens the browser.
4. Add a cadet via the "Add New Cadet" UI.
5. Verify that the cadet is written directly into the local [cadets.json](file:///c:/Users/jayav/OneDrive/Desktop/od%20form%20for%20ncc/cadets.json) file on the disk.
6. Verify that the cadet is preserved across browser refreshes and different browsers.
7. Test the offline fallback:
   - Close the Express server.
   - Double click `index.html` to open it via `file://`.
   - Verify that the app shows "Offline Mode" and supports manual database download/upload.
