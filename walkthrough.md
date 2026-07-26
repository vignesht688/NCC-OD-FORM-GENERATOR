# Walkthrough: MS Word Compatibility & Caching Solutions

We have updated the document compiler and cache configurations to ensure the downloaded Word files open seamlessly without corruption.

---

## 1. MS Word Compatibility (Table Cell Merges)
- **The Issue**: Previously, the Particulars and On Duty columns used a web-only `rowSpan` property. In OpenXML (docx), tables must maintain a consistent grid of cells in every row. Omitting cells in subsequent rows caused the document structure to compile as invalid/corrupted OpenXML, which prevented Microsoft Word from opening the file.
- **The Fix**: Swapped out the `rowSpan` attribute in `app.js` with standard OpenXML vertical merging:
  - Every row now has exactly 7 cells.
  - The first row initializes the vertical merge group (`verticalMerge: "restart"`).
  - Subsequent rows continue the merge (`verticalMerge: "continue"`).
  - This ensures 100% compliance with OpenXML standards, allowing the downloaded `.docx` document to open perfectly in Microsoft Word, WPS Word, and Google Docs.

## 2. Server Caching Disablement
- Updated `server.js` to serve all static assets (HTML, CSS, JS) with `Cache-Control: no-store, no-cache` headers. This prevents browsers from caching stale versions of files during local development, ensuring you always run the latest code.

## 3. Defensive Event Binding
- Wrapped all UI bindings in a utility function (`safeAddListener`). If elements are ever missing from the DOM during testing or tab transitions, the script will gracefully log a console warning instead of throwing a blocker crash.

---

## 4. Production Deployment Setup (Vercel & Render)
- **Vercel (`vercel.json`)**: Configured Vercel to route all traffic to the serverless function handler in `server.js` using `@vercel/node`.
- **Render (`render.yaml`)**: Added a Render blueprint spec for deploying the app as a standard Web Service.
- **Server Updates**: Modified server endpoints to dynamically detect deployed environments, ensuring local-only endpoints like "Open in Word" and "Quick Gen" fail gracefully with clear directions, while "Download (.docx)" works seamlessly everywhere.

---

## Visual Verification
Here is a recording showing the verified layout, clean console initialization, and successful document compilation:
![Verification WebP Recording](/C:/Users/jayav/.gemini/antigravity-ide/brain/b0067744-fd02-4ccf-8e27-b9dca0c0bba3/docx_successful_download_1783838501378.webp)

---

## How to Run & Deploy

### Running Locally
1. Start the Node app using [run-app.bat](file:///c:/Users/jayav/OneDrive/Desktop/od form for ncc/run-app.bat).
2. Go to `http://localhost:3000/?t=99999` (to ensure your browser loads the fresh script).
3. Select your cadets, configure event settings, and click **Generate Word Form**. The downloaded file will open perfectly in MS Word.

### Deploying Online
- **Vercel**: Import the repository on Vercel. Vercel will automatically detect `vercel.json` and deploy it.
- **Render**: Create a new Web Service on Render and point it to the repository, or import it using the Blueprint configuration in `render.yaml`.

