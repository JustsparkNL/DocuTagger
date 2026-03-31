# DocuTagger by Uninova

[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)
[![Office Add-in](https://img.shields.io/badge/Microsoft%20Office-Add--in-orange.svg)](https://learn.microsoft.com/en-us/office/dev/add-ins/)

A Microsoft Word Task Pane Add-in by **Uninova B.V.** that automatically replaces `{{ID}}` placeholders throughout a Word document with the document's SharePoint Document ID.

---

## Features

- Retrieves the document ID from **SharePoint Online** via REST API
- Falls back to **custom XML metadata** embedded in the document when SharePoint is unavailable
- Replaces `{{ID}}` placeholders in:
  - Document body (paragraphs)
  - Shapes and text boxes
  - Primary headers (all sections)
  - Primary footers (all sections)
- Displays real-time status and a count of replacements made
- Hosted on GitHub Pages — no server setup needed for end users

---

## How it works

1. The user opens a Word document and clicks **DocuTagger** in the Home ribbon.
2. The add-in attempts to resolve the document ID in order:
   1. **SharePoint Online** — if the document URL contains `sharepoint.com`, it calls `_api/web/GetFileByServerRelativeUrl()` and reads the `OData__dlc_DocId` field.
   2. **Custom XML fallback** — searches the document's embedded custom XML parts for a `_dlc_DocId` value.
3. Once the ID is resolved, the add-in uses the Word JavaScript API to find every `{{ID}}` occurrence and replaces it with the actual document ID.
4. Status is shown directly in the task pane.

---

## Prerequisites

- Microsoft Word (desktop or online)
- A SharePoint Online site (for automatic ID retrieval) or a Word document with a `_dlc_DocId` custom XML part
- To **sideload** the add-in: admin access or the ability to upload a custom manifest in your Microsoft 365 tenant

---

## Installation (end users)

Sideload the add-in by uploading `manifest.xml` to your Word environment:

**Word Desktop (Windows/Mac):**
1. Go to **Insert → Add-ins → My Add-ins → Upload My Add-in**
2. Browse to `manifest.xml` and click **Upload**
3. The **DocuTagger** button appears in the **Home** tab ribbon

**SharePoint App Catalog (organization-wide):**
1. Upload `manifest.xml` to your tenant's SharePoint App Catalog
2. The add-in will be available to all users in the organization

Once loaded, open any Word document containing `{{ID}}` placeholders and click the **DocuTagger** button.

---

## Development setup

### 1. Install dependencies

```bash
npm install --global http-server
npm install --global office-addin-dev-certs
npx office-addin-dev-certs install
```

### 2. Start the local dev server

```bash
npm run dev-server
```

Or serve the `src/` directory directly:

```bash
http-server src/ --cors -p 3000
```

### 3. Sideload for local testing

Update the `SourceLocation` URL in `manifest.xml` to point to your local server (e.g., `https://localhost:3000/taskpane.html`), then sideload the manifest as described above.

### 4. Build

```bash
npm run build:dev   # development build
npm run build       # production build
```

### 5. Watch mode

```bash
npm run watch
```

### 6. Lint

```bash
npm run lint        # check for issues
npm run lint:fix    # auto-fix issues
```

### VS Code debugging

Use the launch configurations in `.vscode/launch.json` to debug against Word Desktop (Edge Chromium, port 9229).

---

## Deployment

The add-in is hosted on **GitHub Pages** at:

```
https://justsparknl.github.io/DocuTagger/src/taskpane.html
```

Pushing to the `master` branch updates the hosted add-in automatically. End users do not need to re-sideload after updates — the task pane always loads the latest version from GitHub Pages.

---

## Project structure

```
DocuTagger/
├── src/
│   ├── taskpane.html        # Task pane UI and all add-in logic (single-file app)
│   ├── taskpane.css         # Task pane styles (Fluent UI)
│   ├── visiedosis_logo.png  # Add-in icon
│   └── visiedosis.png       # Branding image
├── manifest.xml             # Office Add-in manifest (ID, URLs, permissions)
├── .vscode/
│   ├── launch.json          # Word Desktop debug configurations
│   ├── tasks.json           # npm build/debug tasks
│   └── settings.json        # ESLint settings
├── LICENSE                  # Apache License 2.0
└── README.md
```

---

## Configuration

Key fields in `manifest.xml`:

| Field | Value |
|---|---|
| Add-in ID | `d6776528-5e24-45c6-884e-92d17858b3ab` |
| Version | `1.0.0.1` |
| Provider | Uninova B.V. |
| Default locale | `nl-NL` (Dutch) |
| Hosted URL | `https://justsparknl.github.io/DocuTagger/src/taskpane.html` |
| Permissions | `ReadWriteDocument` |
| Support URL | `https://www.uninova.nl/` |

To point the add-in at a different host, update the `SourceLocation` and `Taskpane.Url` entries in `manifest.xml`.

---

## License

Licensed under the [Apache License 2.0](LICENSE).
