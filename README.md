# officesuite-frontend

Shared static frontend (HTML/CSS/JS, no build step) originally developed for
[OfficeSuite](https://github.com/Decumano/OfficeSuite), a Tauri desktop app, and now also
consumed by [officesuite-web](https://github.com/Decumano/officesuite-web) as a git submodule.

## Structure

- `index.html`, `folio-office-suite.html` — app shell / markup
- `main.js` — application logic (only ever calls `Platform.*`, never a backend directly)
- `styles.css` — styling
- `platform.js` — the backend adapter (see below)
- `auth-gate.js` — login/register screen shown before `main.js` loads when a web backend is present
- `vendor/` — third-party libraries (see [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md))
- `assets/` — static assets

## Platform adapter

`platform.js` isolates every backend call behind a single `Platform` object (`pickWorkFolder`,
`readWorkFile`, `writeWorkFile`, `listWorkFolder`, `createWorkFolder`, `deleteWorkEntry`,
`moveWorkEntry`, `defaultFile`, `saveFile`, `exportPdf`, plus `login`/`register`/`logout`/
`currentUser`). `main.js` never talks to a backend directly.

- **On Tauri**, `platform.js` detects `window.__TAURI__` and delegates to `invoke(...)` calls
  against the host app's Rust commands (see `src-tauri/src/lib.rs` in the OfficeSuite repo for the
  reference implementation of those commands). `auth-gate.js` is a no-op on Tauri — there's no
  account system, so it loads `main.js` immediately.
- **Without Tauri**, the work-folder methods call a web backend's `/api/workspace/*` endpoints
  instead of rejecting as unavailable, `exportPdf` posts HTML to `/api/export/pdf` and downloads
  the resulting blob, and `auth-gate.js` shows a login/register form backed by `/api/auth/*` before
  `main.js` is loaded. `officesuite-web`'s `src/` is the reference implementation of that API.
- **Without any backend at all**, work-folder operations are simply unavailable and the app runs
  off in-memory/`localStorage` state instead (the existing behavior before a work folder is
  picked). `saveFile` falls back to a standard browser download instead of a native save dialog.

To embed this frontend in a new host, either provide a matching backend (Tauri commands or the
`/api/*` HTTP routes, matching the names/params in `platform.js`) or implement an equivalent by
editing `platform.js`.

## Usage in a consuming repo

Add as a submodule:

```
git submodule add https://github.com/Decumano/officesuite-frontend.git web
git submodule update --init --recursive
```
