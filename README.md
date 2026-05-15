# MacTools for PowerPoint

A cross-platform PowerPoint add-in (web, Mac, Windows) — a rebuild of the
legacy Windows-only VSTO "OWG Toolbox" on the Office.js platform.

All features live in a single **task pane**, organised by category.

## Status

Scaffold + Phase 2 reference implementation.

- ✅ Task-pane shell with categorised, collapsible navigation
- ✅ Settings store (per-document + per-user roaming)
- ✅ **Alignment & Distribution** — fully implemented
- 🚧 All other categories — listed with feasibility ratings; handlers land in
  later phases

See [`docs/powerpoint-cross-platform-overview.md`](docs/powerpoint-cross-platform-overview.md)
for the full feature inventory and feasibility analysis.

## Tech stack

TypeScript · React 18 · Fluent UI v9 · webpack · Office.js (`PowerPointApi`)

## Develop

```bash
npm install
npx office-addin-dev-certs install   # one-time: trust the HTTPS dev certificate
npm start                            # sideload into PowerPoint and open the task pane
```

Other scripts:

```bash
npm run dev-server  # webpack dev server only (https://localhost:3000)
npm run validate    # validate the manifest
npm run typecheck   # type-check without emitting
npm run build       # production build to dist/
```

> **Task pane not loading?** Office only loads a task pane over HTTPS with a
> *trusted* certificate. Run `npx office-addin-dev-certs install` once (it adds
> the dev cert to your OS trust store), then restart `npm start`. On Office on
> the web, also confirm `https://localhost:3000/taskpane.html` opens in a
> browser without a certificate warning.

## Project layout

```
manifest.xml                 Office Add-in manifest (PowerPoint, all platforms)
src/
  taskpane/
    index.tsx                Office.onReady bootstrap
    components/              App shell, category panels, feature list
    categories/index.ts      Category + feature registry (the roadmap)
  commands/                  Function-file entry point (reserved)
  lib/
    types.ts                 Feature / Category model
    settings.ts              Per-document + roaming settings store
    powerpoint.ts            Selection + geometry helpers
    alignment.ts             Alignment & distribution operations
docs/                        Feasibility overview
```

## Architecture notes

- **Categories** are data (`src/taskpane/categories/index.ts`). Each feature
  carries a feasibility `rating`; a feature with no `run` handler renders
  disabled, so the roadmap is visible in-product.
- A category may render a generic feature-button list or a **custom panel**
  when it needs richer UI (e.g. Alignment's target selector).
- **Power Align**: the legacy guide- and modifier-key-based modes are not
  expressible on Office.js. They are replaced by an explicit in-pane *target*
  choice — align relative to the selection, first/last shape, or the slide.
