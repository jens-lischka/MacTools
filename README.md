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

## Two manifests

A manifest is only a *pointer* to where the add-in's code is hosted — it does
not contain the app. There are two:

| File | Points at | Use |
|---|---|---|
| `manifest.xml` | `https://jens-lischka.github.io/MacTools/` (GitHub Pages) | **Upload this.** Works on any machine — web, Mac, Windows — no local server. |
| `manifest.local.xml` | `https://localhost:3000/` | Local development only; needs the dev server running on *that* machine. |

## Use it (hosted — recommended)

1. In the GitHub repo: **Settings → Pages → Source → "GitHub Actions"**.
2. Push to the deploy branch (or run the *Deploy to GitHub Pages* workflow
   manually). It builds and publishes `dist/`.
3. Once the deploy succeeds, upload **`manifest.xml`** in PowerPoint
   (Web/Mac: *Add-ins → My Add-ins → Upload My Add-in*). It now loads on every
   machine with no local setup.

> The hosted URL assumes the repository is named `MacTools`. If your repo name
> differs in spelling or case, update the URLs in `manifest.xml` and
> `.github/workflows/deploy.yml` to match `https://<user>.github.io/<repo>/`.

## Develop locally (optional)

```bash
npm install
npx office-addin-dev-certs install   # one-time: trust the HTTPS dev certificate
npm start                            # runs dev server + sideloads manifest.local.xml
```

Other scripts:

```bash
npm run dev-server  # webpack dev server only (https://localhost:3000)
npm run build       # production build to dist/
npm run validate    # validate manifest.xml
npm run typecheck   # type-check without emitting
```

> **Task pane blank or "can't reach localhost"?** That means no dev server is
> running on that machine. The localhost manifest only works where the dev
> server runs. To test across several machines, use the hosted `manifest.xml`
> instead.

## Project layout

```
manifest.xml                 Hosted manifest (GitHub Pages URL) — upload this
manifest.local.xml           Localhost manifest for local development
.github/workflows/deploy.yml Builds and publishes dist/ to GitHub Pages
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
