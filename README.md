# MacTools for PowerPoint

A cross-platform PowerPoint add-in (web, Mac, Windows) — a rebuild of the
legacy Windows-only VSTO "OWG Toolbox" on the Office.js platform.

All features live in a single **task pane**, organised by category.

## Status

All feasible features implemented, with runtime API capability gating.

- ✅ **Alignment & Distribution**
- ✅ **Size & Position** — match size, scale, stretch, straighten lines
- ✅ **Select Same** — fill/outline colour & weight, type, size, font,
  position; hide / show shapes *(needs API 1.10)*
- ✅ **Text & Paragraph** — autofit, wrap, margins, alignment, clear text,
  bullets, text styles, merge/split text, special characters
- ✅ **Shape Creation** — insert shapes, line, text box, numbered circle,
  lay out as row/column
- ✅ **Swap / Pick up & Apply** — swap position/size/fill&outline, pick up
  & apply geometry
- ✅ **Tables** — insert, add/remove rows & columns, table-to-text
  *(needs API 1.8)*
- ✅ **Slides** — export titles, table of contents, sticky notes
- ✅ **Special Shapes** — title, conclusion, footnote, ghost, label
- ✅ **Utilities** — replace fonts, CAGR, file size
- ✅ **Settings** — slide size, default initials, online help
- ⛔ Remaining items have no JavaScript API at all; each is listed in-product
  with the reason.

### API capability gating

The manifest keeps a low `MinVersion` so the add-in installs on every client.
Features that need a newer API set (e.g. Tables → 1.8, Hide/Show → 1.10)
check `Office.context.requirements.isSetSupported` at runtime and disable
themselves cleanly on clients that don't support them, rather than failing
when invoked. The detected API level is shown in the task-pane header.

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
