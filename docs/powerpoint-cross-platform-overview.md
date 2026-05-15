# PowerPoint Add-in — Cross-Platform Rebuild Overview

## 1. Context

The source project (`dtp-ot-win-mmc-owg-develop`, internally **ToolsPowerPoint** /
"OWG Toolbox") is a **VSTO add-in** written in C#. VSTO is Windows-only and
talks to PowerPoint through the **COM object model**. It cannot run on the web
or on macOS.

To make these features available on **web, Mac and Windows**, they must be
rebuilt as an **Office Add-in** (Office.js / "Office JavaScript API"). Office
Add-ins are HTML/JS/CSS, run inside an embedded browser, and use the same
JavaScript API on every platform — so a single codebase covers all three.

This document inventories every PowerPoint feature in the legacy add-in and
rates how realistically each one can be rebuilt on Office.js today. It is an
**assessment for planning**, not an implementation.

---

## 2. The core constraint: Office.js is not COM

VSTO exposes essentially the *entire* PowerPoint application. Office.js exposes
a deliberately smaller, sandboxed surface. The practical gaps that shape this
whole project are:

| Area | VSTO (legacy) | Office.js (web/Mac/Win) |
|---|---|---|
| Shape geometry, fill, line, text, fonts | Full | Full ✅ |
| Selection (shapes / slides / text) | Full | Full ✅ |
| Add shapes / text boxes / lines / tables | Full | Full ✅ |
| Slide tags / custom metadata | Full | Full ✅ (`tags`) |
| Z-order (bring to front / send to back) | Full | **Missing** ❌ |
| Slide guides | Full | **Missing** ❌ |
| Presentation **sections** | Full | **Missing** ❌ |
| Render / export a slide as an image | Full | **Missing** ❌ |
| Charts / SmartArt object model | Full | **Missing** ❌ |
| Picture crop & recolor (grayscale) | Full | **Missing** ❌ |
| Shape geometry-adjustment handles (e.g. chevron angle) | Full | **Missing** ❌ |
| Multiple open documents / windows | Full | **Missing** ❌ |
| Global keyboard shortcut binding | Full | **Missing** ❌ |
| Connectors with real connection sites | Full | Partial 🟡 |
| Table structure editing (insert/move/delete rows) | Full | Partial 🟡 |
| Modifier-key state at click time (Ctrl/Alt/Shift) | Full | **Missing** ❌ |

Everything below is rated against this reality:

- ✅ **Feasible now** — Office.js supports it directly; cross-platform with no asterisks.
- 🟡 **Partial** — the useful core works, but with reduced scope or a workaround
  (often: do the math ourselves on shape coordinates instead of calling a
  native command).
- ❌ **Blocked** — Office.js currently lacks the API; cannot be rebuilt faithfully
  yet. Listed so nothing is silently dropped.

---

## 3. Proposed shape of the new add-in

- **One task pane**, opened from a single ribbon button. All features live
  inside it (per your request) — no scattered ribbon groups.
- The pane is **organised by category** (Alignment, Size & Position, Text, …),
  each category a collapsible section, matching the structure in §4.
- Recommended stack: **TypeScript + React + Fluent UI v9**, which is the
  standard for modern Office Add-ins and gives a native Office look on all
  platforms.
- **Settings** (the legacy "Tools Settings": templates, colours, label text,
  Harvey styles, etc.) are stored with `Office.context.document.settings`
  (per-document) and/or `OfficeRuntime.storage` (per-user roaming) instead of
  the registry/files used by VSTO.
- A single **unified manifest** targets PowerPoint on web, Mac and Windows.

This gives one button → one categorised pane → every feasible feature, working
identically on all three platforms.

---

## 4. Feature inventory by category

Counts are approximate distinct features (legacy menus/duplicates collapsed).

### 4.1 Alignment & Distribution
| Feature | Rating | Notes |
|---|---|---|
| Align Left / Center / Right / Top / Middle / Bottom | ✅ | Compute from selected-shape geometry. |
| Align to Slide / Selection / First / Last selected | ✅ | Slide size and shape bounds are both available. |
| Distribute Horizontally / Vertically | ✅ | Geometry math. |
| Distribute from Top / Bottom / Left / Right | ✅ | Geometry math. |
| Set Custom Gap / gap presets (0–10) | ✅ | Geometry math. |
| Pick up Horizontal / Vertical Gap | ✅ | Read gap, re-apply. |
| Padding / Center to Slide | ✅ | Geometry math. |
| **Power Align** (guide-aware align, modifier-key modes) | 🟡 | Plain align works ✅. The "power" part — align to **guides** and the Ctrl/Alt/Shift modifier modes — is ❌ (no guide API, no modifier state). Replace with explicit in-pane options. |
| Align to Guides | ❌ | No slide-guide API. |
| Align Over (align shape onto a table cell) | 🟡 | Works for shape-to-shape; table-cell targeting is limited by the table API. |

### 4.2 Size & Position
| Feature | Rating | Notes |
|---|---|---|
| Match Size / Width / Height | ✅ | Copy dimensions across selection. |
| Stretch Top / Right / Left / Bottom | ✅ | Geometry math. |
| Fill Top / Right / Left / Bottom | ✅ | Geometry math. |
| Fill Horizontal / Vertical Gap | ✅ | Geometry math. |
| Height↔Width | ✅ | Geometry math. |
| Scale to Value | ✅ | Geometry math. |
| Straighten Line | ✅ | Reset line endpoints. |
| Unify Shapes / Unify Arrows | 🟡 | Size/line unification ✅; arrow-style unification ✅; anything depending on geometry adjustments ❌. |
| Chevrons → 90° / 120°, Chevron Width | ❌ | Needs shape geometry-adjustment handles, not exposed. |
| Crop to Circle | ❌ | Picture crop API not exposed. |
| Lock / Unlock Aspect Ratio | ❌ | `lockAspectRatio` not exposed; can only emulate during an in-pane resize. |

### 4.3 Selection ("Select Same")
| Feature | Rating | Notes |
|---|---|---|
| Select by Fill / Outline colour, Outline weight / dash | ✅ | Scan shapes, compare, `setSelectedShapes`. |
| Select by Type / Height / Width / Size | ✅ | Same approach. |
| Select by Text colour / size / Font name | ✅ | Same approach. |
| Select by Position (top/left/right/bottom) | ✅ | Same approach. |
| Select by Line Arrow begin / end | ✅ | `lineFormat` arrow heads are readable. |
| Show All / Hide Objects | ❌ | Shape visibility toggling not exposed. |

### 4.4 Text & Paragraph
| Feature | Rating | Notes |
|---|---|---|
| Text Styles (Body / Heading / Subheading / Display 1–3) | ✅ | Apply font + paragraph formatting from settings. |
| Resize Shape to Fit Text | ✅ | `textFrame.autoSizeSetting`. |
| Wrap Text | ✅ | `textFrame.wordWrap`. |
| Text-box Margins (increase/decrease + 9 presets) | ✅ | `textFrame` margin properties. |
| Clear Line Breaks / Clear Text in Selection | ✅ | String operations on `textRange`. |
| Merge Text / Split Text / Copy Text | ✅ | Read/write text across shapes. |
| Clear to Basic / Clear to Shape / Clear Bold from Marsh Serif | ✅ | Reset formatting to defaults. |
| Special Characters (French/German/Italian/Spanish, Symbols, Special) | ✅ | Insert characters into the text range. |
| Clean (text cleanup) | ✅ | String normalisation. |
| Fix Bullets / Bullet Lists 1–3 / No Bullets | 🟡 | Bullet formatting in the paragraph API is partial — basic bullet on/off works; rich custom bullet styles may be limited. |
| Paragraph Spacing — Increase / Decrease / Reset | 🟡 | Space-before/after support in the paragraph API is partial; verify against target API version. |

### 4.5 Shape Creation / Insert
| Feature | Rating | Notes |
|---|---|---|
| Insert rectangle / rounded rectangle / oval / circle | ✅ | `shapes.addGeometricShape`. |
| Insert callout / bracket / brace / caret / cross / tick | ✅ | Geometric shape types. |
| Insert text box | ✅ | `shapes.addTextBox`. |
| Insert line / arrow line / arc line | ✅ | `shapes.addLine`. |
| Highlight shape | ✅ | Semi-transparent shape. |
| Label line / Numbered Circle | ✅ | Composite of line/shape + text. |
| Multiply Shape | ✅ | Duplicate + offset. |
| Group as Rows / Columns | ✅ | Arrange + `group`. |
| Connector / Connector (arrow), Auto-Connect, Connector Wizard | 🟡 | A line *between* two shapes is fine; true connectors that re-route when shapes move (connection sites) are not fully exposed. |

### 4.6 Shape Properties — Swap / Pick up & Apply
| Feature | Rating | Notes |
|---|---|---|
| Swap Position / Text / Font | ✅ | Read both, write crossed-over. |
| Swap Fill & Outline / Line Type / Arrow Type / Oval Type / Outline Type / Line Weight | ✅ | All via `fill` / `lineFormat`. |
| Pick up & Apply Object Size & Position (+ Size / Position / W / H) | ✅ | Hold values in pane state or document settings. |
| Apply to Matching Objects | ✅ | Combine with §4.3 matching logic. |
| Pick up / Apply Table Size & Formatting | 🟡 | Limited by the table API; cell-level styling is partial. |

### 4.7 Effects
| Feature | Rating | Notes |
|---|---|---|
| Drop Shadow / Remove Effects | ❌ | Shadow & effect formatting not exposed in Office.js. |
| Fix Grayscale | ❌ | Picture recolour not exposed. |

### 4.8 Tables
| Feature | Rating | Notes |
|---|---|---|
| Format Table / Table Heading / Row Heading / Table Text | 🟡 | Cell text & fill are reachable; full table-style application is limited. |
| Add Column Left/Right, Add Row | 🟡 | Row/column add exists in recent API versions; confirm against target version. |
| Move Column / Row, Remove Last Column / Row | 🟡 | Move is ❌ (no reorder API); delete is partial. |
| Transpose Table | ❌ | No structural transpose; would require rebuilding the table. |
| Split Table | ❌ | No split API. |
| Convert Table to Text | 🟡 | Read cell values ✅, build a text box ✅, but original table removal/layout is approximate. |
| Optimize Table Width | 🟡 | Possible if column widths are writable in the target API version. |

### 4.9 Slide & Presentation Management
| Feature | Rating | Notes |
|---|---|---|
| Table of Contents (Style 1 / 2) | 🟡 | Reading slide titles ✅ and building a TOC slide ✅; clickable links/precise styling partial. |
| Export Slide Titles | ✅ | Read titles, output text/CSV. |
| Sticky Notes — add, place inside/outside slide, add initials/timestamp | 🟡 | Note shape + metadata in `tags` ✅; placement ✅. |
| Sticky Notes — Show All / Hide All / jump Next-Previous | 🟡 | Navigation ✅; show/hide ❌ (no shape visibility API) — would delete/recreate instead. |
| Paste on Slides | 🟡 | Depends on clipboard access, which is constrained in the sandbox. |
| Export as Pictures (96/150/300 dpi, PNG/JPG) | ❌ | No slide-to-image render/export API. |
| Move Slides to Unused Section | ❌ | No sections API. |
| Move Section Before / After, Move to Beginning / End | ❌ | No sections API (slide-level reordering ✅, section-level ❌). |
| Close All (documents) | ❌ | No multi-document/window control. |
| Slide Guides (2/3/4/5/6/12 columns, spacing, delete) | ❌ | No slide-guide API. |

### 4.10 Special Shapes (template-driven)
| Feature | Rating | Notes |
|---|---|---|
| Slide Title / Conclusion / Footnote / Ghost / Labels | ✅ | Insert pre-formatted text boxes driven by stored settings. |
| Harvey Balls (0 / 25 / 50 / 75 / 100 % + custom) | ✅ | Draw with shapes (pie/partial fills). |

### 4.11 Utilities
| Feature | Rating | Notes |
|---|---|---|
| CAGR (compound annual growth rate) | ✅ | Pure calculation + insert result. |
| Replace Fonts | ✅ | Scan shapes, reassign font names. |
| Optimize Font Size | 🟡 | Resizing-to-fit logic works; depends on text-metrics accuracy. |
| File Size | 🟡 | Obtainable via `document.getFileAsync` (size of the sliced file). |
| Airplane Mode (downscale/flash images to shrink file) | ❌ | Picture compression/recolour not exposed. |
| Conversion Assistant — Apply New Colours / Optimize Font | 🟡 | Recolouring shapes & fonts ✅; bulk re-theming partial. |
| Conversion Assistant — Apply Custom Template / Batch Convert | ❌ | No API to swap the presentation template/master. |

### 4.12 Settings & Infrastructure
| Feature | Rating | Notes |
|---|---|---|
| Tools Settings (customise / generate / import / clear) | ✅ | Rebuild on `document.settings` + `OfficeRuntime.storage`. |
| Toolbox / Toolbox Designer | ➖ | Obsolete — the new task pane *is* the toolbox. |
| Shortcut Manager + global keyboard shortcuts | ❌ | Add-ins cannot bind arbitrary global PowerPoint shortcuts; a limited fixed set of add-in shortcuts is possible. |
| Online Help | ✅ | Link out. |

---

## 5. Summary

| Rating | Meaning | Approx. count |
|---|---|---|
| ✅ Feasible now | Rebuildable cleanly, cross-platform | ~70 features |
| 🟡 Partial | Core works; reduced scope or workaround needed | ~25 features |
| ❌ Blocked | Office.js API missing today | ~20 features |

**Headline:** the **large majority** of the toolbox — all of Alignment,
Distribution, Size & Position, Selection, most Text/Paragraph work, Shape
creation, and Swap/Pick-up/Apply — can be rebuilt to run identically on web,
Mac and Windows.

The **blocked** items cluster around four missing API areas: **slide guides**,
**presentation sections**, **slide-image export**, and **picture
crop/recolour/effects**. These should be parked, not redesigned, until/unless
Microsoft extends the API.

## 6. Suggested phasing

1. **Phase 1 — Foundation:** task-pane shell, category navigation, settings
   storage, selection plumbing.
2. **Phase 2 — Geometry features (highest ROI):** Alignment, Distribution,
   Size & Position, Select Same. All ✅, all pure geometry, immediate value.
3. **Phase 3 — Text & shapes:** Text styles, margins, text cleanup, shape
   insertion, Swap / Pick-up & Apply, Special Shapes, Harvey balls.
4. **Phase 4 — Partial features:** Tables, Sticky Notes, TOC, Conversion
   Assistant colour/font tools, bullets & paragraph spacing — implemented to
   the extent the API allows.
5. **Backlog — Blocked:** guides, sections, image export, effects — revisit as
   the Office.js API evolves.

## 7. Open questions for you

- Confirm the **minimum Office.js requirement set / version** to target — a few
  🟡 items (paragraph spacing, table rows) depend on it.
- For **Power Align**, since guide-based modes are blocked, do you want explicit
  in-pane controls (target = slide / selection / first / last) as the replacement?
- Should **settings** be per-document, per-user roaming, or both?
- Is a **macOS + Windows desktop** parity check enough, or must the **web**
  build be first-class from day one (it can be — just confirms test priority)?
