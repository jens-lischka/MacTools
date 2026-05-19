# Lettre

A cross-platform Outlook add-in for building newsletters **inside the compose
window** and inserting them straight into the email body.

It is the task-pane successor to the standalone newsletter-builder prototype:
the same block model and Outlook-safe HTML generator, repackaged as an Office
Add-in that works on Outlook for Windows, Mac, and the web.

## How it works

- Open a new message in Outlook → **Newsletter Builder** button on the compose
  ribbon → the task pane opens.
- Add and arrange blocks (header, heading, text, image, image + text, button,
  two columns, carousel, accordion, divider, spacer, footer).
- **Insert into email** writes the generated HTML directly into the message
  body via `Office.context.mailbox.item.body.setAsync` — no copy/paste, so the
  styling is preserved rather than being mangled by the compose editor.

## Project layout

```
src/newsletter/   Core, UI-free document model
  types.ts        Block union (discriminated on `type`)
  render.ts       Outlook-safe HTML generation — the heart of Lettre
  catalog.ts      Block catalogue, defaults, factory helpers
src/office/
  mailbox.ts      Office.js bridge — writes the newsletter into the message
src/taskpane/     React task pane (Fluent UI)
src/commands/     Manifest function file
```

`src/newsletter/` is intentionally self-contained: the markup rules can be
tuned without touching the task pane, and the task pane can be reworked
without touching the generator.

## Develop

```
npm install
npm start            # sideloads manifest.xml into Outlook over https://localhost:3000
npm run typecheck
npm run build        # production bundle into dist/
```

The manifest points at `https://localhost:3000` for local development. To
deploy, host the contents of `dist/` on an HTTPS origin and replace every
`https://localhost:3000` in `manifest.xml` with that origin.

## Roadmap

- **Usage analytics** — tracking newsletter opens and recipient interaction is
  planned but not yet implemented. It needs a hosted endpoint (open pixels and
  click-redirects cannot be served from Outlook itself), so it will land
  alongside a small collector service.
