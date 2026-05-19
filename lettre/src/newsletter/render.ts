/* HTML generation — Outlook-compatible (tables, inline CSS, MSO/VML).
 *
 * This module is the core of Lettre: given a list of blocks it produces the
 * email HTML. It is intentionally self-contained and free of UI concerns, so
 * the markup rules can be tuned without touching the task pane. */

import type {
  Block,
  HeaderData,
  HeadingData,
  TextData,
  ImageData,
  ImageTextData,
  ButtonData,
  TwoColumnData,
  CarouselData,
  AccordionData,
  DividerData,
  SpacerData,
  FooterData,
} from "./types";

const esc = (s: string): string =>
  String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

const nl2br = (s: string): string => esc(s).replace(/\n/g, "<br>");

/* Outlook's compose editor replaces fonts and colours on paste. The structures
 * that reliably survive are: <table bgcolor> cells for backgrounds, nested
 * coloured cells for text-on-colour regions, and inline styles on inner spans
 * for typography. The renderers below lean on all three. */
const renderers = {
  header: (b: HeaderData): string => `<tr><td bgcolor="${b.bg}" style="background-color:${b.bg};background:${b.bg};padding:36px 32px;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="${b.bg}" style="background-color:${b.bg};">
    <tr><td bgcolor="${b.bg}" style="background-color:${b.bg};font-family:Georgia,'Times New Roman',serif;font-size:30px;font-weight:700;line-height:1.15;letter-spacing:-0.01em;">
      <font color="${b.fg}" face="Georgia,'Times New Roman',serif"><span style="color:${b.fg};font-family:Georgia,'Times New Roman',serif;font-size:30px;font-weight:700;">${esc(b.title)}</span></font>
    </td></tr>
    <tr><td bgcolor="${b.bg}" style="background-color:${b.bg};padding-top:8px;font-family:Arial,Helvetica,sans-serif;font-size:13px;letter-spacing:0.04em;text-transform:uppercase;">
      <font color="${b.fg}" face="Arial,Helvetica,sans-serif"><span style="color:${b.fg};font-family:Arial,Helvetica,sans-serif;font-size:13px;">${esc(b.tagline)}</span></font>
    </td></tr>
  </table>
</td></tr>`,

  heading: (b: HeadingData): string => {
    const size = { h1: 26, h2: 21, h3: 17 }[b.level];
    return `<tr><td style="padding:24px 32px 6px;font-family:Arial,Helvetica,sans-serif;font-size:${size}px;font-weight:700;color:${b.color};text-align:${b.align};line-height:1.3;">${esc(b.text)}</td></tr>`;
  },

  text: (b: TextData): string =>
    `<tr><td style="padding:6px 32px 14px;font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.65;color:${b.color};text-align:${b.align};">${nl2br(b.content)}</td></tr>`,

  image: (b: ImageData): string => {
    const img = `<img src="${esc(b.src)}" alt="${esc(b.alt)}" width="536" style="display:block;width:100%;max-width:536px;height:auto;border:0;outline:none;text-decoration:none;">`;
    const wrap = b.href ? `<a href="${esc(b.href)}" style="text-decoration:none;">${img}</a>` : img;
    return `<tr><td style="padding:16px 32px;">${wrap}</td></tr>`;
  },

  /* The button is wrapped in a bgcolor table cell — Outlook strips background
   * colours from <a> tags but respects them on <td bgcolor>. The MSO branch
   * uses VML so the button keeps its shape in the Word rendering engine. */
  button: (b: ButtonData): string => `<tr><td align="${b.align}" style="padding:18px 32px;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;">
  <tr><td bgcolor="${b.bg}" align="center" style="background-color:${b.bg};background:${b.bg};border-radius:4px;padding:14px 30px;">
    <!--[if mso]>
    <v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" href="${esc(b.href)}" style="height:46px;v-text-anchor:middle;width:220px;" arcsize="10%" stroke="f" fillcolor="${b.bg}">
    <w:anchorlock/><center style="color:${b.fg};font-family:Arial,sans-serif;font-size:15px;font-weight:600;">${esc(b.text)}</center>
    </v:roundrect>
    <![endif]-->
    <!--[if !mso]><!-- -->
    <a href="${esc(b.href)}" style="color:${b.fg};text-decoration:none;font-family:Arial,Helvetica,sans-serif;font-size:15px;font-weight:600;letter-spacing:0.01em;display:inline-block;">
      <font color="${b.fg}" face="Arial,Helvetica,sans-serif"><span style="color:${b.fg};">${esc(b.text)}</span></font>
    </a>
    <!--<![endif]-->
  </td></tr>
</table>
</td></tr>`,

  /* Real HTML table — inline-block divs get collapsed by Outlook. */
  twoColumn: (b: TwoColumnData): string => `<tr><td style="padding:16px 32px;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
  <tr>
    <td width="50%" valign="top" style="width:50%;padding-right:10px;font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#333;">${nl2br(b.left)}</td>
    <td width="50%" valign="top" style="width:50%;padding-left:10px;font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#333;">${nl2br(b.right)}</td>
  </tr>
</table>
</td></tr>`,

  imageText: (b: ImageTextData): string => {
    const imgW = b.imageWidth;
    const imgCell = `<td width="${imgW}" valign="top" style="padding:${b.layout === "image-left" ? "0 16px 0 0" : "0 0 0 16px"};">
<img src="${esc(b.src)}" alt="${esc(b.alt)}" width="${imgW}" style="display:block;width:100%;max-width:${imgW}px;height:auto;border:0;outline:none;">
</td>`;
    const linkHtml =
      b.linkText && b.linkUrl
        ? `<div style="margin-top:12px;"><a href="${esc(b.linkUrl)}" style="color:#c2410c;text-decoration:none;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:600;">${esc(b.linkText)} &rarr;</a></div>`
        : "";
    const textCell = `<td valign="top" style="font-family:Arial,Helvetica,sans-serif;">
<div style="font-size:18px;font-weight:700;color:#1a1a1a;line-height:1.3;">${esc(b.heading)}</div>
<div style="font-size:14px;line-height:1.6;color:#444;margin-top:8px;">${nl2br(b.body)}</div>
${linkHtml}
</td>`;
    const cells = b.layout === "image-left" ? imgCell + textCell : textCell + imgCell;
    return `<tr><td style="padding:16px 32px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>${cells}</tr></table></td></tr>`;
  },

  /* Static gallery — Outlook-safe substitute for an interactive carousel. */
  carousel: (b: CarouselData): string => {
    const n = Math.max(1, b.items.length);
    const pct = (100 / n).toFixed(2);
    const cells = b.items
      .map(
        (item) => `
<td valign="top" width="${pct}%" style="padding:0 4px;">
<img src="${esc(item.src)}" alt="${esc(item.alt)}" style="display:block;width:100%;height:auto;border:0;outline:none;border-radius:3px;">
${item.caption ? `<div style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#666;margin-top:8px;text-align:center;line-height:1.4;letter-spacing:0.02em;">${esc(item.caption)}</div>` : ""}
</td>`,
      )
      .join("");
    return `<tr><td style="padding:16px 28px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>${cells}</tr></table></td></tr>`;
  },

  /* Styled Q&A — Outlook-safe substitute for a collapsible accordion. */
  accordion: (b: AccordionData): string => {
    const items = b.items
      .map(
        (item, i) => `
<tr><td style="padding-bottom:8px;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border:1px solid #e5e3dc;border-collapse:separate;">
<tr><td bgcolor="#f5f4ef" style="background-color:#f5f4ef;background:#f5f4ef;padding:14px 18px;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>
    <td width="36" valign="top" style="width:36px;font-family:Georgia,serif;font-size:15px;font-weight:600;"><font color="#c2410c">${String(i + 1).padStart(2, "0")}</font></td>
    <td valign="top" style="font-family:Arial,Helvetica,sans-serif;font-size:15px;font-weight:600;color:#1a1a1a;">${esc(item.question)}</td>
  </tr></table>
</td></tr>
<tr><td bgcolor="#ffffff" style="background-color:#ffffff;padding:14px 18px;font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#444;border-top:1px solid #e5e3dc;">${nl2br(item.answer)}</td></tr>
</table>
</td></tr>`,
      )
      .join("");
    return `<tr><td style="padding:16px 32px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">${items}</table></td></tr>`;
  },

  divider: (b: DividerData): string =>
    `<tr><td style="padding:10px 32px;"><div style="border-top:1px solid ${b.color};line-height:0;font-size:0;">&nbsp;</div></td></tr>`,

  spacer: (b: SpacerData): string =>
    `<tr><td style="height:${b.height}px;line-height:${b.height}px;font-size:0;">&nbsp;</td></tr>`,

  footer: (b: FooterData): string => `<tr><td bgcolor="#f5f4ef" style="background-color:#f5f4ef;background:#f5f4ef;padding:28px 32px;text-align:center;font-family:Arial,Helvetica,sans-serif;font-size:12px;line-height:1.7;color:#666;">
<div style="font-weight:600;color:#333;">${esc(b.company)}</div>
<div style="color:#666;">${esc(b.address)}</div>
<div style="margin-top:10px;"><a href="${esc(b.unsubUrl)}" style="color:#666;text-decoration:underline;">Unsubscribe</a></div>
</td></tr>`,
};

/** Render a single block to its table-row HTML. */
function renderBlock(block: Block): string {
  // The renderer for `block.type` always matches `block.data`; the index
  // access loses that correlation, so the call is asserted back.
  return (renderers[block.type] as (d: Block["data"]) => string)(block.data);
}

/** Build the complete, Outlook-ready newsletter document. */
export function generateHtml(blocks: Block[]): string {
  const body = blocks.map(renderBlock).join("\n");
  return `<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta charset="UTF-8"><meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width,initial-scale=1">
<!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><![endif]-->
<title>Newsletter</title>
</head>
<body style="margin:0;padding:0;background:#ebe8df;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#ebe8df;">
  <tr><td align="center" style="padding:28px 12px;">
    <!--[if mso | IE]><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" align="center"><tr><td><![endif]-->
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;background:#ffffff;">
${body}
    </table>
    <!--[if mso | IE]></td></tr></table><![endif]-->
  </td></tr>
</table>
</body>
</html>`;
}
