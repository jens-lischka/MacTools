/**
 * Text & paragraph operations — Phase 3.
 *
 * Operate on the text-capable shapes in the selection. Lines, pictures and
 * other shapes without a text frame are filtered out by type so they cannot
 * break the batch.
 */

import { isTextShape } from "./powerpoint";

async function withTextShapes(
  callback: (
    shapes: PowerPoint.Shape[],
    context: PowerPoint.RequestContext,
  ) => void | Promise<void>,
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id,items/type");
    await context.sync();

    const textShapes = selected.items.filter((s) =>
      isTextShape(s.type),
    );
    if (textShapes.length === 0) {
      throw new Error("Select at least one shape that can contain text.");
    }

    await callback(textShapes, context);
    await context.sync();
  });
}

/** Apply an autosize mode to the selected text shapes. */
export async function setAutoSize(mode: PowerPoint.ShapeAutoSize): Promise<void> {
  await withTextShapes((shapes) => {
    shapes.forEach((s) => {
      s.textFrame.autoSizeSetting = mode;
    });
  });
}

/** Turn word wrap on or off for the selected text shapes. */
export async function setWordWrap(wrap: boolean): Promise<void> {
  await withTextShapes((shapes) => {
    shapes.forEach((s) => {
      s.textFrame.wordWrap = wrap;
    });
  });
}

export interface Margins {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

/** Set the internal text-box margins (points) of the selected text shapes. */
export async function setMargins(margins: Margins): Promise<void> {
  await withTextShapes((shapes) => {
    shapes.forEach((s) => {
      s.textFrame.leftMargin = margins.left;
      s.textFrame.rightMargin = margins.right;
      s.textFrame.topMargin = margins.top;
      s.textFrame.bottomMargin = margins.bottom;
    });
  });
}

/** Set paragraph horizontal alignment for all text in the selected shapes. */
export async function setParagraphAlignment(
  alignment: PowerPoint.ParagraphHorizontalAlignment,
): Promise<void> {
  await withTextShapes((shapes) => {
    shapes.forEach((s) => {
      s.textFrame.textRange.paragraphFormat.horizontalAlignment = alignment;
    });
  });
}

/** Replace every run of line breaks with a single space. */
export async function clearLineBreaks(): Promise<void> {
  await withTextShapes(async (shapes, context) => {
    const ranges = shapes.map((s) => s.textFrame.textRange);
    ranges.forEach((r) => r.load("text"));
    await context.sync();

    ranges.forEach((r) => {
      r.text = r.text.replace(/[\r\n\v\f]+/g, " ").trim();
    });
  });
}

/** Delete all text in the selected shapes. */
export async function clearText(): Promise<void> {
  await withTextShapes((shapes) => {
    shapes.forEach((s) => {
      s.textFrame.textRange.text = "";
    });
  });
}

/** Show or hide paragraph bullets in the selected text shapes. */
export async function setBullets(visible: boolean): Promise<void> {
  await withTextShapes((shapes) => {
    shapes.forEach((s) => {
      s.textFrame.textRange.paragraphFormat.bulletFormat.visible = visible;
    });
  });
}

interface BulletLevel {
  type: PowerPoint.BulletFormat["type"];
  style: PowerPoint.BulletFormat["style"];
  visible: PowerPoint.BulletFormat["visible"];
}

/** Character ranges of each paragraph in `text` (handles \r, \n and \r\n). */
function paragraphRanges(text: string): { start: number; length: number }[] {
  const ranges: { start: number; length: number }[] = [];
  let start = 0;
  let i = 0;
  while (i <= text.length) {
    const ch = text[i];
    if (i === text.length || ch === "\r" || ch === "\n") {
      ranges.push({ start, length: i - start });
      if (ch === "\r" && text[i + 1] === "\n") i += 1;
      i += 1;
      start = i;
    } else {
      i += 1;
    }
  }
  return ranges;
}

/** Non-empty paragraph sub-ranges of a text range. */
function paragraphSubranges(
  range: PowerPoint.TextRange,
  text: string,
): PowerPoint.TextRange[] {
  return paragraphRanges(text)
    .filter((p) => p.length > 0)
    .map((p) => range.getSubstring(p.start, p.length));
}

/**
 * Copy the body placeholder's per-indent-level bullet formatting from the
 * slide master onto every paragraph of the selected text shapes, matched by
 * indent level.
 *
 * The JavaScript API only exposes a bullet's type, numbering style and
 * visibility — not its character, colour, size or indent — so this aligns the
 * bullet *kind* per level, not the master's exact bullet appearance.
 */
export async function matchBulletLevels(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id,items/type");
    await context.sync();

    const textShapes = selected.items.filter((s) => isTextShape(s.type));
    if (textShapes.length === 0) {
      throw new Error("Select at least one shape that can contain text.");
    }

    // Locate the body placeholder on the slide master.
    const masterShapes = textShapes[0].getParentSlideMaster().shapes;
    masterShapes.load("items/type");
    await context.sync();

    const placeholders = masterShapes.items.filter(
      (s) => s.type === PowerPoint.ShapeType.placeholder,
    );
    placeholders.forEach((p) => p.placeholderFormat.load("type"));
    await context.sync();

    const body = placeholders.find((p) => {
      const t = p.placeholderFormat.type;
      return (
        t === PowerPoint.PlaceholderType.body ||
        t === PowerPoint.PlaceholderType.content ||
        t === PowerPoint.PlaceholderType.verticalBody
      );
    });
    if (!body) {
      throw new Error("The slide master has no body placeholder to copy bullets from.");
    }

    // Read the master body placeholder's bullet format per indent level.
    const masterRange = body.textFrame.textRange;
    masterRange.load("text");
    await context.sync();

    const masterFormats = paragraphSubranges(masterRange, masterRange.text).map(
      (range) => {
        range.paragraphFormat.load("indentLevel");
        range.paragraphFormat.bulletFormat.load("type,style,visible");
        return range.paragraphFormat;
      },
    );
    await context.sync();

    const levels = new Map<number, BulletLevel>();
    masterFormats.forEach((pf) => {
      if (!levels.has(pf.indentLevel)) {
        levels.set(pf.indentLevel, {
          type: pf.bulletFormat.type,
          style: pf.bulletFormat.style,
          visible: pf.bulletFormat.visible,
        });
      }
    });
    if (levels.size === 0) {
      throw new Error("The master body placeholder has no bullet levels to copy.");
    }

    // Read every target paragraph's indent level.
    const targetRanges = textShapes.map((s) => s.textFrame.textRange);
    targetRanges.forEach((r) => r.load("text"));
    await context.sync();

    const targetParagraphs = targetRanges.flatMap((r) =>
      paragraphSubranges(r, r.text),
    );
    targetParagraphs.forEach((range) =>
      range.paragraphFormat.load("indentLevel"),
    );
    await context.sync();

    // Apply the master level's bullet format to each paragraph.
    targetParagraphs.forEach((range) => {
      const pf = range.paragraphFormat;
      const level = levels.get(pf.indentLevel) ?? levels.get(0);
      if (!level) return;
      if (level.visible !== null) pf.bulletFormat.visible = level.visible;
      if (level.type !== null) pf.bulletFormat.type = level.type;
      if (level.style !== null) pf.bulletFormat.style = level.style;
    });
    await context.sync();
  });
}

export interface TextStyle {
  name: string;
  size: number;
  bold: boolean;
}

/**
 * Apply a font style (font, size, weight). If text is selected inside a
 * shape, the style is applied to that selection only (so a single line can
 * be styled); otherwise it applies to the whole text of each selected shape.
 * Line spacing is left untouched — the JS API does not expose it.
 */
export async function applyTextStyle(style: TextStyle): Promise<void> {
  await PowerPoint.run(async (context) => {
    const selectedRange = context.presentation.getSelectedTextRangeOrNullObject();
    selectedRange.load("isNullObject");
    await context.sync();

    if (!selectedRange.isNullObject) {
      const font = selectedRange.font;
      font.name = style.name;
      font.size = style.size;
      font.bold = style.bold;
      await context.sync();
      return;
    }

    const selected = context.presentation.getSelectedShapes();
    selected.load("items/type");
    await context.sync();
    const textShapes = selected.items.filter((s) => isTextShape(s.type));
    if (textShapes.length === 0) {
      throw new Error("Select a shape, or select the text to style.");
    }
    textShapes.forEach((s) => {
      const font = s.textFrame.textRange.font;
      font.name = style.name;
      font.size = style.size;
      font.bold = style.bold;
    });
    await context.sync();
  });
}

/**
 * Concatenate the text of every selected text shape into the first one.
 * Selected text boxes that are emptied by the merge are deleted.
 */
export async function mergeText(): Promise<void> {
  await withTextShapes(async (shapes, context) => {
    if (shapes.length < 2) {
      throw new Error("Select two or more text shapes to merge.");
    }
    const ranges = shapes.map((s) => s.textFrame.textRange);
    ranges.forEach((r) => r.load("text"));
    await context.sync();

    const merged = ranges
      .map((r) => r.text.trim())
      .filter((t) => t.length > 0)
      .join("\n");
    ranges[0].text = merged;
    shapes.slice(1).forEach((s, i) => {
      if (s.type === PowerPoint.ShapeType.textBox) {
        s.delete();
      } else {
        ranges[i + 1].text = "";
      }
    });
  });
}

/** Split one selected text shape into one text box per paragraph. */
export async function splitText(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/type,items/left,items/top,items/width,items/height");
    const slides = context.presentation.getSelectedSlides();
    slides.load("items/id");
    await context.sync();

    const textShapes = selected.items.filter((s) => isTextShape(s.type));
    if (textShapes.length !== 1) {
      throw new Error("Select exactly one text shape to split.");
    }
    if (slides.items.length === 0) throw new Error("No slide is in view.");

    const shape = textShapes[0];
    const range = shape.textFrame.textRange;
    range.load("text");
    await context.sync();

    // Split on paragraph breaks only — soft line breaks stay within a
    // paragraph, so each resulting text box holds exactly one paragraph.
    const paragraphs = paragraphRanges(range.text)
      .map((p) => range.text.substr(p.start, p.length))
      .filter((p) => p.trim().length > 0);
    if (paragraphs.length < 2) {
      throw new Error("The selected shape has only one paragraph.");
    }

    const boxHeight = shape.height / paragraphs.length;
    const slide = slides.items[0];
    paragraphs.slice(1).forEach((paragraph, i) => {
      slide.shapes.addTextBox(paragraph, {
        left: shape.left,
        top: shape.top + (i + 1) * boxHeight,
        width: shape.width,
        height: boxHeight,
      });
    });
    range.text = paragraphs[0];
    await context.sync();
  });
}

/** Insert a character at the current text cursor / selection. */
export async function insertSpecialCharacter(character: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    const range = context.presentation.getSelectedTextRangeOrNullObject();
    range.load("isNullObject");
    await context.sync();
    if (range.isNullObject) {
      throw new Error("Place the cursor inside a text box first.");
    }
    range.text = character;
    await context.sync();
  });
}
