/**
 * Text & paragraph operations — Phase 3.
 *
 * Operate on the text-capable shapes in the selection. Lines, pictures and
 * other shapes without a text frame are filtered out by type so they cannot
 * break the batch.
 */

/**
 * True for shape types that carry a text frame. Defined as a function (not a
 * module-level constant) so the `PowerPoint` enum is read lazily — referencing
 * it at module load can run before Office.js has defined the namespace.
 */
function isTextShape(type: PowerPoint.Shape["type"]): boolean {
  return (
    type === PowerPoint.ShapeType.geometricShape ||
    type === PowerPoint.ShapeType.textBox ||
    type === PowerPoint.ShapeType.placeholder
  );
}

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

export interface TextStyle {
  size: number;
  bold: boolean;
}

/** Apply a simple font style (size + weight) to the selected text shapes. */
export async function applyTextStyle(style: TextStyle): Promise<void> {
  await withTextShapes((shapes) => {
    shapes.forEach((s) => {
      const font = s.textFrame.textRange.font;
      font.size = style.size;
      font.bold = style.bold;
    });
  });
}

/** Concatenate the text of every selected text shape into the first one. */
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
    ranges.slice(1).forEach((r) => {
      r.text = "";
    });
  });
}

/** Split the lines of one selected text shape into separate text boxes. */
export async function splitText(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/type,items/left,items/top,items/width,items/height");
    const slides = context.presentation.getSelectedSlides();
    slides.load("items/id");
    await context.sync();

    const textShapes = selected.items.filter((s) =>
      isTextShape(s.type),
    );
    if (textShapes.length !== 1) {
      throw new Error("Select exactly one text shape to split.");
    }
    if (slides.items.length === 0) throw new Error("No slide is in view.");

    const shape = textShapes[0];
    const range = shape.textFrame.textRange;
    range.load("text");
    await context.sync();

    const lines = range.text
      .split(/[\r\n\v\f]+/)
      .map((l) => l.trim())
      .filter((l) => l.length > 0);
    if (lines.length < 2) {
      throw new Error("The selected shape has only one line of text.");
    }

    const lineHeight = shape.height / lines.length;
    const slide = slides.items[0];
    lines.slice(1).forEach((line, i) => {
      slide.shapes.addTextBox(line, {
        left: shape.left,
        top: shape.top + (i + 1) * lineHeight,
        width: shape.width,
        height: lineHeight,
      });
    });
    range.text = lines[0];
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
