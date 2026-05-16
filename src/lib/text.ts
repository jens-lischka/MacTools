/**
 * Text & paragraph operations — Phase 3.
 *
 * Operate on the text-capable shapes in the selection. Lines, pictures and
 * other shapes without a text frame are filtered out by type so they cannot
 * break the batch.
 */

const TEXT_SHAPE_TYPES: ReadonlyArray<PowerPoint.Shape["type"]> = [
  PowerPoint.ShapeType.geometricShape,
  PowerPoint.ShapeType.textBox,
  PowerPoint.ShapeType.placeholder,
];

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
      TEXT_SHAPE_TYPES.includes(s.type),
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
