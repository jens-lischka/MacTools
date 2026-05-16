/**
 * "Select Same" operations.
 *
 * The user selects one reference shape; the add-in then selects every shape
 * on the current slide that matches it on the chosen property.
 *
 * Properties split into three groups by how they must be loaded:
 *  - type / size / position need only universally-safe geometry (one pass).
 *  - fill / outline are loaded only on the shape types that support them.
 *  - font is loaded only on text-capable shape types.
 */

export type SameProperty =
  | "fillColor"
  | "lineColor"
  | "lineWeight"
  | "shapeType"
  | "size"
  | "fontName"
  | "positionTop"
  | "positionLeft"
  | "positionRight"
  | "positionBottom";

/** Geometry tolerance in points. */
const SIZE_TOLERANCE = 0.5;
const WEIGHT_TOLERANCE = 0.01;

/** True for shape types that carry a text frame. A function (not a
 *  module-level constant) so the `PowerPoint` enum is read lazily. */
function isTextShape(type: PowerPoint.Shape["type"]): boolean {
  return (
    type === PowerPoint.ShapeType.geometricShape ||
    type === PowerPoint.ShapeType.textBox ||
    type === PowerPoint.ShapeType.placeholder
  );
}

function near(a: number, b: number, tolerance: number): boolean {
  return Math.abs(a - b) < tolerance;
}

function edgeValue(shape: PowerPoint.Shape, property: SameProperty): number {
  switch (property) {
    case "positionTop":
      return shape.top;
    case "positionLeft":
      return shape.left;
    case "positionRight":
      return shape.left + shape.width;
    case "positionBottom":
      return shape.top + shape.height;
    default:
      return NaN;
  }
}

export async function selectSame(property: SameProperty): Promise<void> {
  await PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items/id");
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items/id");
    await context.sync();

    if (selectedShapes.items.length !== 1) {
      throw new Error("Select exactly one reference shape first.");
    }
    if (selectedSlides.items.length === 0) {
      throw new Error("No slide is in view.");
    }

    const referenceId = selectedShapes.items[0].id;
    const slide = selectedSlides.items[0];
    const shapes = slide.shapes;
    shapes.load("items/id,items/type,items/left,items/top,items/width,items/height");
    await context.sync();

    const all = shapes.items;
    const reference = all.find((s) => s.id === referenceId);
    if (!reference) {
      throw new Error("The reference shape is not on the current slide.");
    }

    let matchIds: string[];

    if (property === "shapeType") {
      matchIds = all.filter((s) => s.type === reference.type).map((s) => s.id);
    } else if (property === "size") {
      matchIds = all
        .filter(
          (s) =>
            near(s.width, reference.width, SIZE_TOLERANCE) &&
            near(s.height, reference.height, SIZE_TOLERANCE),
        )
        .map((s) => s.id);
    } else if (
      property === "positionTop" ||
      property === "positionLeft" ||
      property === "positionRight" ||
      property === "positionBottom"
    ) {
      const refEdge = edgeValue(reference, property);
      matchIds = all
        .filter((s) => near(edgeValue(s, property), refEdge, SIZE_TOLERANCE))
        .map((s) => s.id);
    } else if (property === "fontName") {
      matchIds = await selectSameFont(context, all, referenceId);
    } else {
      matchIds = await selectSameStyle(context, all, referenceId, property);
    }

    if (matchIds.length === 0) {
      throw new Error("No matching shapes were found on this slide.");
    }
    slide.setSelectedShapes(matchIds);
    await context.sync();
  });
}

/** Fill / outline matching — second pass over the type-compatible subset. */
async function selectSameStyle(
  context: PowerPoint.RequestContext,
  all: PowerPoint.Shape[],
  referenceId: string,
  property: "fillColor" | "lineColor" | "lineWeight",
): Promise<string[]> {
  const wantsFill = property === "fillColor";

  // Lines have an outline but no fill; restrict each comparison to the shape
  // types that genuinely carry the property.
  const compatible = all.filter((s) =>
    wantsFill
      ? s.type === PowerPoint.ShapeType.geometricShape
      : s.type === PowerPoint.ShapeType.geometricShape ||
        s.type === PowerPoint.ShapeType.line,
  );

  const reference = compatible.find((s) => s.id === referenceId);
  if (!reference) {
    throw new Error(
      `The reference shape has no ${wantsFill ? "fill" : "outline"} to match.`,
    );
  }

  compatible.forEach((s) => {
    if (wantsFill) s.fill.load("foregroundColor");
    else s.lineFormat.load("color,weight");
  });
  await context.sync();

  return compatible
    .filter((s) => {
      if (property === "fillColor") {
        return s.fill.foregroundColor === reference.fill.foregroundColor;
      }
      if (property === "lineColor") {
        return s.lineFormat.color === reference.lineFormat.color;
      }
      return near(s.lineFormat.weight, reference.lineFormat.weight, WEIGHT_TOLERANCE);
    })
    .map((s) => s.id);
}

/** Font-name matching — second pass over the text-capable subset. */
async function selectSameFont(
  context: PowerPoint.RequestContext,
  all: PowerPoint.Shape[],
  referenceId: string,
): Promise<string[]> {
  const textShapes = all.filter((s) => isTextShape(s.type));
  if (!textShapes.some((s) => s.id === referenceId)) {
    throw new Error("The reference shape has no text to match.");
  }

  const fonts = new Map<string, PowerPoint.ShapeFont>();
  textShapes.forEach((s) => {
    const font = s.textFrame.textRange.font;
    font.load("name");
    fonts.set(s.id, font);
  });
  await context.sync();

  const referenceName = fonts.get(referenceId)?.name;
  return textShapes
    .filter((s) => fonts.get(s.id)?.name === referenceName)
    .map((s) => s.id);
}
