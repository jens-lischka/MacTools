/**
 * Shape creation operations — Phase 3.
 *
 * New shapes are added to the slide currently in view, at a default
 * position; the user then moves them as needed.
 */

import { withSelectedShapes } from "./powerpoint";

/** Default placement for a newly inserted shape, points. */
const DEFAULT_PLACEMENT = { left: 120, top: 120, width: 200, height: 120 };

export async function getActiveSlide(
  context: PowerPoint.RequestContext,
): Promise<PowerPoint.Slide> {
  const slides = context.presentation.getSelectedSlides();
  slides.load("items/id");
  await context.sync();
  if (slides.items.length === 0) {
    throw new Error("Open a slide before inserting a shape.");
  }
  return slides.items[0];
}

/** Insert a geometric shape (rectangle, oval, …) on the active slide. */
export async function insertShape(
  shapeType: PowerPoint.GeometricShapeType,
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    slide.shapes.addGeometricShape(shapeType, { ...DEFAULT_PLACEMENT });
    await context.sync();
  });
}

/** Insert a straight line on the active slide. */
export async function insertLine(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
      left: DEFAULT_PLACEMENT.left,
      top: DEFAULT_PLACEMENT.top,
      width: DEFAULT_PLACEMENT.width,
      height: 0,
    });
    await context.sync();
  });
}

/** Insert a text box on the active slide. */
export async function insertTextBox(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    slide.shapes.addTextBox("Text", { ...DEFAULT_PLACEMENT, height: 60 });
    await context.sync();
  });
}

/** Connect the centres of exactly two selected shapes with a straight line. */
export async function insertConnector(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/left,items/top,items/width,items/height");
    await context.sync();
    if (selected.items.length !== 2) {
      throw new Error("Select exactly two shapes to connect.");
    }
    const centre = (s: PowerPoint.Shape) => ({
      x: s.left + s.width / 2,
      y: s.top + s.height / 2,
    });
    const a = centre(selected.items[0]);
    const b = centre(selected.items[1]);
    slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
      left: Math.min(a.x, b.x),
      top: Math.min(a.y, b.y),
      width: Math.abs(b.x - a.x),
      height: Math.abs(b.y - a.y),
    });
    await context.sync();
  });
}

/** Insert a numbered circle (an ellipse with a centred label). */
export async function insertNumberedCircle(label: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    const circle = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.ellipse,
      { left: 120, top: 120, width: 48, height: 48 },
    );
    circle.textFrame.textRange.text = label.trim() || "1";
    circle.textFrame.textRange.paragraphFormat.horizontalAlignment =
      PowerPoint.ParagraphHorizontalAlignment.center;
    await context.sync();
  });
}

export type GroupDirection = "row" | "column";

/** Lay the selected shapes out edge-to-edge as a row or a column. */
export async function groupAsLayout(direction: GroupDirection): Promise<void> {
  const gap = 8;
  await withSelectedShapes(2, (shapes, geometry) => {
    const order = geometry
      .map((g, i) => ({ g, i }))
      .sort((a, b) =>
        direction === "row" ? a.g.left - b.g.left : a.g.top - b.g.top,
      );
    const base = direction === "row" ? order[0].g.top : order[0].g.left;
    let cursor = direction === "row" ? order[0].g.left : order[0].g.top;

    order.forEach((o) => {
      const shape = shapes[o.i];
      if (direction === "row") {
        shape.left = cursor;
        shape.top = base;
        cursor += o.g.width + gap;
      } else {
        shape.top = cursor;
        shape.left = base;
        cursor += o.g.height + gap;
      }
    });
  });
}
