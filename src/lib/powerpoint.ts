/**
 * Thin helpers over the PowerPoint JavaScript API.
 *
 * All geometry is in points (the unit the API uses for shape left/top/
 * width/height). The legacy VSTO add-in did the same maths against the COM
 * model; here we do it against the JS model so it runs on web, Mac and Win.
 */

/** Axis-aligned bounding box, points. */
export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface SelectedShape extends Box {
  id: string;
  type: PowerPoint.Shape["type"];
}

export class NoSelectionError extends Error {
  constructor(minShapes: number) {
    super(
      minShapes <= 1
        ? "Select at least one shape first."
        : `Select ${minShapes} or more shapes first.`,
    );
    this.name = "NoSelectionError";
  }
}

/**
 * Run a callback against the currently selected shapes, with geometry loaded.
 * Throws NoSelectionError when fewer than `minShapes` shapes are selected.
 */
export async function withSelectedShapes(
  minShapes: number,
  callback: (
    shapes: PowerPoint.Shape[],
    geometry: SelectedShape[],
    context: PowerPoint.RequestContext,
  ) => void,
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const collection = context.presentation.getSelectedShapes();
    collection.load("items/id,items/type,items/left,items/top,items/width,items/height");
    await context.sync();

    const shapes = collection.items;
    if (shapes.length < minShapes) throw new NoSelectionError(minShapes);

    const geometry: SelectedShape[] = shapes.map((s) => ({
      id: s.id,
      type: s.type,
      left: s.left,
      top: s.top,
      width: s.width,
      height: s.height,
    }));

    callback(shapes, geometry, context);
    await context.sync();

    // Programmatic geometry changes update the model and the slide thumbnail
    // but do not always repaint the live editing canvas. Re-selecting the
    // shapes nudges PowerPoint to refresh the view.
    const slide = shapes[0].getParentSlide();
    slide.setSelectedShapes(geometry.map((g) => g.id));
    await context.sync();
  });
}

/**
 * True for shape types that carry a text frame. A function (not a
 * module-level constant) so the `PowerPoint` enum is read lazily — referencing
 * it at module load can run before Office.js has defined the namespace.
 */
export function isTextShape(type: PowerPoint.Shape["type"]): boolean {
  return (
    type === PowerPoint.ShapeType.geometricShape ||
    type === PowerPoint.ShapeType.textBox ||
    type === PowerPoint.ShapeType.placeholder
  );
}

/** The first selected slide, or throw if no slide is in view. */
export async function getActiveSlide(
  context: PowerPoint.RequestContext,
): Promise<PowerPoint.Slide> {
  const slides = context.presentation.getSelectedSlides();
  slides.load("items/id");
  await context.sync();
  if (slides.items.length === 0) throw new Error("Open a slide first.");
  return slides.items[0];
}

/** Bounding box that encloses every box in the list. */
export function unionBox(boxes: Box[]): Box {
  const left = Math.min(...boxes.map((b) => b.left));
  const top = Math.min(...boxes.map((b) => b.top));
  const right = Math.max(...boxes.map((b) => b.left + b.width));
  const bottom = Math.max(...boxes.map((b) => b.top + b.height));
  return { left, top, width: right - left, height: bottom - top };
}
