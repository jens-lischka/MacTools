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
}

export class NoSelectionError extends Error {
  constructor() {
    super("Select two or more shapes first.");
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
    collection.load("items/id,items/left,items/top,items/width,items/height");
    await context.sync();

    const shapes = collection.items;
    if (shapes.length < minShapes) throw new NoSelectionError();

    const geometry: SelectedShape[] = shapes.map((s) => ({
      id: s.id,
      left: s.left,
      top: s.top,
      width: s.width,
      height: s.height,
    }));

    callback(shapes, geometry, context);
    await context.sync();
  });
}

/** Bounding box that encloses every box in the list. */
export function unionBox(boxes: Box[]): Box {
  const left = Math.min(...boxes.map((b) => b.left));
  const top = Math.min(...boxes.map((b) => b.top));
  const right = Math.max(...boxes.map((b) => b.left + b.width));
  const bottom = Math.max(...boxes.map((b) => b.top + b.height));
  return { left, top, width: right - left, height: bottom - top };
}
