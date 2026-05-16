/**
 * Shape creation operations — Phase 3.
 *
 * New shapes are added to the slide currently in view, at a default
 * position; the user then moves them as needed.
 */

/** Default placement for a newly inserted shape, points. */
const DEFAULT_PLACEMENT = { left: 120, top: 120, width: 200, height: 120 };

async function getActiveSlide(
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
