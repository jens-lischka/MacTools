/**
 * Shape visibility operations — require PowerPoint API 1.10 (`Shape.visible`).
 */

/** Hide or show every selected shape. */
export async function setSelectedShapesVisible(visible: boolean): Promise<void> {
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id");
    await context.sync();
    if (selected.items.length === 0) {
      throw new Error("Select at least one shape first.");
    }
    selected.items.forEach((s) => {
      s.visible = visible;
    });
    await context.sync();
  });
}

/** Make every shape on the slide in view visible again. */
export async function showAllShapes(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items/id");
    await context.sync();
    if (slides.items.length === 0) throw new Error("Open a slide first.");

    const shapes = slides.items[0].shapes;
    shapes.load("items/id");
    await context.sync();
    shapes.items.forEach((s) => {
      s.visible = true;
    });
    await context.sync();
  });
}
