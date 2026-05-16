/**
 * Swap / Pick up & Apply operations — Phase 3.
 *
 * "Pick up" stores a shape's size and position in roaming settings, so it
 * survives across documents and sessions until overwritten.
 */

import { withSelectedShapes, type Box } from "./powerpoint";
import { getSetting, setSetting } from "./settings";

const PICKUP_KEY = "pickup:sizePosition";

/** Swap the positions of exactly two selected shapes. */
export async function swapPosition(): Promise<void> {
  await withSelectedShapes(2, (shapes, geometry) => {
    if (geometry.length !== 2) {
      throw new Error("Select exactly two shapes to swap.");
    }
    const [a, b] = geometry;
    shapes[0].left = b.left;
    shapes[0].top = b.top;
    shapes[1].left = a.left;
    shapes[1].top = a.top;
  });
}

/** Swap the sizes of exactly two selected shapes. */
export async function swapSize(): Promise<void> {
  await withSelectedShapes(2, (shapes, geometry) => {
    if (geometry.length !== 2) {
      throw new Error("Select exactly two shapes to swap.");
    }
    const [a, b] = geometry;
    shapes[0].width = b.width;
    shapes[0].height = b.height;
    shapes[1].width = a.width;
    shapes[1].height = a.height;
  });
}

/** Swap fill colour and outline (colour, weight, dash) between two shapes. */
export async function swapFillAndOutline(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id,items/type");
    await context.sync();

    if (selected.items.length !== 2) {
      throw new Error("Select exactly two shapes to swap.");
    }
    const [a, b] = selected.items;
    if (
      a.type !== PowerPoint.ShapeType.geometricShape ||
      b.type !== PowerPoint.ShapeType.geometricShape
    ) {
      throw new Error("Both shapes must be standard shapes with a fill.");
    }

    [a, b].forEach((s) => {
      s.fill.load("foregroundColor");
      s.lineFormat.load("color,weight,dashStyle");
    });
    await context.sync();

    const snapshot = (s: PowerPoint.Shape) => ({
      fill: s.fill.foregroundColor,
      color: s.lineFormat.color,
      weight: s.lineFormat.weight,
      dashStyle: s.lineFormat.dashStyle,
    });
    const aStyle = snapshot(a);
    const bStyle = snapshot(b);

    const apply = (s: PowerPoint.Shape, style: ReturnType<typeof snapshot>) => {
      s.fill.setSolidColor(style.fill);
      s.lineFormat.color = style.color;
      s.lineFormat.weight = style.weight;
      s.lineFormat.dashStyle = style.dashStyle;
    };
    apply(a, bStyle);
    apply(b, aStyle);
    await context.sync();
  });
}

/** Store the size and position of the single selected shape. */
export async function pickUpSizePosition(): Promise<void> {
  let picked: Box | null = null;
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/left,items/top,items/width,items/height");
    await context.sync();
    if (selected.items.length !== 1) {
      throw new Error("Select exactly one shape to pick up from.");
    }
    const s = selected.items[0];
    picked = { left: s.left, top: s.top, width: s.width, height: s.height };
  });
  if (picked) await setSetting(PICKUP_KEY, picked, "roaming");
}

/** Apply the picked-up size and position to every selected shape. */
export async function applySizePosition(): Promise<void> {
  const picked = await getSetting<Box | null>(PICKUP_KEY, null);
  if (!picked) {
    throw new Error("Pick up size & position from a shape first.");
  }
  await withSelectedShapes(1, (shapes) => {
    shapes.forEach((s) => {
      s.left = picked.left;
      s.top = picked.top;
      s.width = picked.width;
      s.height = picked.height;
    });
  });
}
