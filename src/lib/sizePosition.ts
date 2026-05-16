/**
 * Size & Position operations — Phase 2.
 *
 * All pure geometry against the selected shapes, so it runs identically on
 * web, Mac and Windows. Units are points.
 */

import { withSelectedShapes, unionBox, type Box } from "./powerpoint";
import { getSetting } from "./settings";

export type Reference = "first" | "last";
export type MatchDimension = "width" | "height" | "both";
export type StretchEdge = "left" | "right" | "top" | "bottom";
export type StretchTarget = "selection" | "slide";

interface SlideSize {
  width: number;
  height: number;
}

const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };

/** Smallest dimension a stretch is allowed to produce, points. */
const MIN_SIZE = 1;

/** Resize every selected shape to match a reference shape's dimensions. */
export async function matchSize(
  dimension: MatchDimension,
  reference: Reference,
): Promise<void> {
  await withSelectedShapes(2, (shapes, geometry) => {
    const ref = reference === "first" ? geometry[0] : geometry[geometry.length - 1];
    shapes.forEach((shape, i) => {
      if (geometry[i].id === ref.id) return;
      if (dimension !== "height") shape.width = ref.width;
      if (dimension !== "width") shape.height = ref.height;
    });
  });
}

/** Scale every selected shape about its own centre by `percent`. */
export async function scaleShapes(percent: number): Promise<void> {
  if (!(percent > 0)) throw new Error("Enter a scale percentage greater than 0.");
  const factor = percent / 100;
  await withSelectedShapes(1, (shapes, geometry) => {
    shapes.forEach((shape, i) => {
      const g = geometry[i];
      const newWidth = g.width * factor;
      const newHeight = g.height * factor;
      shape.left = g.left + (g.width - newWidth) / 2;
      shape.top = g.top + (g.height - newHeight) / 2;
      shape.width = newWidth;
      shape.height = newHeight;
    });
  });
}

/** Extend one edge of every selected shape out to the target box's edge. */
export async function stretchToEdge(
  edge: StretchEdge,
  target: StretchTarget,
): Promise<void> {
  const slideSize = await getSetting<SlideSize>("slideSize", DEFAULT_SLIDE_SIZE);
  const minShapes = target === "slide" ? 1 : 2;

  await withSelectedShapes(minShapes, (shapes, geometry) => {
    const box: Box =
      target === "slide"
        ? { left: 0, top: 0, width: slideSize.width, height: slideSize.height }
        : unionBox(geometry);

    shapes.forEach((shape, i) => {
      const g = geometry[i];
      switch (edge) {
        case "left": {
          const right = g.left + g.width;
          shape.left = box.left;
          shape.width = Math.max(MIN_SIZE, right - box.left);
          break;
        }
        case "right":
          shape.width = Math.max(MIN_SIZE, box.left + box.width - g.left);
          break;
        case "top": {
          const bottom = g.top + g.height;
          shape.top = box.top;
          shape.height = Math.max(MIN_SIZE, bottom - box.top);
          break;
        }
        case "bottom":
          shape.height = Math.max(MIN_SIZE, box.top + box.height - g.top);
          break;
      }
    });
  });
}

/**
 * Straighten selected line shapes to be perfectly horizontal or vertical,
 * whichever they are already closest to. Non-line shapes are ignored.
 */
export async function straightenLines(): Promise<void> {
  await withSelectedShapes(1, (shapes, geometry) => {
    const lines = geometry.filter((g) => g.type === PowerPoint.ShapeType.line);
    if (lines.length === 0) throw new Error("Select at least one line.");

    shapes.forEach((shape, i) => {
      const g = geometry[i];
      if (g.type !== PowerPoint.ShapeType.line) return;
      if (g.width >= g.height) {
        shape.top = g.top + g.height / 2;
        shape.height = 0;
      } else {
        shape.left = g.left + g.width / 2;
        shape.width = 0;
      }
    });
  });
}

/** Close the gaps between selected shapes by extending each to its neighbour. */
export async function fillGap(axis: "horizontal" | "vertical"): Promise<void> {
  await withSelectedShapes(2, (shapes, geometry) => {
    const order = geometry
      .map((g, i) => ({ g, i }))
      .sort((a, b) =>
        axis === "horizontal" ? a.g.left - b.g.left : a.g.top - b.g.top,
      );
    for (let k = 0; k < order.length - 1; k++) {
      const current = order[k];
      const next = order[k + 1];
      const shape = shapes[current.i];
      if (axis === "horizontal") {
        shape.width = Math.max(MIN_SIZE, next.g.left - current.g.left);
      } else {
        shape.height = Math.max(MIN_SIZE, next.g.top - current.g.top);
      }
    }
  });
}

/** Resize every selected shape to match the largest one (by area). */
export async function unifyShapes(): Promise<void> {
  await withSelectedShapes(2, (shapes, geometry) => {
    const largest = geometry.reduce((a, b) =>
      a.width * a.height >= b.width * b.height ? a : b,
    );
    shapes.forEach((shape) => {
      shape.width = largest.width;
      shape.height = largest.height;
    });
  });
}
