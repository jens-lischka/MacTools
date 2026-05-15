import type { ComponentType } from "react";

/** Feasibility rating for rebuilding a legacy feature on Office.js. */
export type Rating = "ok" | "partial" | "blocked";

export interface Feature {
  /** Stable identifier, reused from the legacy ribbon control id where sensible. */
  id: string;
  label: string;
  /** Short explanation shown as a tooltip / secondary text. */
  description?: string;
  rating: Rating;
  /**
   * Action handler. Undefined means "not yet implemented" — the button is
   * shown but disabled. Blocked features never have a handler.
   */
  run?: () => Promise<void>;
}

export interface Category {
  id: string;
  label: string;
  /** Generic feature-button list. Ignored when `panel` is set. */
  features?: Feature[];
  /** Optional custom UI for categories that need more than buttons. */
  panel?: ComponentType;
}
