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
   * PowerPoint API requirement set this feature needs (e.g. "1.8"). When the
   * client does not support it, the feature renders disabled regardless of
   * its rating.
   */
  requiresApi?: string;
  /**
   * Action handler. Undefined means "not yet implemented" — the button is
   * shown but disabled. Blocked features never have a handler.
   */
  run?: () => Promise<void>;
}

export interface Category {
  id: string;
  label: string;
  /**
   * Generic feature-button list. May be combined with `panel` — when both are
   * present the panel renders first, then the feature list (used to keep
   * not-yet-built / blocked features visible alongside implemented ones).
   */
  features?: Feature[];
  /** Optional custom UI for categories that need more than buttons. */
  panel?: ComponentType;
}
