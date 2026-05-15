import { createContext, useContext } from "react";

/** Shared runner so any feature button can execute through the App's
 *  busy-state + error-reporting plumbing. */
export interface ActionRunner {
  run: (label: string, fn: () => Promise<void>) => void;
  busy: boolean;
}

export const ActionContext = createContext<ActionRunner>({
  run: () => undefined,
  busy: false,
});

export const useActions = (): ActionRunner => useContext(ActionContext);
