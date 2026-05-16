import * as React from "react";
import { Button, Tooltip } from "@fluentui/react-components";
import { ToolIcon } from "./ToolIcon";

/** Info button in the header — opens the online manual in a new browser tab. */
export const HelpButton: React.FC = () => (
  <Tooltip content="Open the manual" relationship="label">
    <Button
      appearance="subtle"
      icon={<ToolIcon name="Info" />}
      aria-label="Open the manual"
      onClick={() => window.open("manual.html", "_blank", "noopener")}
    />
  </Tooltip>
);
