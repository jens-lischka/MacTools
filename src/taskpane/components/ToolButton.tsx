import * as React from "react";
import { Button, Tooltip } from "@fluentui/react-components";
import { ToolIcon } from "./ToolIcon";

interface ToolButtonProps {
  icon: string;
  /** Shown as a tooltip and used as the accessible name (buttons are icon-only). */
  label: string;
  onClick: () => void;
  disabled?: boolean;
}

/** An icon-only action button with the label exposed via tooltip + aria-label. */
export const ToolButton: React.FC<ToolButtonProps> = ({
  icon,
  label,
  onClick,
  disabled,
}) => (
  <Tooltip content={label} relationship="label">
    <Button
      icon={<ToolIcon name={icon} />}
      disabled={disabled}
      aria-label={label}
      onClick={onClick}
    />
  </Tooltip>
);
