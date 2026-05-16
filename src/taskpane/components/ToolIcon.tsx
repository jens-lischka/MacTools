import * as React from "react";

/** Renders one of the bundled PNG tool icons from assets/icons/. */
export const ToolIcon: React.FC<{ name: string }> = ({ name }) => (
  <img
    src={`assets/icons/${name}.png`}
    alt=""
    width={20}
    height={20}
    style={{ display: "block" }}
  />
);
