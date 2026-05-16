import { makeStyles, tokens } from "@fluentui/react-components";

/** Layout styles shared by every category panel. */
export const usePanelStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  toolbar: {
    display: "flex",
    flexWrap: "wrap",
    gap: tokens.spacingHorizontalS,
    alignItems: "center",
  },
  grow: { flexGrow: 1, minWidth: "140px" },
  spin: { width: "96px" },
});
