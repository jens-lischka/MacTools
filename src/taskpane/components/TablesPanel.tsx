import * as React from "react";
import {
  Button,
  Field,
  SpinButton,
  Divider,
  Caption1,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  insertTable,
  addTableRow,
  addTableColumn,
  deleteLastTableRow,
  deleteLastTableColumn,
  tableToText,
} from "../../lib/tables";
import { isApiSupported } from "../../lib/capabilities";
import { useActions } from "./ActionContext";
import { ToolIcon } from "./ToolIcon";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  row: { display: "flex", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
  grow: { flexGrow: 1 },
  grid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: tokens.spacingHorizontalS,
  },
});

export const TablesPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const [rows, setRows] = React.useState(3);
  const [columns, setColumns] = React.useState(3);

  if (!isApiSupported("1.8")) {
    return (
      <MessageBar intent="warning">
        <MessageBarBody>
          Table tools need PowerPoint API 1.8, which this client does not
          support. They will activate automatically on a newer build.
        </MessageBarBody>
      </MessageBar>
    );
  }

  const spin =
    (set: (n: number) => void): React.ComponentProps<typeof SpinButton>["onChange"] =>
    (_, d) => {
      const next = d.value ?? Number(d.displayValue);
      if (Number.isFinite(next)) set(next as number);
    };

  return (
    <div className={styles.section}>
      <Field label="Insert table — rows">
        <SpinButton min={1} max={50} value={rows} onChange={spin(setRows)} />
      </Field>
      <Field label="Insert table — columns">
        <div className={styles.row}>
          <SpinButton
            className={styles.grow}
            min={1}
            max={50}
            value={columns}
            onChange={spin(setColumns)}
          />
          <Button
            disabled={busy}
            icon={<ToolIcon name="FormatTable" />}
            onClick={() =>
              run("Insert table", () => insertTable(rows, columns))
            }
          >
            Insert
          </Button>
        </div>
      </Field>

      <Divider />

      <Caption1>The actions below apply to the selected table.</Caption1>
      <div className={styles.grid}>
        <Button disabled={busy} icon={<ToolIcon name="AddRowBottom" />} onClick={() => run("Add row", () => addTableRow())}>
          Add Row
        </Button>
        <Button disabled={busy} icon={<ToolIcon name="AddColumnRight" />} onClick={() => run("Add column", () => addTableColumn())}>
          Add Column
        </Button>
        <Button
          disabled={busy}
          icon={<ToolIcon name="RemoveLastRow" />}
          onClick={() => run("Delete last row", () => deleteLastTableRow())}
        >
          Delete Last Row
        </Button>
        <Button
          disabled={busy}
          icon={<ToolIcon name="RemoveLastColumn" />}
          onClick={() => run("Delete last column", () => deleteLastTableColumn())}
        >
          Delete Last Column
        </Button>
      </div>
      <Button disabled={busy} icon={<ToolIcon name="TableToText" />} onClick={() => run("Table to text", () => tableToText())}>
        Convert Table to Text
      </Button>
    </div>
  );
};
