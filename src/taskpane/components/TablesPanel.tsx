import * as React from "react";
import {
  Field,
  SpinButton,
  Divider,
  Caption1,
  MessageBar,
  MessageBarBody,
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
import { ToolButton } from "./ToolButton";
import { usePanelStyles } from "./panelStyles";
import { spinHandler } from "./spin";

export const TablesPanel: React.FC = () => {
  const styles = usePanelStyles();
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

  return (
    <div className={styles.section}>
      <Field label="Insert table — rows &amp; columns">
        <div className={styles.toolbar}>
          <SpinButton
            className={styles.spin}
            min={1}
            max={50}
            value={rows}
            onChange={spinHandler(setRows)}
          />
          <SpinButton
            className={styles.spin}
            min={1}
            max={50}
            value={columns}
            onChange={spinHandler(setColumns)}
          />
          <ToolButton
            icon="FormatTable"
            label="Insert table"
            disabled={busy}
            onClick={() => run("Insert table", () => insertTable(rows, columns))}
          />
        </div>
      </Field>

      <Divider />

      <Caption1>The actions below apply to the selected table.</Caption1>
      <div className={styles.toolbar}>
        <ToolButton
          icon="AddRowBottom"
          label="Add row"
          disabled={busy}
          onClick={() => run("Add row", () => addTableRow())}
        />
        <ToolButton
          icon="AddColumnRight"
          label="Add column"
          disabled={busy}
          onClick={() => run("Add column", () => addTableColumn())}
        />
        <ToolButton
          icon="RemoveLastRow"
          label="Delete last row"
          disabled={busy}
          onClick={() => run("Delete last row", () => deleteLastTableRow())}
        />
        <ToolButton
          icon="RemoveLastColumn"
          label="Delete last column"
          disabled={busy}
          onClick={() => run("Delete last column", () => deleteLastTableColumn())}
        />
        <ToolButton
          icon="TableToText"
          label="Convert table to text"
          disabled={busy}
          onClick={() => run("Table to text", () => tableToText())}
        />
      </div>
    </div>
  );
};
