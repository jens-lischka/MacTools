/**
 * Table operations — require PowerPoint API 1.8.
 *
 * The panel gates these behind a capability check, so they are only reachable
 * on clients that support the table API.
 */

import { getActiveSlide } from "./powerpoint";

/** Find the selected table, or throw a friendly error. */
async function getSelectedTable(
  context: PowerPoint.RequestContext,
): Promise<{ table: PowerPoint.Table; shape: PowerPoint.Shape }> {
  const selected = context.presentation.getSelectedShapes();
  selected.load("items/type,items/left,items/top,items/width,items/height");
  await context.sync();

  const shape = selected.items.find((s) => s.type === PowerPoint.ShapeType.table);
  if (!shape) throw new Error("Select a table first.");
  return { table: shape.getTable(), shape };
}

/** Insert a new table on the slide in view. */
export async function insertTable(rows: number, columns: number): Promise<void> {
  if (!(rows > 0) || !(columns > 0)) {
    throw new Error("Rows and columns must be greater than 0.");
  }
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    slide.shapes.addTable(rows, columns, {
      left: 60,
      top: 90,
      width: 600,
      height: 60 + rows * 30,
    });
    await context.sync();
  });
}

/** Append a row to the selected table. */
export async function addTableRow(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const { table } = await getSelectedTable(context);
    table.load("rowCount");
    await context.sync();
    table.rows.add(table.rowCount, 1);
    await context.sync();
  });
}

/** Append a column to the selected table. */
export async function addTableColumn(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const { table } = await getSelectedTable(context);
    table.load("columnCount");
    await context.sync();
    table.columns.add(table.columnCount, 1);
    await context.sync();
  });
}

/** Delete the last row of the selected table. */
export async function deleteLastTableRow(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const { table } = await getSelectedTable(context);
    table.load("rowCount");
    await context.sync();
    if (table.rowCount <= 1) throw new Error("A table must keep at least one row.");
    table.rows.deleteRows([table.rows.getItemAt(table.rowCount - 1)]);
    await context.sync();
  });
}

/** Delete the last column of the selected table. */
export async function deleteLastTableColumn(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const { table } = await getSelectedTable(context);
    table.load("columnCount");
    await context.sync();
    if (table.columnCount <= 1) {
      throw new Error("A table must keep at least one column.");
    }
    table.columns.deleteColumns([table.columns.getItemAt(table.columnCount - 1)]);
    await context.sync();
  });
}

/** Convert the selected table's contents into a tab-separated text box. */
export async function tableToText(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    const { table, shape } = await getSelectedTable(context);
    table.load("values");
    await context.sync();

    const text = table.values.map((row) => row.join("\t")).join("\n");
    slide.shapes.addTextBox(text, {
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
    });
    await context.sync();
  });
}
