
const myColumns = ["cool", "age", "city", "country","note"]

Office.onReady(async () => {
  document.getElementById("saveBtn").onclick = saveChanges;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onSelectionChanged.add(onSelectionChanged);
    await context.sync();
  });

  await loadSelectedRow();
});

async function onSelectionChanged(event) {
  try {
    await loadSelectedRow(); // Refresh the task pane with the new selection
  } catch (error) {
    console.error("Error updating task pane:", error);
  }
}
/* 
async function loadSelectedRow() {
  await Excel.run(async (context) => {
    const table = context.workbook.getTables()[0];
    const activeCell = context.workbook.getActiveCell();
    const tableRange = table.getRangeBetweenHeaderAndTotal();
    const rowCount = table.getRowCount();
    let tableData = tableRange.getValues();
    let activeRowIndex;
    await context.sync();
    for (let i = 0; i < rowCount; i++) {
      const rowRange = tableRange.getRow(i);
      if (rowRange.getAddress() === activeCell.getEntireRow().getAddress()) {
        console.log(`Active cell is in table row index: ${i}`);
        activeRowIndex = i;
      }
    }
    const container = document.getElementById("fields");
    container.innerHTML = tableData[activeRowIndex]
      .slice(0, 4)
      .map(
        (val, i) =>
          `<input type="text" id="col${i}" myindex="${activeRowIndex}" value="${val}" />`
      )
      .join("");
    await context.sync();
  });
} */

  async function loadSelectedRow() {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    const activeCell = wb.getActiveCell();

    // ✅ Use the TableCollection on the workbook
    const tables = wb.tables;

    // Optionally check there is at least one table
    const count = tables.getCount();       // ClientResult<number>
    await context.sync();
    if (count.value === 0) {
      document.getElementById("fields").innerHTML = "<em>No tables in this workbook.</em>";
      return;
    }

    // Use the first table (or use `getItem("YourTableName")` if you know it)
    const table = tables.getItemAt(0);
   
    // Work only with the data body (no headers/totals)
    const body = table.getDataBodyRange();
    const hit  = body.getIntersectionOrNullObject(activeCell.getEntireRow());

    // Load what we need
    body.load(["rowIndex", "columnCount"]);
    hit.load(["isNullObject", "rowIndex"]);
    await context.sync();

    // If the selection isn't in the table body, tell the user
    if (hit.isNullObject) {
      document.getElementById("fields").innerHTML = "<em>Select a row inside the table body.</em>";
      return;
    }

    // Compute the row index relative to the table body
    const relativeRowIndex = hit.rowIndex - body.rowIndex;
    const cols = table.columns;
    const colwithValues = []
    for (let i = 0; i < myColumns.length; i++) {
      const colName = myColumns[i];
      const colValue = cols.getItem(colName).getDataBodyRange().getCell(relativeRowIndex, 0);
      colValue.load("values");
      colwithValues.push(colValue);
    }

    /* const city = cols.getItem("city").getDataBodyRange().getCell(relativeRowIndex, 0);
    const note = cols.getItem("note").getDataBodyRange().getCell(relativeRowIndex, 0);
    city.load("values");
    note.load("values");
    console.log("city", city, note); */
    await context.sync();

    const values = [...colwithValues.map(c=>c.values[0])];
    const container = document.getElementById("fields");
    console.log("mycolumns", myColumns)
    container.innerHTML = values
      .map((val, i) =>
        
        `<div><label for="col${i}">${myColumns[i]}</label>
        <input type="text" id="col${i}" data-rowindex="${relativeRowIndex}" value="${String(val[0] ?? "")}" /></div>`
      )
      .join("");
  });
}

/* 
async function loadSelectedRow() {
  await Excel.run(async (context) => {
    const table = context.workbook.getTables()[0];
    const activeCell = context.workbook.getActiveCell();
    const tableRange = table.getRangeBetweenHeaderAndTotal();
    const rowCount = table.getRowCount();
    let tableData = tableRange.getValues();
    let activeRowIndex;
    await context.sync();
    for (let i = 0; i < rowCount; i++) {
      const rowRange = tableRange.getRow(i);
      if (rowRange.getAddress() === activeCell.getEntireRow().getAddress()) {
        console.log(`Active cell is in table row index: ${i}`);
        activeRowIndex = i;
      }
    }
    const container = document.getElementById("fields");
    container.innerHTML = tableData[activeRowIndex]
      .slice(0, 4)
      .map(
        (val, i) =>
          `<input type="text" id="col${i}" myindex="${activeRowIndex}" value="${val}" />`
      )
      .join("");
    await context.sync();
  });
} */

/*  async function loadSelectedRow() {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    const activeCell = wb.getActiveCell();

    // ✅ Use the TableCollection on the workbook
    const tables = wb.tables;

    // Optionally check there is at least one table
    const count = tables.getCount();       // ClientResult<number>
    await context.sync();
    if (count.value === 0) {
      document.getElementById("fields").innerHTML = "<em>No tables in this workbook.</em>";
      return;
    }

    // Use the first table (or use `getItem("YourTableName")` if you know it)
    const table = tables.getItemAt(0);

    // Work only with the data body (no headers/totals)
    const body = table.getDataBodyRange();
    const hit  = body.getIntersectionOrNullObject(activeCell.getEntireRow());

    // Load what we need
    body.load(["rowIndex", "columnCount"]);
    hit.load(["isNullObject", "rowIndex"]);
    await context.sync();

    // If the selection isn't in the table body, tell the user
    if (hit.isNullObject) {
      document.getElementById("fields").innerHTML = "<em>Select a row inside the table body.</em>";
      return;
    }

    // Compute the row index relative to the table body
    const relativeRowIndex = hit.rowIndex - body.rowIndex;
    console.log("relativerow", relativeRowIndex)
    console.log("colcount", body.columnCount)
    // First 4 columns (cap at body width)
    const width  = Math.min(4, body.columnCount);
    console.log("width", width)
    const first4 = body.getRow(relativeRowIndex);
    console.log("first4", first4);
    first4.load("values");
    await context.sync();

    const values = first4.values[0].slice(0,4); // [[...]] -> first row
    const container = document.getElementById("fields");
    container.innerHTML = values
      .map((val, i) =>
        `<input type="text" id="col${i}" data-rowindex="${relativeRowIndex}" value="${String(val ?? "")}" />`
      )
      .join("");
  });
}
*/
async function saveChanges() {
  await Excel.run(async (context) => {
    const wb    = context.workbook;
    const table = wb.tables.getItemAt(0);
    const body  = table.getDataBodyRange();

    const rowIndex = parseInt(document.getElementById("col0").dataset.rowindex, 10);
    if (isNaN(rowIndex)) return;

    const newValues = [
      document.getElementById("col0").value,
      document.getElementById("col1").value,
      document.getElementById("col2").value,
      document.getElementById("col3").value,
    ];

    // Target the specific row and the first 4 columns, then write values
    const target = body.getRow(rowIndex).getResizedRange(0, -1);
    target.load("values");
    await context.sync();
    console.log("targetvalues", target.values)
    
    target.values = [newValues]
    await context.sync();

    await loadSelectedRow();
  });
}
