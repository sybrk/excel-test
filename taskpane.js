const DEFAULT_COLUMNS = [""];
const LS_KEY_COLUMNS = "tableHelper.columns";

function parseColumnsInput(text) {
  // Allow comma- or newline-separated entries
  const parts = (text || "")
    .split(/[\n,]/g)
    .map((s) => s.trim())
    .filter(Boolean);
  // Deduplicate (case sensitive by default; switch to .toLowerCase() if you want case-insensitive)
  const unique = [...new Set(parts)];
  return unique;
}

function loadUserColumns() {
  try {
    const raw = localStorage.getItem(LS_KEY_COLUMNS);
    if (!raw) return [...DEFAULT_COLUMNS];
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && parsed.every((x) => typeof x === "string")) {
      return parsed;
    }
  } catch (e) {
    console.warn(
      "Failed to read columns from localStorage, using defaults.",
      e
    );
  }
  return [...DEFAULT_COLUMNS];
}

function saveUserColumns(columns) {
  try {
    localStorage.setItem(LS_KEY_COLUMNS, JSON.stringify(columns));
    return true;
  } catch (e) {
    console.error("Failed to save columns to localStorage", e);
    return false;
  }
}

function resetUserColumns() {
  try {
    localStorage.removeItem(LS_KEY_COLUMNS);
  } catch {}
}

Office.onReady(async () => {
  document.getElementById("saveBtn").onclick = saveChanges;

  // Settings UI
  document.getElementById("toggleSettingsBtn").addEventListener("click", () => {
    const sec = document.getElementById("settings");
    sec.hidden = !sec.hidden;
    if (!sec.hidden) {
      // Populate textarea with current setting
      const cols = loadUserColumns();
      document.getElementById("columnsInput").value = cols.join(", ");
      document.getElementById("settingsMessage").textContent = "";
    }
  });

  document
    .getElementById("saveSettingsBtn")
    .addEventListener("click", async () => {
      const txt = document.getElementById("columnsInput").value;
      const cols = parseColumnsInput(txt);

      if (cols.length === 0) {
        document.getElementById("settingsMessage").textContent =
          "Please enter at least one column name.";
        return;
      }

      // Optional: cap the number of columns to keep the UI snappy
      if (cols.length > 20) {
        document.getElementById("settingsMessage").textContent =
          "Please keep it to 20 columns or fewer.";
        return;
      }

      const ok = saveUserColumns(cols);
      document.getElementById("settingsMessage").textContent = ok
        ? "Saved."
        : "Could not save settings.";
      // Re-render the editor for the currently selected row with the new columns
      await loadSelectedRow();
    });

  document
    .getElementById("resetSettingsBtn")
    .addEventListener("click", async () => {
      resetUserColumns();
      document.getElementById("columnsInput").value =
        DEFAULT_COLUMNS.join(", ");
      document.getElementById("settingsMessage").textContent =
        "Reset to defaults.";
      await loadSelectedRow();
    });
  // End settings UI

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
    const count = tables.getCount(); // ClientResult<number>
    await context.sync();
    if (count.value === 0) {
      document.getElementById("fields").innerHTML =
        "<em>No tables in this workbook.</em>";
      return;
    }

    // Use the first table (or use `getItem("YourTableName")` if you know it)
    //const table = tables.getItemAt(0);
    const table = tables.getItem("All_Jobs");
    // Work only with the data body (no headers/totals)
    const body = table.getDataBodyRange();
    const hit = body.getIntersectionOrNullObject(activeCell.getEntireRow());

    // Load what we need
    body.load(["rowIndex", "columnCount"]);
    hit.load(["isNullObject", "rowIndex"]);
    await context.sync();

    // If the selection isn't in the table body, tell the user
    if (hit.isNullObject) {
      document.getElementById("fields").innerHTML =
        "<em>Select a row inside the table body.</em>";
      return;
    }

    // Compute the row index relative to the table body
    const relativeRowIndex = hit.rowIndex - body.rowIndex;
    const cols = table.columns;
    let myColumns = loadUserColumns();
    if (myColumns.length === 0) {
      myColumns = [...DEFAULT_COLUMNS];
    }
    console.log("mycolumns", myColumns);
    const colwithValues = [];
    for (let i = 0; i < myColumns.length; i++) {
      const colName = myColumns[i];
      const colValue = cols
        .getItem(colName)
        .getDataBodyRange()
        .getCell(relativeRowIndex, 0);
      colValue.load("values");
      colwithValues.push(colValue);
    }

    await context.sync();

    const values = [...colwithValues.map((c) => c.values[0])];
    const container = document.getElementById("fields");
    console.log("mycolumns", myColumns);
    container.innerHTML = values
      .map(
        (val, i) =>
          `<div class="myColumn"><label for="col_${i}">${
            myColumns[i]
          }</label>
        <input type="text" id="col_${i}"
        data-colName="${myColumns[i]}"
        data-rowindex="${relativeRowIndex}" value="${String(
            val[0] ?? ""
          )}" /></div>`
      )
      .join("");
  });
}

async function saveChanges() {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    //const table = wb.tables.getItemAt(0);
    const table = wb.tables.getItem("All_Jobs");
    const body = table.getDataBodyRange();
    const rowIndex = parseInt(
      document.querySelector(".myColumn input").dataset.rowindex,
      10
    );
    if (isNaN(rowIndex)) return;

    /* const newValues = [
      document.getElementById("col0").value,
      document.getElementById("col1").value,
      document.getElementById("col2").value,
      document.getElementById("col3").value,
    ]; */
    const cols = table.columns;

    cols.load("items/name,items/index");
    await context.sync();

    let myColumns = loadUserColumns();
    const colwithValues = {};
    for (let i = 0; i < myColumns.length; i++) {
      const colName = myColumns[i];
      const tableCol = cols.getItem(colName);
      const body = tableCol.getDataBodyRange();
      const cell     = body.getCell(rowIndex, 0);
      cell.load("values");
      tableCol.load("index");
      await context.sync();
      colwithValues[colName] = {
        value: cell.values[0][0],
        colIndex: tableCol.index
      };
    }
    
    console.log("colwithValues", colwithValues);
    for (let i = 0; i < myColumns.length; i++) {
      const colName = myColumns[i];
      
      const cell = body.getCell(rowIndex, colwithValues[colName].colIndex);
      cell.values = [[document.querySelector(
        `.myColumn #col_${i}`
      ).value]];
     
    }
    await context.sync();
    await loadSelectedRow();
  });
}
