// Keep track of which (row, columnName) cells are unlocked by the user.
const unlockedCells = new Set(); // keys like `${rowIndex}::${columnName}`

function initLockHandlers() {
  document.addEventListener("click", (ev) => {
    const btn = ev.target.closest(".lock-btn");
    if (!btn) return;

    const colName = btn.dataset.colname;
    const rowIndex = parseInt(btn.dataset.rowindex, 10);
    const locked = btn.dataset.locked === "1";

    // Toggle
    setUnlocked(rowIndex, colName, locked); // locked -> unlock (true); unlocked -> lock (false)
    // Update button & input immediately
    btn.dataset.locked = locked ? "0" : "1";
    btn.textContent = locked ? "🔓" : "🔒";
    btn.setAttribute("aria-label", locked ? "Lock" : "Unlock");

    // Flip the input's readonly and type for dates
    const input = document.querySelector(
      `#fields input[data-rowindex="${rowIndex}"][data-colname="${CSS.escape(
        colName
      )}"]`
    );
    if (input) {
      if (locked) {
        // We just unlocked
        input.readOnly = false;
        input.classList.remove("readonly");
        // If this is a date cell rendered as text, switch to type="date"
        if (
          input.type === "text" &&
          input.previousElementSibling?.classList?.contains("fx")
        ) {
          // We can detect date by seeing the meta.kind at render-time; easiest is to
          // replace the input if the displayed value is ISO-like; but simpler:
          // Whenever unlocked and the meta is Date, loadSelectedRow() gave type="date" already.
          // To keep it instant, you could re-call loadSelectedRow(). For now we keep as-is.
        }
      } else {
        // We just locked it
        input.readOnly = true;
        input.classList.add("readonly");
        // Keep current type; next full refresh will normalize
      }
    }
  });
}

// Helpers
function cellKey(rowIndex, colName) {
  return `${rowIndex}::${colName}`;
}
function isUnlocked(rowIndex, colName) {
  return unlockedCells.has(cellKey(rowIndex, colName));
}
function setUnlocked(rowIndex, colName, unlocked) {
  const key = cellKey(rowIndex, colName);
  if (unlocked) unlockedCells.add(key);
  else unlockedCells.delete(key);
}

const DATE_SYSTEM_1904 = false;

const MS_PER_DAY = 86400000;
// Excel serial of 1970-01-01 for each system
const EPOCH_SERIAL_1900 = 25569; // 1900 system
const EPOCH_SERIAL_1904 = 24107; // 1904 system (1900 - 1462)

function excelSerialToISODate(serial, date1904 = DATE_SYSTEM_1904) {
  if (serial == null || isNaN(serial)) return "";
  const epoch = date1904 ? EPOCH_SERIAL_1904 : EPOCH_SERIAL_1900;
  const ms = Math.round((serial - epoch) * MS_PER_DAY);
  const d = new Date(ms); // UTC ms → Date
  return d.toISOString().slice(0, 10); // "YYYY-MM-DD"
}

function isoDateToExcelSerial(iso, date1904 = DATE_SYSTEM_1904) {
  if (!iso) return null;
  const epoch = date1904 ? EPOCH_SERIAL_1904 : EPOCH_SERIAL_1900;
  // Force UTC midnight to avoid TZ shifts
  const ms = Date.parse(iso + "T00:00:00Z");
  return ms / MS_PER_DAY + epoch;
}

function inferCellKind(vJson, numberFormat) {
  if (!vJson) return "Unknown";
  let kind = vJson.basicType || "Unknown"; // "String"|"Double"|"Boolean"|"Error"
  if (vJson.type === "FormattedNumber" || kind === "Double") {
    if (typeof numberFormat === "string" && /[dymhs]/i.test(numberFormat)) {
      return "Date"; // numeric + date/time number format
    }
    return "Number";
  }
  if (kind === "Double") return "Number";
  return kind;
}

/**
 * Adds: { isFormula, formula, iso (for Date) }
 */
async function getRowCellMetaInCurrentContext(
  table,
  relativeRowIndex,
  columnNames
) {
  const header = table.getHeaderRowRange().load("values");
  const body = table.getDataBodyRange();
  const row = body.getRow(relativeRowIndex);

  // Load everything we need in one sync
  row.load(["valuesAsJson", "numberFormat", "text", "formulas"]);
  await table.context.sync();

  const headerNames = header.values[0];
  const rowJson = row.valuesAsJson[0];
  const rowNf = row.numberFormat[0];
  const rowText = row.text[0];
  const rowFx = row.formulas[0];

  const meta = {};

  for (const name of columnNames) {
    const idx = headerNames.indexOf(name);
    if (idx === -1) {
      meta[name] = {
        value: "",
        kind: "String",
        displayText: "",
        numberFormat: null,
        isFormula: false,
        notFound: true,
      };
      continue;
    }

    const vJson = rowJson[idx];
    const nf = rowNf[idx];
    const txt = rowText[idx];
    const fx = rowFx[idx];

    const isFormula = typeof fx === "string" && fx.trim().startsWith("="); // formula detection [2](https://stackoverflow.com/questions/76584702/unable-to-get-selected-range-values-when-column-selected-excel-javascript)
    const kind = inferCellKind(vJson, nf);

    // If it's a date, compute ISO from the *numeric* basicValue
    let iso = "";
    if (kind === "Date" && typeof vJson?.basicValue === "number") {
      iso = excelSerialToISODate(vJson.basicValue, DATE_SYSTEM_1904);
    }

    meta[name] = {
      value: vJson?.basicValue ?? null,
      kind,
      displayText: txt ?? "",
      numberFormat: typeof nf === "string" ? nf : null,
      isFormula,
      formula: isFormula ? fx : undefined,
      iso,
    };
  }

  return meta;
}

const DEFAULT_COLUMNS = [""];
const LS_KEY_COLUMNS = "tableHelper.columns";
let tableName = "";

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

  initLockHandlers();

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

async function loadSelectedRow() {
  await Excel.run(async (context) => {
    //const sheet = context.workbook.worksheets.getActiveWorksheet();
    const wb = context.workbook;
    const sheet = wb.worksheets.getActiveWorksheet();
    const tmpTable = sheet.tables.getItemAt(0);
    tmpTable.load("name");
    await context.sync();
    //const
    tableName = tmpTable.name;
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
    try {
      tables.getItem(tableName);
      console.log(`Using table: ${tableName}`);
      if (!tableName){
        document.getElementById("fields").innerHTML =
        `<em>Error loading table ${tableName}</em>`;
      return;
      }
    } catch (error) {
      console.error("Table not found:", error);
      document.getElementById("fields").innerHTML =
        `<em>Error loading table ${tableName}</em>`;
      return;
    }
    const table = tables.getItem(tableName);
   
    //const table = tables.getItem("All_Jobs");
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
    
    let myColumns = loadUserColumns();
    if (myColumns.length === 0) {
      myColumns = [...DEFAULT_COLUMNS];
    }

    // Build { ColumnName -> CellMeta } including formula flags

    const cellsMeta = await getRowCellMetaInCurrentContext(
      table,
      relativeRowIndex,
      myColumns
    );

    const html = myColumns
      .map((name, i) => {
        const meta = cellsMeta[name];
        if (!meta || meta.notFound) {
          return `
          <div class="field">
            <label>${name} <span style="color:#b33;">(header not found)</span></label>
            <input type="text" readonly class="readonly" value="" />
          </div>`;
        }

        // Decide input type
        const isDate = meta.kind === "Date";
        const isFormula = meta.isFormula;
        const unlocked = isFormula && isUnlocked(relativeRowIndex, name);
        const editable = !isFormula || unlocked;
         const isBoolean = meta.kind === "Boolean";

        // Show a lock button only for formulas
        const lockBtn = isFormula
          ? `<button class="lock-btn" type="button"
                   aria-label="${unlocked ? "Lock" : "Unlock"}"
                   data-colname="${name}" data-rowindex="${relativeRowIndex}"
                   data-locked="${unlocked ? "0" : "1"}">${
              unlocked ? "🔓" : "🔒"
            }</button>`
          : "";

        const badge = isFormula
          ? `<span class="pill">FORMULA</span>`
          : `<small style="opacity:.7">[${meta.kind}]</small>`;

        const formulaLine = isFormula
          ? `<div class="fx">${meta.formula}</div>`
          : "";

        // If editable and it's a date => <input type="date"> with ISO.
        const type = editable && isDate ? "date" : "text";
        const valueAttr =
          type === "date" ? meta.iso || "" : String(meta.displayText ?? "");
        const roAttrs = editable ? "" : `readonly class="readonly"`;

        return `
        <div class="field">
          <label for="col_${i}">${name} ${badge} ${lockBtn}</label>
          
          <input
            id="col_${i}"
            data-colname="${name}"
            data-rowindex="${relativeRowIndex}"
            ${isDate ? `type="date"` : isBoolean ? `type="checkbox"` : `type="text"` }
            ${roAttrs}
            value="${valueAttr}"
            ${isBoolean && valueAttr == 'TRUE' ? `checked` : ''} />
        </div>`;
      })
      .join("");

    fields.innerHTML = html;
    fields.setAttribute("data-rowindex", String(relativeRowIndex));
  });
}

async function saveChanges() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItemAt(0);
    const header = table.getHeaderRowRange().load("values");
    const body = table.getDataBodyRange();
    await context.sync();

    const fields = document.getElementById("fields");
    const rowIndex = parseInt(fields.getAttribute("data-rowindex"), 10);
    if (Number.isNaN(rowIndex)) return;

    const headerNames = header.values[0];
    const myColumns = loadUserColumns();

    // Need fresh meta to know which cells are formulas (and dates)
    const cellsMeta = await getRowCellMetaInCurrentContext(
      table,
      rowIndex,
      myColumns
    );

    for (let i = 0; i < myColumns.length; i++) {
      const name = myColumns[i];
      const meta = cellsMeta[name];
      const input = document.getElementById(`col_${i}`);
      if (!meta || !input || meta.notFound) continue;

      const idx = headerNames.indexOf(name);
      if (idx === -1) continue;

      // Skip if it's a formula cell and user did not unlock it
      if (meta.isFormula && !isUnlocked(rowIndex, name)) continue;

      const cell = body.getCell(rowIndex, idx);

      if (input.type === "date") {
        const iso = input.value;
        if (!iso) {
          cell.values = [[""]];
        } else {
          const serial = isoDateToExcelSerial(iso, DATE_SYSTEM_1904);
          cell.values = [[serial]];
          //cell.numberFormat = [["yyyy-mm-dd"]]; // ensure it stays a date [2](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablecollection?view=excel-js-preview)
        }
      } else if (input.type === "checkbox") {
        const checked = input.checked;
        cell.values = [[checked]];
        
      } else {
        let val = input.value ?? "";
        // If user typed a leading "=", treat it as a literal string (not a formula)
        if (val.trim().startsWith("=")) {
          cell.numberFormat = [["@"]]; // Text format (2D array API) [2](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablecollection?view=excel-js-preview)
        }
        cell.values = [[val]];
      }
    }

    await context.sync();
  });
}
