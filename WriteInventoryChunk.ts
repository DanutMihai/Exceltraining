/**
 * WriteInventoryChunk  —  Office Script for the SIM inventory export flow.
 *
 * Called once per chunk from Power Automate ("Run script"). Writes SharePoint
 * rows into Table_query without touching the calculated columns.
 *
 * Why a script instead of "Add a row into a table": that action is one HTTP
 * call per row, roughly 1 row/second. 60,000 rows would take about 17 hours
 * and would not survive the flow's run duration limit. This writes 5,000 rows
 * in a single setValues() call.
 *
 * Column layout of Table_query (29 columns, A..AC):
 *   A..S   FirstName .. ID          <- written from SharePoint  (19 columns)
 *   T      IsPhoneValid             <- NOT written: it is a check formula
 *   U      IMEI                     <- written from SharePoint  (1 column)
 *   V..AC  PhoneClean .. HasError    <- NOT written: check formulas
 *
 * So each row of `rows` must be exactly 20 values, in this order:
 *   FirstName, LastName, PhoneNr, Plan, ICC_ID, GD_ID, StartDate, SIM_Country,
 *   Provider, CostCenter, UBR_Code, Legal entity, IsVR, BAN, SOC, Status,
 *   SIM Type, Email, ID,      <- 19 values, land in A..S
 *   IMEI                       <- 20th value, lands in U
 *
 * Everything must arrive as a STRING. ICC_ID is 18-22 digits and IMEI is 15;
 * Excel keeps only 15 significant digits, so anything sent as a JSON number is
 * silently truncated before it ever reaches the sheet.
 */

interface ChunkJob {
  rows: string[][];      // this chunk's rows, 20 strings each
  startIndex: number;    // 0-based offset of this chunk within the data body
  totalRows: number;     // total rows the finished export will have
  firstChunk: boolean;
  lastChunk: boolean;
}

function main(workbook: ExcelScript.Workbook, payload: string): string {
  const SHEET_NAME = "query";
  const TABLE_NAME = "Table_query";
  const HEADER_ROWS = 1;
  const LEFT_BLOCK = 19;   // A..S
  const IMEI_COL = 20;     // U, 0-based
  const TEXT_COLS = [2, 4, 6, 20];  // C PhoneNr, E ICC_ID, G StartDate, U IMEI

  const job: ChunkJob = JSON.parse(payload);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  const table = workbook.getTable(TABLE_NAME);
  const app = workbook.getApplication();

  // Each Run script call is its own session, so re-assert manual mode every
  // time. Without this Excel recalculates the whole sheet after every chunk.
  app.setCalculationMode(ExcelScript.CalculationMode.manual);

  if (job.firstChunk) {
    // Resize once, up front. Growing the table propagates the calculated
    // columns (V..AC) down all rows in one operation, so the script never
    // writes a formula itself and the two stay in sync automatically.
    const totalCols = table.getRange().getColumnCount();
    table.resize(
      sheet.getRangeByIndexes(0, 0, HEADER_ROWS + job.totalRows, totalCols)
    );

    // Force Text on the identifier columns BEFORE any data lands, so long
    // digit strings are not coerced to numbers on the way in.
    // Note: setNumberFormat (singular) takes one string for the whole range;
    // setNumberFormats (plural) is the string[][] overload.
    for (const col of TEXT_COLS) {
      sheet.getRangeByIndexes(HEADER_ROWS, col, job.totalRows, 1)
           .setNumberFormat("@");
    }
  }

  const n = job.rows.length;
  if (n > 0) {
    const top = HEADER_ROWS + job.startIndex;

    // A..S in one call.
    sheet.getRangeByIndexes(top, 0, n, LEFT_BLOCK)
         .setValues(job.rows.map(r => r.slice(0, LEFT_BLOCK)));

    // U on its own, stepping over the IsPhoneValid formula in T.
    sheet.getRangeByIndexes(top, IMEI_COL, n, 1)
         .setValues(job.rows.map(r => [r[LEFT_BLOCK]]));
  }

  if (job.lastChunk) {
    // Restore automatic calculation and force one full pass, so the delivered
    // file carries real values rather than a manual-mode workbook that looks
    // empty until the user presses F9.
    //
    // WATCH THIS STEP. It is the one most likely to hit the 120-second Run
    // script timeout on a large export. If it does, drop the calculate() call,
    // leave the mode automatic, and let Excel recalculate when the user opens
    // the file - or run this branch as its own separate Run script action.
    app.setCalculationMode(ExcelScript.CalculationMode.automatic);
    app.calculate(ExcelScript.CalculationType.fullRebuild);
  }

  return `${n} rows written at offset ${job.startIndex}`;
}
