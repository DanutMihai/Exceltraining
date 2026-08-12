/**
 * ValidateAndMatchImport  —  the import engine.
 *
 * Runs against an uploaded workbook sitting in a SharePoint staging library.
 * For one page of rows it: re-validates every field from scratch, matches
 * against the SharePoint inventory, and returns both the create/update
 * decisions for your $batch builder and the report rows for the HTML report.
 *
 * WHY THE RULES ARE RE-IMPLEMENTED HERE RATHER THAN READ FROM THE WORKBOOK
 * The check columns are formulas in a file the user controls. They can paste
 * values over them, delete them, or hand back a file that never recalculated -
 * and in every one of those cases the columns cheerfully report "OK". This
 * script therefore ignores columns T and V..AC entirely and recomputes from
 * the raw data. The Excel formulas stay as the live preview the user edits
 * against; this file is the gate. They must be kept in step - if you change a
 * rule in one, change it in the other.
 *
 * PAGING
 * Office Scripts caps the response at 5 MB and the run at 120 seconds, so the
 * flow calls this once per page. Each call reads only its own slice of rows,
 * so cost per call is flat regardless of how big the workbook is. The
 * inventory is passed on every call; at 60k items x 3 fields that is ~4 MB
 * against a 30 MB parameter limit.
 *
 * FLOW SHAPE
 *   Get items (ICC_ID, PhoneNr, ID)  ->  Select to compact array
 *   Do until page > hasMore:
 *      Run script  ->  decisions + reportRows
 *      Build $batch body from decisions  ->  Send an HTTP request to SharePoint
 *      Append reportRows to varReportRows
 *   Compose REPORT json  ->  inject into the HTML template
 */

/* ------------------------------------------------------------------ types */

interface InventoryItem {
  id: number;
  iccId: string;
  phoneNr: string;
  // Optional extra fields. Any key matching a table header is diffed against
  // the workbook value to populate `changes` in the report.
  [key: string]: unknown;
}

interface ImportJob {
  tableName: string;
  country: string;        // country scope for this run; rows for others are skipped
  inventory: InventoryItem[];
  page: number;           // 0-based
  pageSize: number;       // rows per call; 2000 is a good starting point
  /**
   * "exceptions" (default) returns report rows only for failed / warning rows,
   * plus a skipped-by-country tally. "all" returns a row for every line.
   *
   * Use "exceptions". A 60k-row import produces ~10 MB of report JSON in "all"
   * mode, which blows the 5 MB response limit and would render a table no
   * human is going to read anyway. The counts still reconcile either way.
   */
  reportDetail: string;

  /** Set false to omit `decisions` from the response once batches are working. */
  includeDecisions: boolean;

  /**
   * Everything needed to emit ready-to-POST $batch bodies.
   *
   * Building the multipart string in Power Automate turned out to be a dead end:
   * the modern designer validates a Select's text-mode map as JSON, so a raw
   * multi-line body cannot be produced there, and the Apply-to-each alternative
   * is one string append per row. The script already holds every value, so it
   * builds the bodies here - which also lets the field-name mapping happen in
   * the same place, with proper JSON escaping for free.
   */
  batchConfig: BatchConfig;
}

interface BatchConfig {
  site: string;        // https://deutschebank.sharepoint.com/sites/simri
  listGuid: string;    // 6b659861-abd0-4e45-b74e-63e3f69f2648
  entityType: string;  // SP.Data.Global_x0020_SIM_x0020_InventoryListItem
  runId: string;
  batchSize: number;   // operations per batch; 100
  fieldMap: FieldMapEntry[];
}

/**
 * excel  - the Table_query header
 * internal - the SharePoint InternalName (Legal entity -> Legal_x0020_entity)
 * type   - TypeAsString. "DateTime", "Number", "Boolean" get converted;
 *          anything else is sent as text. Omit a field here and it is skipped.
 */
interface FieldMapEntry {
  excel: string;
  internal: string;
  type: string;
}

interface BatchRequest {
  boundary: string;   // goes in the Content-Type header of the HTTP action
  body: string;       // POST this verbatim
  rows: number[];     // Excel rows in this batch, in order, for response matching
}

interface FieldChange { field: string; from: string; to: string; }

interface Decision {
  excelRow: number;
  action: "create" | "update";
  spId: number | null;
  fields: { [key: string]: string };
}

interface ReportRow {
  excelRow: number;
  action: "created" | "updated" | "warning" | "skipped" | "failed";
  name: string;
  iccId: string;
  phoneNr: string;
  simCountry: string;
  spId: number | null;
  matchedOn: string;
  message: string;
  changes: FieldChange[];
}

interface CountryTally { country: string; rows: number; }

interface ImportResult {
  page: number;
  pageSize: number;
  totalDataRows: number;
  rowsInPage: number;
  hasMore: boolean;
  counts: { created: number; updated: number; warning: number; skipped: number; failed: number };
  skippedByCountry: CountryTally[];
  decisions: Decision[];
  reportRows: ReportRow[];
  batches: BatchRequest[];
  /**
   * This page's failed and warning rows as ready-to-append plain text, one per
   * line, CRLF separated. For the log list's WorkHistory column.
   *
   * Built here rather than in the flow because a Select action's map is
   * validated as JSON, so it cannot emit a bare string - the same wall the
   * $batch body ran into. The flow appends this to a string variable once per
   * page, which is 30 appends at 60k rows instead of one per row.
   */
  exceptionsText: string;
}

/** Fields pushed to SharePoint. Excludes the check columns and IsPhoneValid. */
const DATA_FIELDS: string[] = [
  "FirstName", "LastName", "PhoneNr", "Plan", "ICC_ID", "GD_ID", "StartDate",
  "SIM_Country", "Provider", "CostCenter", "UBR_Code", "Legal entity", "IsVR",
  "BAN", "SOC", "Status", "SIM Type", "Email", "ID", "IMEI"
];

/* ------------------------------------------------------------------- main */

/**
 * Returns the ImportResult as a JSON STRING, not an object.
 *
 * Power Automate builds a schema from main's return type, and it cannot
 * describe an index signature - Decision.fields has dynamic keys, one per
 * SharePoint column, so the editor rejects it with
 *     "Type literal contains invalid type 'Unrecognized'".
 * Enumerating all 20 fields as named properties would work, but "Legal entity"
 * and "SIM Type" contain spaces and would have to be renamed, which then has to
 * be undone when building the $batch body.
 *
 * A string return sidesteps the schema entirely. In the flow, parse it:
 *     json(body('Run_script')?['result'])
 * and everything below is addressable as normal, e.g.
 *     body('Parse_result')?['decisions']
 *     body('Parse_result')?['counts']?['failed']
 */
function main(workbook: ExcelScript.Workbook, payload: string): string {
  const job: ImportJob = JSON.parse(payload);
  const pageSize: number = Math.max(1, Math.floor(job.pageSize || 2000));
  const page: number = Math.max(0, Math.floor(job.page || 0));

  const table: ExcelScript.Table = workbook.getTable(job.tableName || "Table_query");
  if (!table) { throw new Error("Table '" + job.tableName + "' not found in the uploaded workbook."); }

  const sheet: ExcelScript.Worksheet = table.getWorksheet();
  const headerRange: ExcelScript.Range = table.getHeaderRowRange();
  const headers: string[] = headerRange.getValues()[0].map((h) => String(h).trim());
  const headerRowIndex: number = headerRange.getRowIndex();
  const firstColumn: number = headerRange.getColumnIndex();
  const totalDataRows: number = table.getRowCount();

  const col: { [key: string]: number } = {};
  for (let i = 0; i < headers.length; i++) { col[headers[i].toLowerCase()] = i; }
  const at = (row: (string | number | boolean)[], name: string): string => {
    const i: number = col[name.toLowerCase()];
    return i === undefined ? "" : cellToText(row[i]);
  };

  const reference: Reference = readConfig(workbook);

  // Two indexes, built once. O(1) lookups.
  // Doing this join in Power Automate would be a Filter array per row - O(n^2),
  // which at 60k rows does not finish.
  const byId: Map<number, InventoryItem> = new Map<number, InventoryItem>();
  const byKey: Map<string, InventoryItem> = new Map<string, InventoryItem>();
  const duplicateKeys: Map<string, number> = new Map<string, number>();
  for (const item of (job.inventory || [])) {
    byId.set(Number(item.id), item);
    const k: string = matchKey(String(item.iccId ?? ""), String(item.phoneNr ?? ""));
    // If SharePoint itself holds two items with the same ICC_ID + PhoneNr, a
    // keyless Excel row cannot be resolved to one of them. Record it and fail
    // those rows rather than silently overwriting whichever landed last.
    if (byKey.has(k)) { duplicateKeys.set(k, (duplicateKeys.get(k) || 1) + 1); }
    byKey.set(k, item);
  }

  // Scan the whole ID / ICC_ID / PhoneNr columns once per call, so two Excel
  // rows targeting the same SharePoint item are caught even when they land on
  // different pages. Three columns over 60k rows is a cheap read; without it,
  // row 4 quietly overwrites row 3's update and the report shows both as fine.
  const targets: TargetScan = scanTargets(sheet, headerRowIndex, firstColumn, headers, totalDataRows, col);

  const start: number = page * pageSize;
  const take: number = Math.max(0, Math.min(pageSize, totalDataRows - start));
  const counts = { created: 0, updated: 0, warning: 0, skipped: 0, failed: 0 };
  const decisions: Decision[] = [];
  const reportRows: ReportRow[] = [];
  const skippedByCountry: Map<string, number> = new Map<string, number>();

  if (take > 0) {
    const block: (string | number | boolean)[][] = sheet
      .getRangeByIndexes(headerRowIndex + 1 + start, firstColumn, take, headers.length)
      .getValues();

    for (let r = 0; r < block.length; r++) {
      const row = block[r];
      const excelRow: number = headerRowIndex + 2 + start + r;   // 1-based sheet row

      const iccRaw: string = at(row, "ICC_ID");
      const phoneRaw: string = at(row, "PhoneNr");
      const country: string = at(row, "SIM_Country");
      const name: string = (at(row, "FirstName") + " " + at(row, "LastName")).trim();

      // Blank spare rows are not an error, they are just empty.
      if (iccRaw === "" && phoneRaw === "" && at(row, "FirstName") === "") { continue; }

      // ---- country scope -------------------------------------------------
      if (country.toLowerCase() !== String(job.country || "").toLowerCase()) {
        counts.skipped++;
        const label: string = country === "" ? "(blank)" : country;
        skippedByCountry.set(label, (skippedByCountry.get(label) || 0) + 1);
        pushRow(reportRows, job.reportDetail, {
          excelRow, action: "skipped", name, iccId: iccRaw, phoneNr: phoneRaw,
          simCountry: country, spId: null, matchedOn: "-",
          message: "SIM_Country is " + label + "; this run is scoped to " + job.country + ".",
          changes: []
        });
        continue;
      }

      // ---- validation ----------------------------------------------------
      const verdict: Verdict = validateRow(row, at, reference);
      if (verdict.errors.length > 0) {
        counts.failed++;
        pushRow(reportRows, job.reportDetail, {
          excelRow, action: "failed", name, iccId: iccRaw, phoneNr: phoneRaw,
          simCountry: country, spId: null, matchedOn: "-",
          message: verdict.errors.join("  //  ") + "  Row rejected; nothing written.",
          changes: []
        });
        continue;
      }

      // ---- match ---------------------------------------------------------
      // Order: explicit ID wins, then ICC_ID + PhoneNr, then create.
      const key: string = matchKey(iccRaw, phoneRaw);
      const idText: string = at(row, "ID").trim();
      let hit: InventoryItem = undefined;
      let matchedOn: string = "-";
      let matchError: string = "";

      if (idText !== "") {
        if (!isDigits(idText)) {
          matchError = "ID: \"" + idText + "\" is not a number. Leave it exactly as exported, "
                     + "or clear it to create a new item.";
        } else if (targets.idCounts.get(idText) > 1) {
          matchError = "ID " + idText + " is used by " + targets.idCounts.get(idText)
                     + " rows in this file. Each SharePoint item may be targeted once per import.";
        } else {
          hit = byId.get(Number(idText));
          if (!hit) {
            // Never silently create here. A stale ID usually means the item was
            // deleted, or the row was pasted in from another country's export -
            // creating a duplicate would be the worst possible response.
            matchError = "ID " + idText + " was not found in the " + job.country
                       + " inventory. It may have been deleted, or the row may belong to "
                       + "another country. Clear the ID column to create this as a new item.";
          } else {
            matchedOn = "ID";
          }
        }
      } else if (duplicateKeys.has(key)) {
        matchError = "ICC_ID + PhoneNr matches " + duplicateKeys.get(key)
                   + " existing SharePoint items, so the target is ambiguous. "
                   + "Fill in the ID column to say which one.";
      } else if (targets.keyCounts.get(key) > 1) {
        matchError = "ICC_ID + PhoneNr appears on " + targets.keyCounts.get(key)
                   + " rows in this file. Remove the duplicates or fill in the ID column.";
      } else {
        hit = byKey.get(key);
        if (hit) { matchedOn = "ICC_ID + PhoneNr"; }
      }

      if (matchError !== "") {
        counts.failed++;
        pushRow(reportRows, job.reportDetail, {
          excelRow, action: "failed", name, iccId: iccRaw, phoneNr: phoneRaw,
          simCountry: country, spId: null, matchedOn: "-",
          message: matchError + " Row rejected; nothing written.", changes: []
        });
        continue;
      }

      const fields: { [key: string]: string } = {};
      for (const f of DATA_FIELDS) { fields[f] = at(row, f); }
      const changes: FieldChange[] = hit ? diff(hit, fields) : [];
      const warned: boolean = verdict.warnings.length > 0;

      if (hit) {
        decisions.push({ excelRow, action: "update", spId: hit.id, fields });
        if (warned) { counts.warning++; } else { counts.updated++; }
        pushRow(reportRows, job.reportDetail, {
          excelRow, action: warned ? "warning" : "updated", name,
          iccId: iccRaw, phoneNr: phoneRaw, simCountry: country,
          spId: hit.id, matchedOn,
          message: warned ? "Updated. " + verdict.warnings.join("  //  ") : "Item updated.",
          changes
        });
      } else {
        decisions.push({ excelRow, action: "create", spId: null, fields });
        if (warned) { counts.warning++; } else { counts.created++; }
        pushRow(reportRows, job.reportDetail, {
          excelRow, action: warned ? "warning" : "created", name,
          iccId: iccRaw, phoneNr: phoneRaw, simCountry: country,
          spId: null, matchedOn: "-",
          message: (warned ? verdict.warnings.join("  //  ") + " " : "")
                 + "No ID and no ICC_ID + PhoneNr match. New list item created.",
          changes: []
        });
      }
    }
  }

  const skippedTally: CountryTally[] = [];
  skippedByCountry.forEach((n: number, c: string) => { skippedTally.push({ country: c, rows: n }); });

  const result: ImportResult = {
    page, pageSize, totalDataRows, rowsInPage: take,
    hasMore: start + take < totalDataRows,
    counts, skippedByCountry: skippedTally,
    decisions: job.includeDecisions === true ? decisions : [],
    reportRows,
    batches: buildBatches(decisions, job.batchConfig),
    exceptionsText: buildExceptionsText(reportRows)
  };
  return JSON.stringify(result);
}

/**
 * Failed and warning rows as plain text, one per line.
 *
 * Only those two actions, whatever reportDetail is set to - in "all" mode the
 * report carries every row, but a log column does not want 60,000 lines of
 * "Item updated."
 *
 * Capped at 200 lines per page. SharePoint's multi-line text column tops out at
 * 63,999 characters, and the full detail lives in the HTML report anyway.
 */
function buildExceptionsText(rows: ReportRow[]): string {
  const CRLF: string = "\r\n";
  const lines: string[] = [];
  let skipped: number = 0;
  for (const r of rows) {
    if (r.action !== "failed" && r.action !== "warning") { continue; }
    if (lines.length >= 200) { skipped++; continue; }
    lines.push(r.action + " · row " + r.excelRow + " · " + r.message);
  }
  if (skipped > 0) {
    lines.push("… and " + skipped + " more on this page - see the full report.");
  }
  return lines.join(CRLF);
}

/* ------------------------------------------------------------ batch bodies */

/**
 * Turns decisions into complete multipart/mixed $batch payloads.
 *
 * One changeset per operation, so a rejected row does not roll back the other
 * 99 - per-row reporting depends on that isolation. The changeset suffix is the
 * Excel row number, so the response can be matched back to the spreadsheet
 * rather than by position.
 *
 * Line endings are CRLF and blank lines are load-bearing: one after the
 * changeset's Content-Transfer-Encoding, one between the request headers and
 * the JSON body. Miss either and SharePoint reads the body as a header.
 */
function buildBatches(decisions: Decision[], cfg: BatchConfig): BatchRequest[] {
  const out: BatchRequest[] = [];
  if (decisions.length === 0) { return out; }

  // Fail loudly. Returning an empty array here is the worst possible outcome:
  // the Apply to each finds nothing to send, no rows are written, and the run
  // still reports "3 updated" because the counts were tallied before this
  // point. A misconfigured payload must stop the flow, not produce a clean
  // report about work that never happened.
  if (!cfg) {
    throw new Error(
      "batchConfig is missing from the payload. " + decisions.length +
      " rows were matched but no $batch bodies can be built, so nothing would " +
      "be written. Add batchConfig (site, listGuid, entityType, runId, " +
      "batchSize, fieldMap) to the Run script payload."
    );
  }
  const missing: string[] = [];
  if (!cfg.site) { missing.push("site"); }
  if (!cfg.listGuid) { missing.push("listGuid"); }
  if (!cfg.entityType) { missing.push("entityType"); }
  if (!cfg.runId) { missing.push("runId"); }
  if (missing.length > 0) {
    throw new Error("batchConfig is incomplete - missing: " + missing.join(", ") +
                    ". Received keys: " + Object.keys(cfg).join(", "));
  }
  if (!cfg.fieldMap || cfg.fieldMap.length === 0) {
    throw new Error(
      "batchConfig.fieldMap is empty. Every item body would contain only " +
      "__metadata and no field values, so SharePoint would accept the writes " +
      "and store nothing. Populate it from " +
      "_api/web/lists(guid'...')/fields?$select=Title,InternalName,TypeAsString"
    );
  }

  const size: number = Math.max(1, Math.floor(cfg.batchSize || 100));
  const CRLF: string = "\r\n";
  const itemsUrl: string = cfg.site + "/_api/web/lists(guid'" + cfg.listGuid + "')/items";

  for (let offset = 0; offset < decisions.length; offset += size) {
    const slice: Decision[] = decisions.slice(offset, offset + size);
    const boundary: string = "batch_" + cfg.runId + "_" + (offset / size);
    const rows: number[] = [];
    let body: string = "";

    for (const d of slice) {
      const cs: string = "changeset_" + cfg.runId + "_" + d.excelRow;
      const isUpdate: boolean = d.action === "update";
      const url: string = isUpdate ? itemsUrl + "(" + d.spId + ")" : itemsUrl;
      rows.push(d.excelRow);

      body += "--" + boundary + CRLF;
      body += "Content-Type: multipart/mixed; boundary=" + cs + CRLF + CRLF;
      body += "--" + cs + CRLF;
      body += "Content-Type: application/http" + CRLF;
      body += "Content-Transfer-Encoding: binary" + CRLF + CRLF;
      body += "POST " + url + " HTTP/1.1" + CRLF;
      body += "Content-Type: application/json;odata=verbose" + CRLF;
      body += "Accept: application/json;odata=verbose" + CRLF;
      if (isUpdate) {
        // IF-MATCH * overwrites regardless of version. Swap in the item's etag
        // if a concurrent edit should fail rather than be silently overwritten.
        body += "IF-MATCH: *" + CRLF;
        // MERGE, never PUT: PUT blanks every field not present in the payload.
        body += "X-HTTP-Method: MERGE" + CRLF;
      }
      body += CRLF;
      body += buildItemJson(d.fields, cfg) + CRLF + CRLF;
      body += "--" + cs + "--" + CRLF;
    }
    body += "--" + boundary + "--";

    out.push({ boundary, body, rows });
  }
  return out;
}

/** JSON.stringify handles the escaping, so a quote in a name cannot break the body. */
function buildItemJson(
  fields: { [key: string]: string },
  cfg: BatchConfig
): string {
  const payload: { [key: string]: unknown } = {};
  payload["__metadata"] = { "type": cfg.entityType };

  for (const entry of (cfg.fieldMap || [])) {
    // ID is the target of an update and assigned by SharePoint on create.
    // Sending it either way is an error.
    if (entry.internal === "ID" || entry.excel === "ID") { continue; }
    const raw: string = fields[entry.excel];
    if (raw === undefined) { continue; }
    payload[entry.internal] = convertValue(raw, entry.type);
  }
  return JSON.stringify(payload);
}

/**
 * SharePoint types are strict. A Date column rejects "15-03-2026"; it wants
 * ISO 8601. A Number column rejects "" but accepts null.
 */
function convertValue(raw: string, type: string): unknown {
  const s: string = String(raw == null ? "" : raw).trim();

  if (type === "DateTime") {
    if (s === "") { return null; }
    return toIso(s);
  }
  if (type === "Number" || type === "Currency") {
    if (s === "") { return null; }
    const n: number = Number(s);
    return isNaN(n) ? null : n;
  }
  if (type === "Boolean") {
    const l: string = s.toLowerCase();
    return l === "true" || l === "yes" || l === "1" || l === "y";
  }
  if (type === "MultiChoice") {
    // Multi-select choice columns need a typed collection, not a bare string.
    // Single-select "Choice" does NOT - it takes the plain value, which is why
    // it falls through to the text branch below.
    const parts: string[] = s === "" ? [] : s.split(";");
    const cleaned: string[] = [];
    for (const p of parts) { if (p.trim() !== "") { cleaned.push(p.trim()); } }
    return { "__metadata": { "type": "Collection(Edm.String)" }, "results": cleaned };
  }
  if (type === "Lookup" || type === "User" || type === "LookupMulti" || type === "UserMulti") {
    // These take a numeric ID against a different property name (FieldId, not
    // Field), so a display string here would 400 the whole batch. Fail with a
    // message that says what to do instead.
    throw new Error(
      "fieldMap contains a " + type + " column. SharePoint needs its numeric ID " +
      "written to the companion property - map it as {\"excel\":\"<header>\"," +
      "\"internal\":\"<InternalName>Id\",\"type\":\"Number\"} and put the ID in " +
      "that Excel column, or drop it from fieldMap to leave the field untouched."
    );
  }
  return s;
}

/** dd-mm-yyyy (any of - / .) or an Excel serial -> ISO 8601. "" if unparseable. */
function toIso(s: string): string {
  const serial: number = Number(s);
  if (!isNaN(serial) && serial > 20000 && serial < 80000) {
    return new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString();
  }
  if (s.length === 10) {
    const day: number = Number(s.substring(0, 2));
    const month: number = Number(s.substring(3, 5));
    const year: number = Number(s.substring(6, 10));
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
      return new Date(Date.UTC(year, month - 1, day)).toISOString();
    }
  }
  return "";
}

/* --------------------------------------------------------- duplicate scan */

interface TargetScan {
  idCounts: Map<string, number>;
  keyCounts: Map<string, number>;
}

/**
 * Counts how many rows in the WHOLE table target each SharePoint ID and each
 * ICC_ID + PhoneNr pair. Runs on every call so the answer is page-independent:
 * two rows aiming at the same item are caught even when they are 40,000 rows
 * apart. Reads three columns, so it costs one range read per call.
 */
function scanTargets(
  sheet: ExcelScript.Worksheet,
  headerRowIndex: number,
  firstColumn: number,
  headers: string[],
  totalDataRows: number,
  col: { [key: string]: number }
): TargetScan {
  const idCounts: Map<string, number> = new Map<string, number>();
  const keyCounts: Map<string, number> = new Map<string, number>();
  if (totalDataRows <= 0) { return { idCounts, keyCounts }; }

  // Read the three columns individually rather than the whole table. Excel I/O
  // dominates this script's runtime, and this scan runs on EVERY page - reading
  // all 29 columns cost 29x more than necessary. At 60k rows over 30 pages that
  // is 52 million cells instead of 5.4 million.
  const readColumn = (name: string): string[] => {
    const i: number = col[name];
    const out: string[] = [];
    if (i === undefined) { return out; }
    const values: (string | number | boolean)[][] = sheet
      .getRangeByIndexes(headerRowIndex + 1, firstColumn + i, totalDataRows, 1)
      .getValues();
    for (const row of values) { out.push(cellToText(row[0]).trim()); }
    return out;
  };

  const ids: string[] = readColumn("id");
  const iccs: string[] = readColumn("icc_id");
  const phones: string[] = readColumn("phonenr");

  for (let r = 0; r < totalDataRows; r++) {
    const id: string = r < ids.length ? ids[r] : "";
    if (id !== "") { idCounts.set(id, (idCounts.get(id) || 0) + 1); }
    const icc: string = r < iccs.length ? iccs[r] : "";
    const phone: string = r < phones.length ? phones[r] : "";
    if (icc !== "" || phone !== "") {
      const k: string = matchKey(icc, phone);
      keyCounts.set(k, (keyCounts.get(k) || 0) + 1);
    }
  }
  return { idCounts, keyCounts };
}

/**
 * In "exceptions" mode only failed and warning rows are returned in detail.
 * Created / updated / skipped still increment the counts, so the report's
 * totals reconcile - you just do not carry 60,000 "Item updated." lines
 * through the flow and into the browser.
 */
function pushRow(sink: ReportRow[], mode: string, row: ReportRow): void {
  if (mode === "all") { sink.push(row); return; }
  if (row.action === "failed" || row.action === "warning") { sink.push(row); }
}

/* -------------------------------------------------------------- reference */

interface Reference {
  dialByCountry: Map<string, string>;
  statuses: string[];
  simTypes: string[];
}

/**
 * Reads tblCountries and the two allowed-value lists off the Config sheet, so
 * the dial codes and pick lists stay maintained in one place - the workbook -
 * rather than hardcoded here.
 */
function readConfig(workbook: ExcelScript.Workbook): Reference {
  const dialByCountry: Map<string, string> = new Map<string, string>();
  const countries: ExcelScript.Table = workbook.getTable("tblCountries");
  if (countries) {
    const rows: (string | number | boolean)[][] = countries.getRangeBetweenHeaderAndTotal().getValues();
    const head: string[] = countries.getHeaderRowRange().getValues()[0].map((h) => String(h).trim().toLowerCase());
    const ni: number = head.indexOf("name");
    const di: number = head.indexOf("dialcode");
    for (const row of rows) {
      const nm: string = cellToText(row[ni]).trim();
      const dc: string = cellToText(row[di]).trim();
      if (nm !== "" && dc !== "") { dialByCountry.set(nm.toLowerCase(), dc); }
    }
  }

  const cfg: ExcelScript.Worksheet = workbook.getWorksheet("Config");
  const readList = (column: number): string[] => {
    const out: string[] = [];
    if (!cfg) { return out; }
    const values: (string | number | boolean)[][] =
      cfg.getRangeByIndexes(1, column, 20, 1).getValues();
    for (const v of values) {
      const s: string = cellToText(v[0]).trim();
      if (s !== "") { out.push(s); }
    }
    return out;
  };
  return { dialByCountry, statuses: readList(6), simTypes: readList(7) };  // G, H
}

/* ------------------------------------------------------------- validation */

interface Verdict { errors: string[]; warnings: string[]; }

function validateRow(
  row: (string | number | boolean)[],
  at: (row: (string | number | boolean)[], name: string) => string,
  reference: Reference
): Verdict {
  const errors: string[] = [];
  const warnings: string[] = [];

  // ---- PhoneNr ----------------------------------------------------------
  const phone: string = normalisePhone(at(row, "PhoneNr"));
  const country: string = at(row, "SIM_Country");
  if (phone === "") {
    errors.push("PhoneNr: empty or contains no digits.");
  } else if (!isDigits(phone)) {
    errors.push("PhoneNr: digits only - no letters or symbols.");
  } else if (phone.length < 8) {
    errors.push("PhoneNr: Too short - " + phone.length + " digits (min 8).");
  } else if (phone.length > 15) {
    errors.push("PhoneNr: Too long - " + phone.length + " digits (max 15).");
  } else {
    const dial: string = reference.dialByCountry.get(country.toLowerCase());
    if (!dial) {
      warnings.push("WARNING: SIM_Country \"" + country + "\" is not in tblCountries; "
                  + "the country code could not be checked.");
    } else if (phone.substring(0, dial.length) !== dial) {
      errors.push("PhoneNr: wrong country code - must start with +" + dial + " for " + country + ".");
    }
  }

  // ---- ICC_ID -----------------------------------------------------------
  const icc: string = at(row, "ICC_ID").trim();
  if (icc === "") {
    errors.push("ICC_ID: empty.");
  } else if (!isDigits(icc)) {
    errors.push("ICC_ID: digits only - no +, -, spaces or any other character.");
  } else if (icc.substring(0, 2) !== "89") {
    errors.push("ICC_ID: must start with 89 (ITU-T E.118 prefix).");
  } else if (icc.length < 18) {
    errors.push("ICC_ID: Too short - " + icc.length + " digits (min 18).");
  } else if (icc.length > 22) {
    errors.push("ICC_ID: Too long - " + icc.length + " digits (max 22).");
  } else if (!luhnOk(icc)) {
    warnings.push("WARNING: ICC_ID failed its Luhn check digit - likely a typo or transposition.");
  }

  // ---- IMEI (optional field) --------------------------------------------
  const imei: string = at(row, "IMEI").trim();
  if (imei !== "") {
    if (!isDigits(imei)) {
      errors.push("IMEI: digits only - no +, -, spaces or any other character.");
    } else if (imei.length !== 15) {
      errors.push("IMEI: must be exactly 15 digits - this one has " + imei.length + ".");
    } else if (!luhnOk(imei)) {
      warnings.push("WARNING: IMEI failed its Luhn check digit - likely a typo or transposition.");
    }
  }

  // ---- StartDate ---------------------------------------------------------
  const dateVerdict: string = checkStartDate(row, at);
  if (dateVerdict !== "") {
    if (dateVerdict.indexOf("WARNING:") === 0) { warnings.push(dateVerdict); }
    else { errors.push("StartDate: " + dateVerdict); }
  }

  // ---- Status / SIM Type -------------------------------------------------
  const status: string = at(row, "Status").trim();
  if (status !== "" && reference.statuses.length > 0) {
    if (reference.statuses.indexOf(status) < 0) {
      const ci: number = indexOfIgnoreCase(reference.statuses, status);
      if (ci >= 0) {
        warnings.push("WARNING: Status has the wrong capitalisation - should be \""
                    + reference.statuses[ci] + "\".");
      } else {
        errors.push("Status: must be one of " + reference.statuses.join(", ") + ".");
      }
    }
  }
  const simType: string = at(row, "SIM Type").trim();
  if (simType !== "" && reference.simTypes.length > 0) {
    if (reference.simTypes.indexOf(simType) < 0) {
      const ci: number = indexOfIgnoreCase(reference.simTypes, simType);
      if (ci >= 0) {
        warnings.push("WARNING: SIM Type has the wrong capitalisation - should be \""
                    + reference.simTypes[ci] + "\".");
      } else {
        errors.push("SIM Type: must be one of " + reference.simTypes.join(", ") + ".");
      }
    }
  }

  return { errors, warnings };
}

/** "" when fine, otherwise a message (prefixed WARNING: for non-blocking). */
function checkStartDate(
  row: (string | number | boolean)[],
  at: (row: (string | number | boolean)[], name: string) => string
): string {
  const raw: string = at(row, "StartDate").trim();
  if (raw === "") { return ""; }

  // A real Excel date arrives as a serial number.
  const asNumber: number = Number(raw);
  if (raw !== "" && !isNaN(asNumber) && raw.indexOf("-") < 0 && raw.indexOf("/") < 0
      && raw.indexOf(".") < 0 && asNumber > 20000 && asNumber < 80000) {
    const d: Date = new Date(Date.UTC(1899, 11, 30) + asNumber * 86400000);
    const year: number = d.getUTCFullYear();
    if (year < 2000 || year > 2100) {
      return "WARNING: StartDate is a real date but falls outside 2000-2100 - check the day/month order.";
    }
    return "";
  }

  if (raw.length !== 10) {
    return "must be dd-mm-yyyy - 10 characters, this one has " + raw.length + ".";
  }
  const s1: string = raw.charAt(2);
  const s2: string = raw.charAt(5);
  if (s1 !== s2 || "-/.".indexOf(s1) < 0) {
    return "separators must be in positions 3 and 6, matching, and be - or / or .";
  }
  const dd: string = raw.substring(0, 2);
  const mm: string = raw.substring(3, 5);
  const yy: string = raw.substring(6, 10);
  if (!isDigits(dd) || !isDigits(mm) || !isDigits(yy)) {
    return "day, month and year must all be digits.";
  }
  const day: number = Number(dd), month: number = Number(mm), year: number = Number(yy);
  if (day < 1 || day > 31) { return "day out of range - dd-mm-yyyy puts the DAY first."; }
  if (month < 1 || month > 12) { return "month out of range - dd-mm-yyyy puts the MONTH second."; }
  if (year < 2000 || year > 2100) { return "year out of range (2000-2100)."; }
  const probe: Date = new Date(Date.UTC(year, month - 1, day));
  if (probe.getUTCDate() !== day || probe.getUTCMonth() !== month - 1) {
    return "not a real calendar date (e.g. 31-02).";
  }
  return "";
}

/* ----------------------------------------------------------------- helpers */

/** Strips + space nbsp - ( ) . then converts a leading 00 to bare digits. */
function normalisePhone(raw: string): string {
  let s: string = String(raw == null ? "" : raw).trim();
  const drop: string[] = ["+", " ", " ", "-", "(", ")", "."];
  for (const ch of drop) { s = s.split(ch).join(""); }
  if (s.substring(0, 2) === "00") { s = s.substring(2); }
  return s;
}

function matchKey(icc: string, phone: string): string {
  return String(icc == null ? "" : icc).trim() + "|" + normalisePhone(phone);
}

function isDigits(s: string): boolean {
  if (s.length === 0) { return false; }
  for (let i = 0; i < s.length; i++) {
    const c: number = s.charCodeAt(i);
    if (c < 48 || c > 57) { return false; }
  }
  return true;
}

/** Luhn mod-10 over the whole string including its trailing check digit. */
function luhnOk(s: string): boolean {
  let sum: number = 0;
  let double: boolean = false;
  for (let i = s.length - 1; i >= 0; i--) {
    let d: number = s.charCodeAt(i) - 48;
    if (double) { d *= 2; if (d > 9) { d -= 9; } }
    sum += d;
    double = !double;
  }
  return sum % 10 === 0;
}

function indexOfIgnoreCase(list: string[], value: string): number {
  const v: string = value.toLowerCase();
  for (let i = 0; i < list.length; i++) {
    if (list[i].toLowerCase() === v) { return i; }
  }
  return -1;
}

/**
 * Excel hands back numbers, booleans and strings. Long identifiers must never
 * go through Number(), so everything is stringified without arithmetic.
 */
function cellToText(v: string | number | boolean): string {
  if (v === null || v === undefined) { return ""; }
  if (typeof v === "string") { return v; }
  if (typeof v === "boolean") { return v ? "TRUE" : "FALSE"; }
  return String(v);
}

/** Diffs the SharePoint item against the workbook row, for the report. */
function diff(item: InventoryItem, fields: { [key: string]: string }): FieldChange[] {
  const out: FieldChange[] = [];
  for (const field of DATA_FIELDS) {
    if (field === "ID") { continue; }
    const before: unknown = item[field];
    if (before === undefined || before === null) { continue; }
    const from: string = String(before);
    const to: string = fields[field] === undefined ? "" : fields[field];
    if (from !== to) { out.push({ field, from, to }); }
  }
  return out;
}
