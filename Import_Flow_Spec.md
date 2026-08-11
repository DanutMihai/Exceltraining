# SIM Inventory Import — build spec

Single source of truth. Everything earlier is deleted; nothing here is superseded.

**Environment**

| | |
|---|---|
| Site | `https://deutschebank.sharepoint.com/sites/simri` |
| List | Global SIM Inventory |
| List GUID | `6b659861-abd0-4e45-b74e-63e3f69f2648` |
| Entity type | `SP.Data.Global_x0020_SIM_x0020_InventoryListItem` |
| Workbook table | `Table_query` (29 columns, A–AC) |
| Office Script | `ValidateAndMatchImport` |
| Report template | `Import_Report_Template.html` |

**Division of labour.** Power Automate orchestrates and does I/O. The Office Script does all
the thinking: validation, matching, and building the `$batch` payloads. Nothing that needs a
loop over rows happens in the designer.

**Naming.** Power Automate converts spaces to underscores in expressions — an action named
`Get inventory` is referenced as `body('Get_inventory')`. Use the names given here exactly.

---

## 0. Prerequisites

**Index `SIM_Country`.** List settings → Indexed columns. Without it, `Get items` with a filter
fails past 5,000 items with the list view threshold error.

**Trigger concurrency = 1.** Trigger → Settings → Concurrency Control → On → Degree of
parallelism `1`. Two simultaneous imports of the same country both read the inventory before
either writes, both decide "create" for the same row, and duplicate it. This is the most likely
way to corrupt the list.

**Two libraries:** `SIM Imports/Staging` and `SIM Imports/Reports`. Keep both away from the
library holding the master template.

**Get the field map** (one-off). Run in a browser or a scratch flow:

```
_api/web/lists(guid'6b659861-abd0-4e45-b74e-63e3f69f2648')/fields?$select=Title,InternalName,TypeAsString&$filter=Hidden eq false and ReadOnlyField eq false
```

You need this to fill `batchConfig.fieldMap` in §5a. A column missing from the map is silently
not written.

---

## 1. Trigger — PowerApps (V2)

| Input | Type | Name |
|---|---|---|
| File | File | `FileContent` |
| Text | Text | `Country` |
| Text | Text | `ActionedBy` |

References: `triggerBody()?['file']?['contentBytes']`, `triggerBody()?['text']` (Country),
`triggerBody()?['text_1']` (ActionedBy). Confirm the exact suffixes from your first run — they
depend on input order.

---

## 2. Initialize variables

Root level only; Power Automate rejects `Initialize variable` inside a Scope or loop.

| Name | Type | Initial value |
|---|---|---|
| `varRunId` | String | `guid()` |
| `varFileName` | String | `concat(formatDateTime(utcNow(),'yyyy-MM-dd_HH-mm-ss'),'_',triggerBody()?['text'],'_SIM_Import.xlsx')` |
| `varPage` | Integer | `0` |
| `varHasMore` | Boolean | `true` |
| `varReportRows` | Array | `[]` |
| `varCreated` | Integer | `0` |
| `varUpdated` | Integer | `0` |
| `varWarning` | Integer | `0` |
| `varSkipped` | Integer | `0` |
| `varFailed` | Integer | `0` |

Five separate counters rather than one object: Power Automate can't increment a property inside
an object variable, so an object means rebuilding the whole thing on every page.

---

## 3. Scope — Main

Everything from §3a to §8 goes inside one Scope named `Scope - Main`. That is what makes the
failure handler in §9 possible.

### 3a. Condition — `Validate inputs`

```
or(empty(triggerBody()?['text']), empty(triggerBody()?['file']?['contentBytes']))
```

**If yes** → `Terminate`, status `Failed`, message naming what was missing.

### 3b. SharePoint — `Stage workbook` (Create file)

| Field | Value |
|---|---|
| Folder Path | `/SIM Imports/Staging` |
| File Name | `variables('varFileName')` |
| File Content | `triggerBody()?['file']?['contentBytes']` |

Retry: Exponential, 4, `PT5S`.

Keep the upload. When someone disputes what the import did, the submitted bytes are the only
evidence.

---

## 4. Load the inventory

### 4a. SharePoint — `Get inventory` (Get items)

| Field | Value |
|---|---|
| List | Global SIM Inventory |
| Filter Query | `SIM_Country eq '@{replace(triggerBody()?['text'],'''','''''')}'` |
| Top Count | `5000` |
| Limit Columns by View | a view containing only ID, ICC_ID, PhoneNr |

Settings → **Pagination On**, Threshold `100000`. Retry: Exponential, 4.

`Limit Columns by View` is not cosmetic — the default returns every field including system
ones, which on 60,000 items is the difference between a few MB and a few hundred.

### 4b. Data Operation — `Shape inventory` (Select)

**From:** `body('Get_inventory')?['value']`
**Map** (text mode):

```json
{
  "id": @{item()?['ID']},
  "iccId": "@{item()?['ICC_ID']}",
  "phoneNr": "@{item()?['PhoneNr']}"
}
```

`id` unquoted so it stays numeric; `iccId` quoted so 18–22 digits survive as a string.

One `Select` maps the whole array. Never use `Apply to each` for shaping.

---

## 5. Do until — `Pages`

**Condition:** `variables('varHasMore')` is equal to `false`
**Limits:** Count `100`, Timeout `PT2H`

The count limit is the fail-safe against a bug that never flips `hasMore`.

### 5a. Excel Online (Business) — `Validate and match` (Run script)

| Field | Value |
|---|---|
| File | the staged file (use the ID from `Stage workbook`) |
| Script | `ValidateAndMatchImport` |
| payload | below |

```
string(json(concat('{
 "tableName":"Table_query",
 "country":"', triggerBody()?['text'], '",
 "page":', variables('varPage'), ',
 "pageSize":2000,
 "reportDetail":"exceptions",
 "includeDecisions":false,
 "batchConfig":{
   "site":"https://deutschebank.sharepoint.com/sites/simri",
   "listGuid":"6b659861-abd0-4e45-b74e-63e3f69f2648",
   "entityType":"SP.Data.Global_x0020_SIM_x0020_InventoryListItem",
   "runId":"', variables('varRunId'), '",
   "batchSize":100,
   "fieldMap":', string(outputs('Compose_field_map')), '
 },
 "inventory":', string(body('Shape_inventory')), '
}')))
```

Put the field map in a `Compose` named `Compose field map` at the top of the flow:

```json
[
  {"excel":"FirstName","internal":"FirstName","type":"Text"},
  {"excel":"LastName","internal":"LastName","type":"Text"},
  {"excel":"PhoneNr","internal":"PhoneNr","type":"Text"},
  {"excel":"ICC_ID","internal":"ICC_ID","type":"Text"},
  {"excel":"StartDate","internal":"StartDate","type":"DateTime"},
  {"excel":"Legal entity","internal":"Legal_x0020_entity","type":"Text"},
  {"excel":"SIM Type","internal":"SIM_x0020_Type","type":"Text"}
]
```

Fill in all 20 from the §0 lookup. `type` drives conversion: `DateTime` → ISO 8601 (SharePoint
rejects `15-03-2026`), `Number`/`Currency` → numeric with blank as `null`, `Boolean` →
true/false, anything else → text. Never include `ID`; the script skips it anyway.

Retry: Exponential, 4, `PT10S`.

### 5b. Data Operation — `Parse result` (Parse JSON)

**Content:** `json(body('Validate_and_match')?['result'])`

The script returns a JSON string because Power Automate can't build a schema for dynamic field
keys. Generate the schema from a real sample after the first run rather than writing it by hand.

Returned shape:

```
counts            {created, updated, warning, skipped, failed}   this page only
skippedByCountry  [{country, rows}]
reportRows        failed + warning rows only (reportDetail: exceptions)
batches           [{boundary, body, rows}]  ready-to-POST payloads
hasMore           boolean
```

### 5c. Apply to each — `Send batches`

**From:** `body('Parse_result')?['batches']`
**Concurrency:** off (sequential)

This iterates over *batches*, not rows: 2,000 rows ÷ 100 per batch = at most 20 iterations,
regardless of file size. The 60,000 rows never enter a loop.

**SharePoint — `Send batch` (Send an HTTP request to SharePoint)**

| Field | Value |
|---|---|
| Site Address | `https://deutschebank.sharepoint.com/sites/simri` |
| Method | `POST` |
| Uri | `_api/$batch` |
| Headers | `Content-Type` : `multipart/mixed; boundary=@{items('Send_batches')?['boundary']}` |
| Body | `@{items('Send_batches')?['body']}` |

Retry: Exponential, 4, `PT10S`.

The action's Uri is relative; the `POST` lines inside the body are absolute. The script handles
the latter.

**⚠️ Scanning the response — do not skip this.**

`$batch` returns HTTP 200 even when every operation inside it failed. Power Automate shows the
action green. The per-operation status lives in the response body.

**How rows are matched to results: by position, not by boundary.** SharePoint generates its own
`--batchresponse_<guid>` and `--changesetresponse_<guid>` names; it does *not* echo the changeset
boundaries you sent, so the Excel row number cannot be recovered from them. What OData does
guarantee is that operation results come back **in request order**. That is exactly why the
script returns `rows` — the Excel row numbers for that batch, in the order they were written.

Four actions, all inside `Apply to each — batches`, after `Send batch`:

**1. `Compose status parts`**

```
skip(split(string(body('Send_batch')), 'HTTP/1.1 '), 1)
```

Splitting on `HTTP/1.1 ` gives one element per operation result, each beginning with its status
code. `skip(...,1)` drops the multipart preamble before the first result.

**2. `Compose row results`** — a `Select`

**From:** `range(0, length(items('Send_batches')?['rows']))`
**Map** (text mode):

```json
{
  "excelRow": @{items('Send_batches')?['rows'][item()]},
  "status": "@{substring(outputs('Compose_status_parts')[item()], 0, 3)}"
}
```

`range()` gives the index positions, which index both arrays in step.

**3. `Condition — result count matches`**

```
equals(length(outputs('Compose_status_parts')), length(items('Send_batches')?['rows']))
```

**If no** → the positional mapping is unsafe, so don't guess. Append every row in
`items('Send_batches')?['rows']` to `varReportRows` as failed with the message
"Batch response could not be matched to rows — verify these manually", and record the raw
response. Silently mismatching results to rows would put a wrong outcome against a row, which is
worse than admitting it's unknown.

**4. `Filter failed operations`**

**From:** `body('Compose_row_results')`
**Condition:** `startsWith(item()?['status'], '2')` **is equal to** `false`

Append each to `varReportRows` as `action: "failed"` with the status code and the error text
from the corresponding part.

Inspect one real response before building this. Confirm that `HTTP/1.1 ` appears exactly once
per operation — an error body that happens to contain the string would shift the alignment, and
the count guard in step 3 is what catches that.

### 5d. Accumulate

In order:

| Action | Value |
|---|---|
| `Compose merged rows` | `union(variables('varReportRows'), body('Parse_result')?['reportRows'])` |
| `Set varReportRows` | `outputs('Compose_merged_rows')` |
| `Increment varCreated` | `body('Parse_result')?['counts']?['created']` |
| `Increment varUpdated` | `body('Parse_result')?['counts']?['updated']` |
| `Increment varWarning` | `body('Parse_result')?['counts']?['warning']` |
| `Increment varSkipped` | `body('Parse_result')?['counts']?['skipped']` |
| `Increment varFailed` | `body('Parse_result')?['counts']?['failed']` |
| `Set varHasMore` | `body('Parse_result')?['hasMore']` |
| `Increment varPage` | `1` |

`union` merges in one action per page instead of one per row. It de-duplicates identical
objects, but every report row carries a unique `excelRow`, so nothing is ever dropped.

Because `reportDetail` is `exceptions`, only failed and warning rows accumulate. Created,
updated and skipped are counted but not listed — a 60,000-row import in `all` mode produces
~10 MB of JSON, past the 5 MB response limit and unreadable anyway. The totals still reconcile.

---

## 6. Build the report

**`Compose report meta`**

```json
{
  "fileName": "@{variables('varFileName')}",
  "country": "@{triggerBody()?['text']}",
  "runId": "@{variables('varRunId')}",
  "startedUtc": "@{triggerBody()?['headers']?['x-ms-workflow-run-start-time']}",
  "finishedUtc": "@{utcNow()}",
  "operator": "@{triggerBody()?['text_1']}",
  "listName": "Global SIM Inventory",
  "gate": "@{if(equals(variables('varFailed'),0),'OK','BLOCKED')}"
}
```

**`Compose report json`**

```
{ "meta": @{outputs('Compose_report_meta')}, "rows": @{variables('varReportRows')} }
```

**`Get report template`** — SharePoint Get file content → `Import_Report_Template.html`

**`Compose report html`**

```
replace(base64ToString(body('Get_report_template')?['$content']), '__REPORT_JSON__', string(outputs('Compose_report_json')))
```

This requires the template's mock data block to have been replaced with the bare token
`__REPORT_JSON__`, so the line reads `const REPORT = __REPORT_JSON__;`.

**`Save report`** — Create file in `/SIM Imports/Reports`, name
`replace(variables('varFileName'),'.xlsx','.html')`

---

## 7. Send the report

Office 365 Outlook — Send an email (V2)

| Field | Value |
|---|---|
| To | `triggerBody()?['text_1']` |
| Subject | `SIM import — @{triggerBody()?['text']} — @{variables('varFailed')} failed, @{variables('varCreated')} created, @{variables('varUpdated')} updated` |
| Body | short summary + link to the saved report |
| Attachment Name | `replace(variables('varFileName'),'.xlsx','.html')` |
| Attachment Content | `base64(outputs('Compose_report_html'))` |

Counts belong in the subject line — most recipients will never open the attachment.

Attach the HTML rather than pasting it into the body: email clients strip JavaScript, which
would leave a dead, unfilterable table.

---

## 8. `Create log item`

One item in an `Import Log` list: RunId, Country, ActionedBy, FileName, Started, Finished,
Created, Updated, Warning, Skipped, Failed, ReportUrl, Status = `Completed`.

Flow run history expires after 28 days on most plans. This list is what answers "who changed
this SIM and when" a year later.

---

## 9. Scope — Catch

Configure run after → **has failed**, **is skipped**, **has timed out** on `Scope - Main`.

1. `Send failure email` to ActionedBy and to you. Put `result('Scope_-_Main')` in the body — it
   returns the status and error of every action inside, which is what you'll debug from.
2. `Create log item` with Status = `Failed` and the error.
3. `Terminate`, status `Failed`.

Without the Terminate, a flow that failed halfway reports success.

---

## Matching rules

Applied per row, in this order. Rows failing any check are reported and never written.

1. **Blank row** → ignored silently (all of FirstName, ICC_ID, PhoneNr empty).
2. **`SIM_Country` ≠ the trigger's Country** → skipped, tallied by country.
3. **Validation** — PhoneNr, ICC_ID, IMEI, StartDate, Status, SIM Type, recomputed from raw
   data in TypeScript. The workbook's check columns are ignored entirely: they are formulas in
   a file the user controls and will report `OK` if overwritten or left unrecalculated.
4. **`ID` present in Excel:**
   - found in the inventory → **update** that item
   - not found → **failed**. Never falls through to create — a stale ID means the item was
     deleted or the row came from another country's export, and creating would duplicate it
   - not numeric, or the same ID on more than one row in the file → **failed**
5. **No `ID`:** match on `ICC_ID` + normalised `PhoneNr`
   - one match → **update**
   - no match → **create**
   - matches two SharePoint items, or the pair appears on two rows in the file → **failed**,
     ambiguous

Phone normalisation strips `+`, spaces, non-breaking spaces, `-`, `(`, `)`, `.` and converts a
leading `00`, so `+40712345678`, `0040 712 345 678` and `40712345678` all produce the same key.

Duplicate detection scans the whole ID / ICC_ID / PhoneNr columns on every call, so two rows
targeting the same item are caught even when they are 40,000 rows apart on different pages.

---

## Load profile at 60,000 rows

| Stage | Calls |
|---|---|
| Get inventory (paginated) | ~12 |
| Run script (2,000 rows/page) | 30 |
| `$batch` POSTs (100 ops each) | ~600 |
| **Total** | **~640** |

Per-item `Create item` / `Update item` would be 60,000 calls — against a limit of 600 per
connection per 60 seconds, that is 100 minutes of throttle budget before any latency.

Longest single `Apply to each`: 20 iterations.

---

## Test order

1. **3 rows**, one update / one create / one deliberately invalid. Inspect the raw `$batch`
   response by hand before trusting the parser.
2. **The same file again.** Must produce 3 updates, 0 creates. If it creates duplicates the
   matching is broken — this is the property you most need and the test most often skipped.
3. **Isolation:** a batch of 3 where the middle row is invalid. Rows 1 and 3 must still commit.
   If all three roll back, changesets aren't isolating and per-row reporting won't work.
4. **Edge rows:** stale ID, duplicate ICC+phone, another country, blank row.
5. **500 rows** to exercise paging.
6. Full volume.
