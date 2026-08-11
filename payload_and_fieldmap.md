# `Compose field map` and the Run script `payload`

Both ready to paste. Based on your XML field dump, with `Title` remapped and the three
new columns kept.

---

## 1. `Compose field map` — Inputs

```json
[
  {"excel":"LastName","internal":"Title","type":"Text"},
  {"excel":"FirstName","internal":"FirstName","type":"Text"},
  {"excel":"LastName","internal":"LastName","type":"Text"},
  {"excel":"PhoneNr","internal":"PhoneNr","type":"Text"},
  {"excel":"Plan","internal":"Plan","type":"Text"},
  {"excel":"ICC_ID","internal":"ICC_ID","type":"Text"},
  {"excel":"GD_ID","internal":"GD_ID","type":"Text"},
  {"excel":"StartDate","internal":"StartDate","type":"Text"},
  {"excel":"SIM_Country","internal":"SIM_Country","type":"Text"},
  {"excel":"Provider","internal":"Provider","type":"Text"},
  {"excel":"CostCenter","internal":"CostCenter","type":"Text"},
  {"excel":"UBR_Code","internal":"UBR_Code","type":"Text"},
  {"excel":"Legal entity","internal":"Legalentity","type":"Text"},
  {"excel":"IsVR","internal":"IsVR","type":"Boolean"},
  {"excel":"BAN","internal":"BAN","type":"Text"},
  {"excel":"SOC","internal":"SOC","type":"Text"},
  {"excel":"Status","internal":"Status","type":"Choice"},
  {"excel":"SIM Type","internal":"SIMType","type":"Text"},
  {"excel":"Email","internal":"Email","type":"Text"},
  {"excel":"IMEI","internal":"IMEI","type":"Text"},
  {"excel":"Location","internal":"Location","type":"Text"},
  {"excel":"SIMRI_REQ_ID","internal":"SIMRI_REQ_ID","type":"Text"},
  {"excel":"NGCC_SNOW_TICKET_ID","internal":"NGCC_SNOW_TICKET_ID","type":"Text"}
]
```

Notes on the three changes from your version:

- **`LastName` appears twice**, once as `Title` and once as `LastName`. That's deliberate, not a
  mistake — both SharePoint fields get the surname. Delete the first line only if `Title` is
  `Required: false` on the list.
- **`Location`, `SIMRI_REQ_ID`, `NGCC_SNOW_TICKET_ID` stay.** They aren't `Table_query` headers,
  so the script skips them, and `MERGE` leaves the existing SharePoint values untouched. Keeping
  them here means they'll start working automatically if you ever add those columns to the
  export.
- **`StartDate` is `Text`** because that's what your field dump said. If the SharePoint column is
  actually a Date column, change it to `DateTime` — the script then converts `15-03-2026` to ISO
  8601, which is what a Date column requires.

---

## 2. Run script → `payload`

One line. Paste as an expression.

```
string(json(concat('{"tableName":"Table_query","country":"',triggerBody()?['text'],'","page":',variables('varPage'),',"pageSize":2000,"reportDetail":"exceptions","includeDecisions":false,"batchConfig":{"site":"https://deutschebank.sharepoint.com/sites/simri","listGuid":"6b659861-abd0-4e45-b74e-63e3f69f2648","entityType":"SP.Data.Global_x0020_SIM_x0020_InventoryListItem","runId":"',variables('varRunId'),'","batchSize":100,"fieldMap":',string(outputs('Compose_field_map')),'},"inventory":',string(body('Shape_inventory')),'}')))
```

The `json(...)` wrapper is a validation step — if the concatenation produces malformed JSON it
fails here with a parse error, rather than inside the script with something cryptic. `string(...)`
turns it back into the string the parameter expects.

### What each part resolves to

| Field | Source | Note |
|---|---|---|
| `country` | trigger | check the `text` suffix matches your input order |
| `page` | `varPage` | must be the variable, not `0`, or the loop never advances |
| `pageSize` | `2000` | rows per script call |
| `reportDetail` | `exceptions` | `all` while testing 205 rows, `exceptions` at volume |
| `includeDecisions` | `false` | `true` to see decisions alongside batches while debugging |
| `runId` | `varRunId` | becomes the batch and changeset boundaries |
| `batchSize` | `100` | operations per `$batch` POST |
| `fieldMap` | Compose | above |
| `inventory` | `Shape_inventory` | the `{id, iccId, phoneNr}` array |

### If action names differ

`outputs('Compose_field_map')` and `body('Shape_inventory')` must match your actual action names
with spaces replaced by underscores. `Compose field map` → `Compose_field_map`.

---

## 3. First run

Set `reportDetail` to `all` and `includeDecisions` to `true` for the first attempt. You'll see
every row's outcome plus the decisions next to the batches, which makes a mismatch obvious.
Switch both back before you go near 60,000 rows.

Expected on your 205-row file: `batches` with one entry (5 decisions, under the 100 batch size),
`counts` summing to 5, and `hasMore: false`.
