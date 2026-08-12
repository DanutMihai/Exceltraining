# SIM Inventory Logs — logging system

List: **SIM Inventory Logs**. One item per import run.

---

## Why this list exists

Power Automate run history expires — 28 days on most plans. After that, the only record that an
import happened, who ran it, what it changed and whether anyone should worry is whatever you
wrote to SharePoint. A year from now the question will be "who changed this SIM and when", and
the run that did it will be long gone.

The list is not a nice-to-have. It is the only durable audit trail the process has.

---

## Three log points, not one

```
Trigger
Initialize variables
Compose field map
Compose Flow Identity
  ┌──────────────────────────────────────────────┐
  │ LOG POINT 1 — Create item · Status: Running  │   ← before Scope - Main
  │ Set varLogItemId                             │
  └──────────────────────────────────────────────┘
Scope - Main
    Validate inputs · Stage workbook
    Get inventory · Shape inventory
    Do until - Pages   ( Run script → batches → counters )
    Build report · Save report
  ┌──────────────────────────────────────────────┐
  │ LOG POINT 2 — Update item · Completed        │   ← last action INSIDE Scope - Main
  └──────────────────────────────────────────────┘
    Send success email
Scope - Catch   (run after: failed / skipped / timed out)
  ┌──────────────────────────────────────────────┐
  │ LOG POINT 3 — Update item · Failed           │   ← first action in the catch
  └──────────────────────────────────────────────┘
    Send failure email · Terminate
```

**Create at the start, update at the end.** Four kinds of failure reach neither terminal action:
a timeout or run-duration limit, an admin cancelling the run, a dropped connection, or the catch
scope itself failing. With end-only logging those runs leave no record — and they are exactly
the ones worth investigating. Created first, the item exists from second one and simply stays at
`Running`, which a stuck-runs view then surfaces.

Log point 2 sits **inside** the Scope, before the email. Outside it, a failure sending the email
would leave the log claiming `Running` on a run that actually succeeded.

---

## Columns

| Column | Type | Notes |
|---|---|---|
| **Title** | Single line of text | Required by default. |
| **RunId** | Single line of text | **Index this.** Join key to the run, the report and the staged workbook. |
| **Status** | Choice: `Running`, `Completed`, `Completed with rejections`, `Failed` | **Index this.** Default `Running`. |
| **Country** | Single line of text | Index it if you have more than a handful. |
| **ActionedBy_email** | Single line of text | Text, not Person — see below. |
| **Started** | Date and Time (include time) | |
| **Finished** | Date and Time (include time) | Empty while `Running`. |
| **Created_Count** | Number, 0 decimals | |
| **Updated_Count** | Number, 0 decimals | |
| **Warning_Count** | Number, 0 decimals | |
| **Skipped_Count** | Number, 0 decimals | |
| **Failed_Count** | Number, 0 decimals | Numbers, not text — so `Failed_Count gt 0` works as a filter. |
| **Rows_Processed** | Number, 0 decimals | |
| **Upload_Report** | Hyperlink | The generated report. Empty on failure. |
| **SourceFile** | Hyperlink | The staged workbook — evidence of what was submitted. |
| **FlowRun** | Hyperlink | The run history, while it exists. |
| **WorkHistory** | Multiple lines, **plain text** | Summary plus exception lines. |
| **ErrorMessage** | Multiple lines, plain text | Catch path only. Separate so you can filter "is not empty". |

**Text, not a Person column,** for ActionedBy_email: Person columns need a numeric user ID on
write, and a text email survives someone leaving the company.

**Plain text, not enhanced rich text** on both multi-line columns — rich text stores HTML, which
makes exports and API reads unreadable.

**Versioning off.** Every item is written twice by design; versioning doubles storage for no
audit value the fields don't already carry.

---

## ⚠️ The two URLs — read this before building

`{Link}` is a property of the **list item** schema, returned by `Get items` and
`Get file properties`. **`Create file` returns a file schema** — `Id`, `Name`, `Path`,
`LastModified`, `Size`, `ETag`, `FileLocator` — and I would not rely on `{Link}` being in it.

So don't reference `body('Save_report')?['{Link}']`. Build the URLs deterministically instead:

**Upload_Report Url**

```
@{concat('https://deutschebank.sharepoint.com/sites/simri/SIM%20Imports/Reports/', replace(replace(variables('varFileName'),'.xlsx','.html'),' ','%20'))}
```

**SourceFile Url**

```
@{concat('https://deutschebank.sharepoint.com/sites/simri/SIM%20Imports/Staging/', replace(variables('varFileName'),' ','%20'))}
```

Both are built from `varFileName`, which you control, so they cannot drift with connector output
shapes. `varFileName` has no spaces today — the `replace` is insurance in case the country name
ever gains one ("United Arab Emirates" would).

**Confirm the folder segments** match your libraries exactly, including case. If your library is
a document library at the site root, the path is `/SIM%20Imports/Reports/`; if it sits under
`/Shared%20Documents/`, include that too. Open one saved report in a browser and copy the URL —
that is the ground truth.

**Alternative, if you'd rather not hardcode:** add `Get file properties` after `Save report`,
pass it `body('Save_report')?['ItemId']`, and use `body('Get_file_properties')?['{Link}']`. Two
extra API calls per run, and it survives a library move. The composed version is simpler and I'd
start there.

---

## Log point 1 — Create item (before `Scope - Main`)

| Field | Value |
|---|---|
| Title | `RUNNING · @{formatDateTime(variables('varStartedUtc'),'yyyy-MM-dd HH:mm')} · @{triggerBody()?['text']}` |
| RunId | `@{variables('varRunId')}` |
| Status | `Running` |
| Country | `@{triggerBody()?['text']}` |
| ActionedBy_email | `@{triggerBody()?['text_1']}` |
| Started | `@{variables('varStartedUtc')}` |
| Finished | *leave empty* |
| Created_Count … Rows_Processed | `0` |
| Upload_Report Url / Description | *leave empty* |
| SourceFile Url | `@{concat('https://deutschebank.sharepoint.com/sites/simri/SIM%20Imports/Staging/', replace(variables('varFileName'),' ','%20'))}` |
| SourceFile Description | `@{variables('varFileName')}` |
| FlowRun Url | `@{outputs('Compose_Flow_Identity')}` |
| FlowRun Description | `Flow run` |
| WorkHistory | `Run started at @{formatDateTime(variables('varStartedUtc'),'dd-MM-yyyy HH:mm')} UTC. Awaiting completion.` |
| ErrorMessage | *leave empty* |

Counts as `0` rather than empty: a `Running` row that shows blanks looks like a data problem,
and a Number column left empty sorts unpredictably.

`SourceFile` is safe to set here — the URL is derived from `varFileName`, which exists before
`Stage workbook` runs. If staging then fails, the link 404s, which is itself accurate: no file
was staged.

**Immediately after:** `Set variable` `varLogItemId` (Integer) = `body('Create_log_item')?['ID']`.

Retry: Exponential, 4. On the **next** action set *Configure run after* to include **has failed**
and **is skipped**, so a logging hiccup never blocks an import.

---

## Exceptions text — no `Select` action needed

The Select action's map is validated as JSON, so it cannot emit a bare string. Rather than work
around that, the script returns the text ready-made.

**Add a variable** (root level, §2): `varExceptions`, String, initial value empty.

**In the page loop (§5d), after `Set varHasMore`**, add:

**`Append to string variable`** — Name `varExceptions`, Value:

```
@{body('Parse_result')?['exceptionsText']}@{if(empty(body('Parse_result')?['exceptionsText']),'',decodeUriComponent('%0D%0A'))}
```

One append per page — 30 at 60,000 rows, not one per row. The trailing newline is only added
when the page actually produced exceptions, so clean pages don't pad the value with blank lines.

`exceptionsText` contains only **failed** and **warning** rows, whatever `reportDetail` is set
to, capped at 200 lines per page with a "… and N more" notice. A log column does not want 60,000
lines of "Item updated."

---

## Log point 2 — Update item (last action inside `Scope - Main`)

**Id:** `@{variables('varLogItemId')}`

| Field | Value |
|---|---|
| Title | `@{formatDateTime(variables('varStartedUtc'),'yyyy-MM-dd HH:mm')} · @{triggerBody()?['text']} · @{variables('varCreated')}C / @{variables('varUpdated')}U / @{variables('varFailed')}F` |
| RunId | `@{variables('varRunId')}` |
| Status | `@{if(equals(variables('varFailed'),0),'Completed','Completed with rejections')}` |
| Country | `@{triggerBody()?['text']}` |
| ActionedBy_email | `@{triggerBody()?['text_1']}` |
| Started | `@{variables('varStartedUtc')}` |
| Finished | `@{utcNow()}` |
| Created_Count | `@{variables('varCreated')}` |
| Updated_Count | `@{variables('varUpdated')}` |
| Warning_Count | `@{variables('varWarning')}` |
| Skipped_Count | `@{variables('varSkipped')}` |
| Failed_Count | `@{variables('varFailed')}` |
| Rows_Processed | `@{add(add(add(add(variables('varCreated'),variables('varUpdated')),variables('varWarning')),variables('varSkipped')),variables('varFailed'))}` |
| Upload_Report Url | `@{concat('https://deutschebank.sharepoint.com/sites/simri/SIM%20Imports/Reports/', replace(replace(variables('varFileName'),'.xlsx','.html'),' ','%20'))}` |
| Upload_Report Description | `Import report` |
| SourceFile Url | `@{concat('https://deutschebank.sharepoint.com/sites/simri/SIM%20Imports/Staging/', replace(variables('varFileName'),' ','%20'))}` |
| SourceFile Description | `@{variables('varFileName')}` |
| FlowRun Url | `@{outputs('Compose_Flow_Identity')}` |
| FlowRun Description | `Flow run` |
| ErrorMessage | *leave empty* |

⚠️ **Every field is repopulated, including those set at log point 1.** SharePoint's `Update item`
writes the whole item — a field left blank in the action is written as blank, wiping what was
there. The repetition is deliberate.

### WorkHistory

```
Status: @{if(equals(variables('varFailed'),0),'Completed','Completed with rejections')}
Country: @{triggerBody()?['text']}
Rows: @{add(add(add(add(variables('varCreated'),variables('varUpdated')),variables('varWarning')),variables('varSkipped')),variables('varFailed'))} processed — @{variables('varCreated')} created, @{variables('varUpdated')} updated, @{variables('varWarning')} warning, @{variables('varSkipped')} skipped, @{variables('varFailed')} failed
Started: @{formatDateTime(variables('varStartedUtc'),'dd-MM-yyyy HH:mm')} UTC
Finished: @{formatDateTime(utcNow(),'dd-MM-yyyy HH:mm')} UTC
Source: @{variables('varFileName')}
Run: @{variables('varRunId')}

Exceptions:
@{if(greater(length(variables('varExceptions')),60000),concat(substring(variables('varExceptions'),0,60000),decodeUriComponent('%0D%0A'),'… truncated, see the full report.'),variables('varExceptions'))}
```

`decodeUriComponent('%0D%0A')` produces a line break — Power Automate rejects a literal newline
inside an expression.

The `if(greater(length(...),60000),...)` guard is the last line of defence: the script caps at
200 lines per page, but 30 pages of exceptions could still pass SharePoint's 63,999-character
limit on a multi-line text column, and exceeding it fails the whole Update item.

---

## Log point 3 — Update item (first action in `Scope - Catch`)

**Id:** `@{variables('varLogItemId')}`

Written out in full rather than "same as above, except" — the whole point is that no field is
left blank by accident.

| Field | Value |
|---|---|
| Title | `FAILED · @{formatDateTime(variables('varStartedUtc'),'yyyy-MM-dd HH:mm')} · @{triggerBody()?['text']}` |
| RunId | `@{variables('varRunId')}` |
| Status | `Failed` |
| Country | `@{triggerBody()?['text']}` |
| ActionedBy_email | `@{triggerBody()?['text_1']}` |
| Started | `@{variables('varStartedUtc')}` |
| Finished | `@{utcNow()}` |
| Created_Count | `@{variables('varCreated')}` |
| Updated_Count | `@{variables('varUpdated')}` |
| Warning_Count | `@{variables('varWarning')}` |
| Skipped_Count | `@{variables('varSkipped')}` |
| Failed_Count | `@{variables('varFailed')}` |
| Rows_Processed | `@{add(add(add(add(variables('varCreated'),variables('varUpdated')),variables('varWarning')),variables('varSkipped')),variables('varFailed'))}` |
| Upload_Report Url | *leave empty — no report was produced* |
| Upload_Report Description | *leave empty* |
| SourceFile Url | `@{concat('https://deutschebank.sharepoint.com/sites/simri/SIM%20Imports/Staging/', replace(variables('varFileName'),' ','%20'))}` |
| SourceFile Description | `@{variables('varFileName')}` |
| FlowRun Url | `@{outputs('Compose_Flow_Identity')}` |
| FlowRun Description | `Flow run` |
| ErrorMessage | `@{string(result('Scope_-_Main'))}` |

The counters keep whatever they reached before the failure — that is how far the run got, and it
is the number that tells you whether anything was written.

`string(result('Scope_-_Main'))` returns the status and error of every action in the scope.
Verbose, which is wrong for an email and right for a log.

**Do not reference any action inside `Scope - Main` here** — if the flow failed before it ran,
the reference is unresolvable and the catch's Update item fails too. `varExceptions` is a
variable, so it is always readable; it simply holds whatever accumulated before the failure.

### WorkHistory (catch)

```
Status: FAILED
Country: @{triggerBody()?['text']}
Written before the failure: @{variables('varCreated')} created, @{variables('varUpdated')} updated
Started: @{formatDateTime(variables('varStartedUtc'),'dd-MM-yyyy HH:mm')} UTC
Failed at: @{formatDateTime(utcNow(),'dd-MM-yyyy HH:mm')} UTC
Source: @{variables('varFileName')}
Run: @{variables('varRunId')}

NOTE: this run was NOT rolled back. The rows counted above were committed to the list.
Re-uploading the same file is safe — they will match on their ID and update rather than duplicate.

Error:
@{result('Scope_-_Main')[0]?['error']?['message']}
```

That NOTE earns its place. "Failed" reads like nothing happened; here it usually means several
batches did commit. Someone reading this in six months needs to know before they "clean up".

---

## Verify these against your first run

Everything above is derived from values you control, except four references to other actions.
Open the first run in history, expand each action, and check the **Outputs** panel:

| Reference | Check | If it's wrong |
|---|---|---|
| `body('Create_log_item')?['ID']` | `Create item` outputs contain `ID` (a number) | Use `?['Id']` — casing varies by connector version |
| `outputs('Compose_Flow_Identity')` | The Compose output is a bare URL string, not an object | If it is an object, index into it: `?['url']` |
| `triggerBody()?['text_1']` | Equals the ActionedBy email, not the Country | Swap for `?['text']` / `?['text_2']` — the suffix depends on input order |
| `result('Scope_-_Main')` | Returns an array in the catch path | If the Scope is named differently, match it exactly with spaces as underscores |

And open the two composed URLs in a browser. A 404 means the folder segment is wrong — copy a
real report's URL and adjust the literal.

**Field-by-field sanity checks worth doing once:**

- **Started and Finished show a real time**, not blank. Both take ISO 8601, which `utcNow()` and `varStartedUtc` already produce.
- **The five counts sum to Rows_Processed.** If they don't, a counter increment is missing from §5d.
- **Status is one of the four choice values exactly.** A typo creates a new value silently in some list configurations and breaks every view filter.
- **Upload_Report is empty on the failure path** and populated on success.
- **WorkHistory contains the Exceptions block** on a run that had rejections, and an empty one when clean.

---

## Making the logging fail-safe

**Guard `varLogItemId`.** If log point 1 failed, it is `0` and both Update actions error against
item 0. Wrap each in a Condition on `greater(variables('varLogItemId'),0)`, or accept that the
failure email still sends because it runs after.

**Let logging fail without killing the import.** On the action following log point 1, set
*Configure run after* to include **has failed** and **is skipped**.

**Don't reference actions that may not have run.** In the catch path, anything inside
`Scope - Main` may never have executed. The values above are all variables or literals for
exactly that reason.

An error handler that can fail is not an error handler.

---

## Views worth creating

| View | Filter | Use |
|---|---|---|
| **Needs attention** | `Failed_Count > 0` OR `Status = Failed` | The daily check. |
| **Stuck runs** | `Status = Running` AND `Started < [Today]-1` | Runs that reached neither terminal action. Nothing else surfaces these. |
| **Last 30 days** | `Started >= [Today]-30`, sorted descending | The default. |
| **By country** | grouped on Country | Volume and error-rate patterns per market. |

Default view sorted by `Started` descending, showing Status, Country, Rows_Processed,
Failed_Count, ActionedBy_email, Upload_Report. Keep WorkHistory out of list views — several
hundred characters destroys the layout. It reads fine in the item form.

---

## Retention and permissions

Nothing expires on its own. At a few runs a day the list stays small for years, so keeping
everything is the simplest policy. If it ever needs trimming, a scheduled flow deleting
`Status = Completed` items older than two years leaves every failure and rejection intact —
those are the ones anyone comes back to.

Set permissions to read-only for users, write for the flow's connection identity. A log anyone
can edit answers no questions.
