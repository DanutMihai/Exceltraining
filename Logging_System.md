# SIM Inventory Logs — logging system

List: **SIM Inventory Logs**. One item per import run.

---

## Why this list exists

Power Automate run history expires — 28 days on most plans. After that, the only record that an
import happened, who ran it, what it changed and whether anyone should worry is whatever you
wrote to SharePoint. A year from now the question will be "who changed this SIM and when", and
the flow run that did it will be long gone.

So the list is not a nice-to-have log. It is the only durable audit trail the process has.

---

## Three log points, not one

Logging only on success is the common mistake. It records the runs you least need to ask about
and misses every run you do.

```
Trigger
Initialize variables
Compose field map
  ┌──────────────────────────────────────────────┐
  │ LOG POINT 1 — Create item · Status: Running  │   ← before Scope - Main
  │ Set varLogItemId                             │
  └──────────────────────────────────────────────┘
Scope - Main
    Validate inputs · Stage workbook
    Get inventory · Shape inventory
    Do until - Pages  ( Run script → batches → counters )
    Build report · Save report
  ┌──────────────────────────────────────────────┐
  │ LOG POINT 2 — Update item · Completed        │   ← last action inside Scope - Main
  └──────────────────────────────────────────────┘
    Send success email
Scope - Catch   (run after: failed / skipped / timed out)
  ┌──────────────────────────────────────────────┐
  │ LOG POINT 3 — Update item · Failed           │   ← first action in the catch
  └──────────────────────────────────────────────┘
    Send failure email · Terminate
```

**Create at the start, update at the end.** The reason is the class of failure that reaches
neither terminal action:

- the flow times out or hits its run-duration limit
- an admin cancels the run
- the connection drops mid-run
- the catch scope itself fails

In all four cases a run that touched your list leaves no record at all if you only log at the
end. With the create-first pattern the item exists from second one, and simply stays at
`Running` — which is itself the signal. A view filtered to `Status = Running` and `Started`
older than two hours is your stuck-run monitor, and it costs nothing to maintain.

Log point 2 goes **inside** `Scope - Main`, as its last action, before the email. If it sat
outside, a failure in the email step would leave the log saying `Running` on a run that actually
succeeded.

---

## Columns

| Column | Type | Notes |
|---|---|---|
| **Title** | Single line of text | Required by default. Human-readable summary — see inputs below. |
| **RunId** | Single line of text | **Index this.** The join key to the flow run, the report and the staged workbook. |
| **Status** | Choice: `Running`, `Completed`, `Completed with rejections`, `Failed` | **Index this.** Default `Running`. Every view filters on it. |
| **Country** | Single line of text | Index if you have more than a handful of countries. |
| **ActionedBy_email** | Single line of text | Not a Person column — see below. |
| **Started** | Date and Time (include time) | SharePoint's built-in `Created` is roughly the *finish* time; pair them for duration. |
| **Finished** | Date and Time (include time) | Empty while `Running`. |
| **Created_Count** | Number, 0 decimals | |
| **Updated_Count** | Number, 0 decimals | |
| **Warning_Count** | Number, 0 decimals | |
| **Skipped_Count** | Number, 0 decimals | |
| **Failed_Count** | Number, 0 decimals | Numbers, not text — so you can filter `Failed_Count gt 0`, chart them and total a month. |
| **Rows_Processed** | Number, 0 decimals | Sum of the five. Saves a calculated column and reads better in views. |
| **Upload_Report** | Hyperlink | The generated HTML report. Empty on failure. |
| **SourceFile** | Hyperlink | The staged workbook — the evidence of what was actually submitted. |
| **FlowRun** | Hyperlink | Straight to the run history, while it still exists. |
| **WorkHistory** | Multiple lines of text, **plain text** | Human-readable summary plus the exception lines. |
| **ErrorMessage** | Multiple lines of text, plain text | Catch path only. Separate from WorkHistory so you can filter "is not empty". |

**Why `ActionedBy_email` is text, not a Person column.** Person columns need a numeric user ID
on write, so a text email means one less lookup per run. It also survives someone leaving the
company, which a Person column does not.

**Plain text, not enhanced rich text**, on both multi-line columns. Rich text stores HTML, which
makes the values unreadable in exports and in the API.

**Set list versioning off.** These items are written twice each by design; versioning doubles
the storage for no audit value you don't already have in the fields.

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
| FlowRun Url | `@{outputs('Compose_Flow_Identity')}` |
| FlowRun Description | `Flow run` |
| WorkHistory | `Run started. Awaiting completion.` |

Everything else: leave empty.

Immediately after, **Set variable** `varLogItemId` (Integer) = `body('Create_log_item')?['ID']`.
Add it to the §2 variable list.

Retry: Exponential, 4. If this action fails the run should still proceed — set
**Configure run after** on the next action to include *has failed*, so a logging problem never
blocks an import.

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
| Upload_Report Url | `@{body('Save_report')?['{Link}']}` |
| Upload_Report Description | `Import report` |
| SourceFile Url | `@{coalesce(body('Stage_workbook')?['{Link}'],'')}` |
| SourceFile Description | `@{variables('varFileName')}` |
| FlowRun Url | `@{outputs('Compose_Flow_Identity')}` |
| FlowRun Description | `Flow run` |

⚠️ **Repopulate every field, including the ones set at log point 1.** SharePoint's `Update item`
action sends the whole item — a field left blank in the action is written as blank, wiping what
was there. Country, ActionedBy_email, Started and RunId are all repeated above for that reason,
not by accident.

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
@{join(select(take(variables('varReportRows'),200), concat(item()?['action'], ' · row ', string(item()?['excelRow']), ' · ', item()?['message'])), decodeUriComponent('%0D%0A'))}
```

`decodeUriComponent('%0D%0A')` is how you produce a line break inside an expression — Power
Automate rejects a literal newline there.

`take(...,200)` caps the list. `varReportRows` holds only failed and warning rows, so it is
normally short, but a badly broken file could produce thousands and a multi-line text column
tops out at 63,999 characters. The full detail is in the report.

---

## Log point 3 — Update item (first action in `Scope - Catch`)

**Id:** `@{variables('varLogItemId')}`

Same fields as log point 2, with these differences:

| Field | Value |
|---|---|
| Title | `FAILED · @{formatDateTime(variables('varStartedUtc'),'yyyy-MM-dd HH:mm')} · @{triggerBody()?['text']}` |
| Status | `Failed` |
| Finished | `@{utcNow()}` |
| Upload_Report Url | leave empty — no report was produced |
| ErrorMessage | `@{string(result('Scope_-_Main'))}` |

The counters keep whatever they reached before the failure. That is the point: they tell you how
far the run got.

`string(result('Scope_-_Main'))` returns the status and error of every action in the scope.
Verbose, which is wrong for an email and right for a log — it is the whole picture six weeks
later when the run history has expired.

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

That NOTE earns its place. "Failed" reads like nothing happened; in this design it usually means
several batches did commit. Someone reading this in six months needs to know that before they
try to "clean up".

---

## Making the logging itself fail-safe

An error handler that can fail is not an error handler.

**Guard every reference to an action that might not have run.** If the flow fails before
`Stage workbook`, `body('Stage_workbook')?['{Link}']` has no value and the catch's Update item
fails too — taking your failure email with it. Wrap those:

```
@{coalesce(body('Stage_workbook')?['{Link}'],'')}
```

**Let logging fail without killing the import.** On log point 1, set the following action's
**Configure run after** to include *has failed* and *is skipped*. A SharePoint hiccup writing a
log row should never stop an import that was otherwise fine.

**`varLogItemId` may be 0.** If log point 1 failed, both Update actions target item 0 and error.
Guard them with a Condition on `greater(variables('varLogItemId'),0)`, or accept that the
failure email still sends because it runs after.

---

## Views worth creating

| View | Filter | Use |
|---|---|---|
| **Needs attention** | `Failed_Count > 0` OR `Status = Failed` | The daily check. |
| **Stuck runs** | `Status = Running` AND `Started < [Today]-1` | Runs that reached neither terminal action — timeouts, cancellations, a broken catch. Nothing else surfaces these. |
| **Last 30 days** | `Started >= [Today]-30`, sorted descending | The default. |
| **By country** | grouped on Country | Volume and error-rate patterns per market. |

Sort the default view by `Started` descending, and show Status, Country, Rows_Processed,
Failed_Count, ActionedBy_email, Upload_Report. Keep WorkHistory out of list views — it is
several hundred characters and destroys the layout. It reads fine in the item form.

---

## Retention

Nothing here expires on its own. At a few runs a day the list stays small for years, so the
simplest policy is to keep everything. If it ever needs trimming, a scheduled flow deleting
`Status = Completed` items older than two years leaves every failure and every rejection intact —
those are the ones anyone ever comes back to.

Set list permissions to read-only for users, write for the flow's connection identity. A log
anyone can edit answers no questions.
