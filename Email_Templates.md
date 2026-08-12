# Import emails — success and failure

Paste into **Send an email (V2)**. Click the `</>` (code view) toggle on the Body field first,
or Power Automate escapes the HTML and mails your recipients raw markup.

**Why these look old-fashioned under the hood.** Outlook on Windows renders with the Word engine.
No flexbox, no grid, no CSS variables, no `<style>` blocks you can rely on, and border-radius is
ignored. So: layout tables, inline styles, 600px fixed width. Everything degrades to a plain
readable block rather than collapsing. The colours are the same semantics as the workbook and
the report.

---

## 1. Success email

### Subject

```
SIM import @{if(equals(variables('varFailed'),0),'complete','— action needed')} · @{triggerBody()?['text']} · @{variables('varCreated')} created, @{variables('varUpdated')} updated@{if(equals(variables('varFailed'),0),'',concat(', ',variables('varFailed'),' rejected'))}
```

Renders as either:

- `SIM import complete · Romania · 3 created, 12 updated`
- `SIM import — action needed · Romania · 3 created, 12 updated, 4 rejected`

Counts in the subject line because most people never open the attachment. No emoji — corporate
filters treat them inconsistently and they age badly. Add `✅`/`⚠️` at the front if you disagree.

### Body

```html
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f4f6f9;margin:0;padding:24px 0;">
<tr><td align="center">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:600px;max-width:600px;background:#ffffff;border:1px solid #e3e8ef;">

  <tr><td style="height:4px;line-height:4px;font-size:0;background:@{if(equals(variables('varFailed'),0),'#4caf72','#e0a63a')};">&nbsp;</td></tr>

  <tr><td style="padding:26px 30px 4px;font-family:Segoe UI,Arial,sans-serif;">
    <div style="font-size:11px;letter-spacing:.8px;text-transform:uppercase;color:#8a94a6;font-weight:700;">Global SIM Inventory</div>
    <div style="font-size:21px;font-weight:600;color:#1a2233;padding-top:6px;">@{if(equals(variables('varFailed'),0),'Import complete','Import completed with rejections')}</div>
    <div style="font-size:14px;color:#5b6779;padding-top:8px;line-height:1.5;">@{triggerBody()?['text']} · @{add(add(add(add(variables('varCreated'),variables('varUpdated')),variables('varWarning')),variables('varSkipped')),variables('varFailed'))} rows processed</div>
  </td></tr>

  <tr><td style="padding:20px 30px 4px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="font-family:Segoe UI,Arial,sans-serif;">
      <tr>
        <td width="20%" align="center" style="padding:12px 4px;background:#f7f9fc;border:1px solid #e3e8ef;">
          <div style="font-size:22px;font-weight:700;color:#006100;">@{variables('varCreated')}</div>
          <div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#8a94a6;font-weight:700;padding-top:2px;">Created</div></td>
        <td width="20%" align="center" style="padding:12px 4px;background:#f7f9fc;border:1px solid #e3e8ef;">
          <div style="font-size:22px;font-weight:700;color:#1f4e79;">@{variables('varUpdated')}</div>
          <div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#8a94a6;font-weight:700;padding-top:2px;">Updated</div></td>
        <td width="20%" align="center" style="padding:12px 4px;background:#f7f9fc;border:1px solid #e3e8ef;">
          <div style="font-size:22px;font-weight:700;color:#9c5700;">@{variables('varWarning')}</div>
          <div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#8a94a6;font-weight:700;padding-top:2px;">Warning</div></td>
        <td width="20%" align="center" style="padding:12px 4px;background:#f7f9fc;border:1px solid #e3e8ef;">
          <div style="font-size:22px;font-weight:700;color:#5b6779;">@{variables('varSkipped')}</div>
          <div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#8a94a6;font-weight:700;padding-top:2px;">Skipped</div></td>
        <td width="20%" align="center" style="padding:12px 4px;background:@{if(equals(variables('varFailed'),0),'#f7f9fc','#ffc7ce')};border:1px solid @{if(equals(variables('varFailed'),0),'#e3e8ef','#f3adb6')};">
          <div style="font-size:22px;font-weight:700;color:#9c0006;">@{variables('varFailed')}</div>
          <div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#8a94a6;font-weight:700;padding-top:2px;">Failed</div></td>
      </tr>
    </table>
  </td></tr>

  <tr><td style="padding:18px 30px 0;font-family:Segoe UI,Arial,sans-serif;font-size:13.5px;color:#5b6779;line-height:1.6;">
    @{if(equals(variables('varFailed'),0),
      'Every row was written. Nothing further is needed.',
      concat('<b style="color:#9c0006;">', variables('varFailed'), ' row(s) were rejected and not written.</b> Everything else was. Open the attached report, fix those rows in the workbook, and upload it again — rows that already landed will update rather than duplicate.'))}
  </td></tr>

  <tr><td style="padding:20px 30px 0;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>
      <td style="background:#1a2233;padding:11px 22px;">
        <a href="@{body('Save_report')?['{Link}']}" style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;font-weight:600;color:#ffffff;text-decoration:none;display:inline-block;">Open the full report</a>
      </td>
    </tr></table>
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:12px;color:#8a94a6;padding-top:10px;">Also attached to this email. Open it in a browser — filtering and search need JavaScript, which email clients strip.</div>
  </td></tr>

  <tr><td style="padding:22px 30px 0;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="font-family:Segoe UI,Arial,sans-serif;font-size:12.5px;border-top:1px solid #eef1f6;">
      <tr><td style="padding:12px 0 3px;color:#8a94a6;width:120px;">Source file</td><td style="padding:12px 0 3px;color:#1a2233;">@{variables('varFileName')}</td></tr>
      <tr><td style="padding:3px 0;color:#8a94a6;">Run by</td><td style="padding:3px 0;color:#1a2233;">@{triggerBody()?['text_1']}</td></tr>
      <tr><td style="padding:3px 0;color:#8a94a6;">Started</td><td style="padding:3px 0;color:#1a2233;">@{formatDateTime(variables('varStartedUtc'),'dd-MM-yyyy HH:mm')} UTC</td></tr>
      <tr><td style="padding:3px 0 14px;color:#8a94a6;">Run ID</td><td style="padding:3px 0 14px;"><a href="@{outputs('Compose_Flow_Identity')}" style="color:#1f4e79;">@{variables('varRunId')}</a></td></tr>
    </table>
  </td></tr>

  <tr><td style="padding:14px 30px 24px;font-family:Segoe UI,Arial,sans-serif;font-size:11.5px;color:#8a94a6;border-top:1px solid #eef1f6;line-height:1.6;">
    Automated message from the SIM Inventory import flow. Do not reply.
  </td></tr>

</table>
</td></tr></table>
```

---

## 2. Failure email

### Subject

```
SIM import FAILED · @{triggerBody()?['text']} · @{variables('varFileName')}
```

### Body

```html
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f4f6f9;margin:0;padding:24px 0;">
<tr><td align="center">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:600px;max-width:600px;background:#ffffff;border:1px solid #e3e8ef;">

  <tr><td style="height:4px;line-height:4px;font-size:0;background:#d4626f;">&nbsp;</td></tr>

  <tr><td style="padding:26px 30px 4px;font-family:Segoe UI,Arial,sans-serif;">
    <div style="font-size:11px;letter-spacing:.8px;text-transform:uppercase;color:#8a94a6;font-weight:700;">Global SIM Inventory</div>
    <div style="font-size:21px;font-weight:600;color:#9c0006;padding-top:6px;">Import failed</div>
    <div style="font-size:14px;color:#5b6779;padding-top:8px;line-height:1.5;">@{triggerBody()?['text']} · the run stopped before finishing</div>
  </td></tr>

  <tr><td style="padding:20px 30px 0;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#fff6f7;border:1px solid #f3adb6;">
      <tr><td style="padding:16px 18px;font-family:Segoe UI,Arial,sans-serif;font-size:13.5px;color:#1a2233;line-height:1.6;">
        <b style="color:#9c0006;">The list may have been partly updated.</b><br>
        Rows written before the failure were committed and are not rolled back. Counted so far:
        <b>@{variables('varCreated')} created</b>, <b>@{variables('varUpdated')} updated</b>.
        Re-uploading the same file is safe — those rows will match on their ID and update rather
        than duplicate.
      </td></tr>
    </table>
  </td></tr>

  <tr><td style="padding:20px 30px 0;font-family:Segoe UI,Arial,sans-serif;">
    <div style="font-size:11px;letter-spacing:.8px;text-transform:uppercase;color:#8a94a6;font-weight:700;padding-bottom:8px;">What failed</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f7f9fc;border:1px solid #e3e8ef;">
      <tr><td style="padding:14px 16px;font-family:Consolas,Menlo,monospace;font-size:11.5px;color:#5b6779;line-height:1.6;word-break:break-word;">
        @{result('Scope_-_Main')[0]?['error']?['message']}
      </td></tr>
    </table>
  </td></tr>

  <tr><td style="padding:20px 30px 0;font-family:Segoe UI,Arial,sans-serif;font-size:13.5px;color:#5b6779;line-height:1.7;">
    <div style="font-size:11px;letter-spacing:.8px;text-transform:uppercase;color:#8a94a6;font-weight:700;padding-bottom:8px;">Next steps</div>
    1. Open the run below and find the red action.<br>
    2. If it is the Run script step, the message names the problem — usually a missing field or a validation rule.<br>
    3. Fix it, then re-upload the same file. Nothing needs undoing first.
  </td></tr>

  <tr><td style="padding:20px 30px 0;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>
      <td style="background:#9c0006;padding:11px 22px;">
        <a href="@{outputs('Compose_Flow_Identity')}" style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;font-weight:600;color:#ffffff;text-decoration:none;display:inline-block;">Open the failed run</a>
      </td>
    </tr></table>
  </td></tr>

  <tr><td style="padding:22px 30px 0;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="font-family:Segoe UI,Arial,sans-serif;font-size:12.5px;border-top:1px solid #eef1f6;">
      <tr><td style="padding:12px 0 3px;color:#8a94a6;width:120px;">Source file</td><td style="padding:12px 0 3px;color:#1a2233;">@{variables('varFileName')}</td></tr>
      <tr><td style="padding:3px 0;color:#8a94a6;">Uploaded by</td><td style="padding:3px 0;color:#1a2233;">@{triggerBody()?['text_1']}</td></tr>
      <tr><td style="padding:3px 0;color:#8a94a6;">Started</td><td style="padding:3px 0;color:#1a2233;">@{formatDateTime(variables('varStartedUtc'),'dd-MM-yyyy HH:mm')} UTC</td></tr>
      <tr><td style="padding:3px 0 14px;color:#8a94a6;">Run ID</td><td style="padding:3px 0 14px;color:#1a2233;">@{variables('varRunId')}</td></tr>
    </table>
  </td></tr>

  <tr><td style="padding:14px 30px 24px;font-family:Segoe UI,Arial,sans-serif;font-size:11.5px;color:#8a94a6;border-top:1px solid #eef1f6;line-height:1.6;">
    Automated message from the SIM Inventory import flow. Do not reply.
  </td></tr>

</table>
</td></tr></table>
```

Send this one to **both** ActionedBy and yourself — the user can't act on most failures, and you
want to know before they tell you.

---

## Notes

**The partial-write warning is the important part of the failure email.** "Import failed" reads
like nothing happened, when in fact every batch committed before the failure is still there.
Someone who assumes a rollback will do something destructive to "clean up". The counters are
accurate up to the point of failure because they're incremented per page.

**`result('Scope_-_Main')[0]?['error']?['message']`** returns the first failed action's message.
For the full picture use `string(result('Scope_-_Main'))`, but it's verbose — better in the log
list than in an email.

**`body('Save_report')?['{Link}']`** is the SharePoint URL of the saved report. If your Create
file action is named differently, adjust. If you skip saving a copy, drop the button and rely on
the attachment.

**Test in Outlook desktop, not just the web client.** The web client renders these near
perfectly; the desktop Word engine is where layout breaks. If the KPI row wraps oddly there,
drop it to three cells (created / updated / failed) — five columns at 600px is tight.
