# PowerApps feature banner — bulk inventory upload

For an **HTML text** control. Two constraints shaped it:

**No JavaScript.** PowerApps strips it entirely, so nothing is interactive. Dismissing the
banner is done with a PowerApps variable, not in the HTML — see the bottom.

**No media queries.** They key off the *viewport*, not the control, so an HTML text control that
shrinks with its container won't trigger them. Instead the layout uses `flex-wrap` with a
`flex-basis` on each card: the three steps sit side by side when there's room and stack when
there isn't, driven by the control's actual rendered width. That's genuinely container-responsive
without a single media query.

**Single quotes throughout.** PowerApps strings are double-quoted, and every `"` inside would
have to be doubled to `""`. Using `'` for attributes means you paste this between two double
quotes and nothing needs escaping.

---

## The HTML

```html
<div style='font-family:Segoe UI,system-ui,sans-serif;background:#ffffff;border:1px solid #e3e8ef;border-left:4px solid #1f4e79;border-radius:10px;padding:0;overflow:hidden;box-sizing:border-box;width:100%'>
  <div style='padding:16px 18px 14px'>

    <div style='display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-bottom:8px'>
      <span style='display:inline-block;background:#c6efce;color:#006100;font-size:10px;font-weight:700;letter-spacing:.8px;padding:3px 9px;border-radius:999px'>NEW</span>
      <span style='font-size:16px;font-weight:600;color:#1a2233'>Bulk SIM inventory upload</span>
    </div>

    <div style='font-size:13px;color:#5b6779;line-height:1.55;margin-bottom:14px;max-width:70ch'>
      Export your country's inventory to Excel, edit as many rows as you like, and upload the file back. Every row is validated before anything is written to the list.
    </div>

    <div style='display:flex;flex-wrap:wrap;gap:10px'>

      <div style='flex:1 1 190px;background:#f7f9fc;border:1px solid #e3e8ef;border-radius:8px;padding:11px 13px;box-sizing:border-box'>
        <div style='font-size:10px;font-weight:700;letter-spacing:.6px;color:#8a94a6;margin-bottom:3px'>STEP 1 — EXPORT</div>
        <div style='font-size:12.5px;color:#1a2233;line-height:1.5'>Pick a country and download the workbook. It arrives with dropdowns and validation already built in.</div>
      </div>

      <div style='flex:1 1 190px;background:#f7f9fc;border:1px solid #e3e8ef;border-radius:8px;padding:11px 13px;box-sizing:border-box'>
        <div style='font-size:10px;font-weight:700;letter-spacing:.6px;color:#8a94a6;margin-bottom:3px'>STEP 2 — EDIT</div>
        <div style='font-size:12.5px;color:#1a2233;line-height:1.5'>Change existing rows or add new ones. Bad phone numbers and ICC_IDs turn red as you type.</div>
      </div>

      <div style='flex:1 1 190px;background:#f7f9fc;border:1px solid #e3e8ef;border-radius:8px;padding:11px 13px;box-sizing:border-box'>
        <div style='font-size:10px;font-weight:700;letter-spacing:.6px;color:#8a94a6;margin-bottom:3px'>STEP 3 — UPLOAD</div>
        <div style='font-size:12.5px;color:#1a2233;line-height:1.5'>Upload the file back. You get an emailed report showing exactly what was created, updated or rejected.</div>
      </div>

    </div>

    <div style='display:flex;flex-wrap:wrap;gap:6px 16px;margin-top:13px;padding-top:12px;border-top:1px solid #eef1f6;font-size:11.5px;color:#8a94a6'>
      <span>Rows with errors are rejected individually — the rest still go through.</span>
      <span>Re-uploading the same file is safe.</span>
    </div>

  </div>
</div>
```

## In PowerApps

Add an **HTML text** control and set:

```
HtmlText = "<paste the HTML here>"
```

Because every attribute uses single quotes, you can paste it between the outer double quotes
untouched.

| Property | Value |
|---|---|
| `Width` | `Parent.Width - 40` |
| `Height` | `Self.Height` won't auto-fit — see below |
| `AutoHeight` | `true` |
| `Size` | `13` (base font; the inline sizes override per element) |
| `PaddingTop/Bottom/Left/Right` | `0` — the HTML does its own padding |

**Height.** HTML text controls don't reflow their height automatically in every PowerApps
version. If `AutoHeight` isn't available or misbehaves, set an explicit height with room for the
stacked layout:

```
Height = If(Parent.Width < 700, 420, 260)
```

That's the one place a PowerApps-side breakpoint is needed — the HTML reflows on its own, but the
control's box doesn't know it happened.

## Making it dismissible

No JavaScript, so the toggle lives in PowerApps.

**On the screen's `OnVisible`:**

```
Set(varShowUploadBanner, true)
```

**On the HTML text control:**

```
Visible = varShowUploadBanner
```

**Add a small X button** on top of the banner's top-right corner:

```
Text    = "✕"
OnSelect = Set(varShowUploadBanner, false)
Visible = varShowUploadBanner
```

To make it stick between sessions, patch a per-user record instead:

```
OnSelect = Patch(UserPrefs, LookUp(UserPrefs, Email = User().Email),
                 {BannerDismissed: true})
```

---

## Compact variant

For a narrow sidebar or a spot where the three cards are too much:

```html
<div style='font-family:Segoe UI,system-ui,sans-serif;display:flex;flex-wrap:wrap;align-items:center;gap:10px;background:#f7f9fc;border:1px solid #e3e8ef;border-left:4px solid #1f4e79;border-radius:8px;padding:11px 14px;box-sizing:border-box;width:100%'>
  <span style='display:inline-block;background:#c6efce;color:#006100;font-size:10px;font-weight:700;letter-spacing:.8px;padding:3px 9px;border-radius:999px'>NEW</span>
  <span style='font-size:13.5px;font-weight:600;color:#1a2233'>Bulk SIM inventory upload</span>
  <span style='font-size:12.5px;color:#5b6779;flex:1 1 240px'>Export to Excel, edit, upload back. Validation runs before anything is written.</span>
</div>
```

Same wrap behaviour — the description drops to its own line below about 400px.

---

## A note on the copy

The two lines under the divider are the ones worth keeping if you trim anything:

- *"Rows with errors are rejected individually — the rest still go through."*
- *"Re-uploading the same file is safe."*

Both pre-empt a support question. Without the first, someone seeing "4 rejected" assumes the
whole upload failed and re-does work that already succeeded. Without the second, nobody dares
re-upload after a partial failure — which is exactly what they should do.
