# Outlook Pushover Alerter — VSTO Add-in (Build & Install)

This is a real installable Outlook COM add-in (C# / VSTO). Once built, it loads
automatically with Outlook and appears as a **"Pushover Alerts"** ribbon tab.
When mail arrives from a watched sender, it sends a **Pushover** push
notification using that sender's alert type (text + sound).

> **Platform:** Classic Outlook for Windows (desktop). VSTO add-ins don't run in
> the new Outlook, Outlook on the web, or Outlook for Mac.

---

## What's in this folder (complete, ready-to-open project)

| File | Role |
|---|---|
| `OutlookPushoverAlerter.sln` | Solution file — **open this in Visual Studio** |
| `OutlookPushoverAlerter.csproj` | Project file (references, build config, VSTO flavor) |
| `ThisAddIn.cs` | Add-in entry point — hooks the Inbox and wires the ribbon |
| `ThisAddIn.Designer.cs` | VSTO-generated plumbing (`Globals`, base class) |
| `ThisAddIn.Designer.xml` | VSTO designer file |
| `PushoverAlerter.cs` | The engine (watch list, alert types, settings, Pushover send) |
| `Ribbon1.cs` | Ribbon button callbacks |
| `Ribbon1.xml` | The "Pushover Alerts" ribbon tab layout (embedded resource) |
| `WatchListForm.cs` | Dialog: manage watched senders + their alert type |
| `AlertTypeForm.cs` | Dialog: manage alert types (name, sound, text) |
| `PushoverSettingsForm.cs` | Dialog: API token, user key, default type, test button |
| `app.config` | Target runtime |
| `Properties\AssemblyInfo.cs` | Assembly metadata |
| `OutlookPushoverAlerter_TemporaryKey.pfx` | Manifest-signing key (self-signed; see note) |
| `install.ps1` / `uninstall.ps1` | Build + register / unregister the add-in |

You do **not** need to scaffold anything — just open the solution and build.

---

## Step 1 — Get a Pushover account (one time)

1. Sign up at https://pushover.net and install the Pushover app on your phone.
2. Your **User Key** is on the dashboard after logging in.
3. Create an **Application/API Token**: dashboard → *Create an Application/API Token*.
   The token it gives you is the **API token**.

You'll paste both into **Pushover Settings** inside the add-in later.

---

## Step 2 — Install the tools (one time)

1. Download **Visual Studio Community 2022** (free): https://visualstudio.microsoft.com/
2. In the installer, on the **Workloads** screen, check
   **"Office/SharePoint development"** (includes the VSTO runtime and templates).
3. .NET Framework 4.7.2+ is already on Windows 10/11.

---

## Step 3 — Open and build

1. Keep all files in this folder together.
2. Double-click **`OutlookPushoverAlerter.sln`** to open it in Visual Studio.
3. Accept any one-time **retarget** prompt; allow **NuGet/assembly restore** if offered.
4. Set the configuration to **Release**, then **Build → Build Solution** (Ctrl+Shift+B).

### Test it immediately
Press **F5** (Start). Outlook launches with the add-in loaded; you'll see the
**Pushover Alerts** tab. Open **Pushover Settings**, paste your token + user key,
and click **Send test** — your phone should buzz. Close Outlook to stop debugging.

### Or build + register from PowerShell
With Outlook closed, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

This builds Release and registers the add-in under
`HKCU\...\Outlook\Addins\OutlookPushoverAlerter`. To remove it, run `uninstall.ps1`.

---

## Step 4 — Configure

1. **Alert Types** — create the types you want, e.g.
   - `Urgent` — sound `siren`, text "Urgent email just arrived."
   - `Work` — sound `bike`, text "New work email."
2. **Watch List** — add the addresses/domains to watch and pick an alert type for
   each. Entries with no type fall back to the **default** alert type set in
   **Pushover Settings**.
3. Leave **Alerting** toggled **On**. New matching mail now pushes automatically.

---

## How matching works

- A watch entry containing `@` matches that **exact address**.
- A watch entry with no `@` is a **domain** and matches the domain and any
  subdomain (`acme.com` also matches `mail.acme.com`).
- Alerts fire as mail arrives while Outlook is open. **Rescan Unread** sweeps
  unread mail already in the Inbox (read mail is skipped so you're not spammed).

---

## Configuration fields (advanced)

Defaults live at the top of `PushoverAlerter.cs`; most are also editable through
the dialogs / `settings.txt`:

| Field | Default | Meaning |
|---|---|---|
| `ListFolder` | `%APPDATA%\OutlookPushoverAlerter` | Where the data files live |
| `Enabled` | `true` | Master switch for automatic alerting |
| `DefaultAlertType` | `""` | Alert type used when a watch entry has none |
| `TargetAddresses` | `""` | If set, only alert on mail sent to these (semicolon separated) |
| `LogToFile` | `true` | Append every send/failure to `alerts.log` |

`message` is capped at 1024 chars and `title` at 250, per Pushover's limits.

---

## Note on the signing key

`OutlookPushoverAlerter_TemporaryKey.pfx` is a **self-signed** code-signing key
(no password) used to sign the ClickOnce/VSTO manifests, exactly like a Visual
Studio-generated temporary key. It's fine for personal use.

If Visual Studio ever complains the key can't be used, regenerate it in two
clicks: **Project → Properties → Signing → "Create Test Certificate…"** (leave the
password blank), then rebuild. To produce a no-prompt installer for other PCs,
sign with a real code-signing certificate instead.

---

## Troubleshooting

- **No ribbon tab after install:** **File → Options → Add-ins → COM Add-ins → Go…**
  and tick *OutlookPushoverAlerter*. If `LoadBehavior` got set to 2, the add-in
  errored on load — check it isn't disabled under **Disabled Items**.
- **`GetCustomUI` returns nothing (no tab):** almost always a namespace mismatch —
  confirm the project's default namespace is `OutlookPushoverAlerter`, matching the
  `OutlookPushoverAlerter.Ribbon1.xml` resource string in `Ribbon1.cs`.
- **Test fails with an HTTP error:** double-check the API **token** and **user key**
  (they're easy to swap). The error text from Pushover is shown in the dialog and
  written to `alerts.log`.
- **An Office reference won't resolve** (yellow warning on `office` or
  `Microsoft.Office.Interop.Outlook`): your installed interop version differs from
  the `15.0.0.0` in the `.csproj`. Right-click **References → Add Reference →
  Assemblies → Extensions**, tick **Microsoft Outlook xx.x Object Library**, remove
  the greyed-out one, and rebuild.
- **Notifications arrive but Outlook feels laggy:** sends already run on a
  background thread, so this is usually unrelated; check your network.
