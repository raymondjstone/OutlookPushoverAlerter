# Outlook Pushover Alerter

A classic-Outlook (Windows desktop) COM add-in that watches your Inbox and sends
a **[Pushover](https://pushover.net)** push notification whenever mail arrives
from a sender you care about. Each watched address or domain is mapped to an
**alert type** that decides the **message text** and the **Pushover sound** used —
including your own **custom uploaded sounds**.

Built as a VSTO C# add-in: it hooks the Inbox `ItemAdd` event, processes each new
mail through an engine (`PushoverAlerter`), and is driven by a **"Pushover Alerts"**
ribbon tab.

> **Platform:** Classic Outlook for Windows only. VSTO add-ins do **not** run in
> the new Outlook, Outlook on the web, or Outlook for Mac.

---

## Table of contents

- [What it does](#what-it-does)
- [Quick start](#quick-start)
- [The ribbon](#the-ribbon)
- [Concepts: watch entries & alert types](#concepts-watch-entries--alert-types)
- [Custom Pushover sounds](#custom-pushover-sounds)
- [Testing a sound or alert type](#testing-a-sound-or-alert-type)
- [How sender matching works](#how-sender-matching-works)
- [Data files](#data-files)
- [Settings reference](#settings-reference)
- [Architecture / source map](#architecture--source-map)
- [Build & install](#build--install)
- [Troubleshooting](#troubleshooting)
- [FAQ](#faq)

---

## What it does

When a new message lands in the Inbox, the add-in:

1. Resolves the sender's real **SMTP address** (it unwraps Exchange `EX`
   addresses to their primary SMTP address, so internal senders match too).
2. Checks that address — and its domain — against your **watch list**.
3. If it matches, looks up the **alert type** assigned to that sender (falling
   back to your default type).
4. Sends a **Pushover** notification on a background thread (so a slow network
   never blocks Outlook). The notification **title** is the alert type's name;
   the **body** is the type's text followed by the actual `From:` and `Subject:`.

There is also a **Rescan Unread** action that sweeps mail already sitting unread
in the Inbox and alerts on any matching senders (read mail is skipped so a rescan
never spams you with old messages).

---

## Quick start

1. **Get Pushover credentials** — sign up at https://pushover.net, install the
   phone app, copy your **User Key**, and create an **Application/API Token**.
2. **Build & load the add-in** (see [Build & install](#build--install)).
3. In Outlook, open the **Pushover Alerts** ribbon tab → **Pushover Settings**,
   paste your **API token** and **User Key**, and click **Send test** — your
   phone should buzz.
4. Open **Alert Types** and create one or two types (name, sound, text).
5. Select a message and click **Watch Sender** (or **Watch Domain**), or open
   **Watch List** to add senders by hand and assign each an alert type.
6. Leave **Alerting** toggled **On**. New matching mail now pushes automatically.

---

## The ribbon

| Button | Action |
|---|---|
| **Watch Sender** | Add the selected message's sender **address** to the watch list |
| **Watch Domain** | Add the selected sender's whole **domain** to the watch list |
| **Watch List** | Add/remove watched senders and assign each one an alert type |
| **Alert Types** | Create/edit alert types (name, sound, text); **Test** a type live |
| **Custom Sounds** | Add/remove the names of custom sounds uploaded to your Pushover app; **Test** one live |
| **Pushover Settings** | Enter your API token + user key, pick the default type, send a test |
| **Edit Files** | Open the data files (`watchlist`, `alerttypes`, `customsounds`, `settings`) in Notepad |
| **Send Test** | Fire a test notification to confirm credentials work |
| **Rescan Unread** | Re-check unread Inbox mail and alert on watched senders |
| **Alerting On/Off** | Master switch for automatic alerting |
| **Confirmations On/Off** | Show/suppress pop-ups for manual actions |

---

## Concepts: watch entries & alert types

**Watch entry** — a sender pattern (an exact `address@domain` or a bare
`domain`) mapped to the **name** of an alert type. Stored in `watchlist.txt`.
An entry with no alert type uses the **default** type from settings.

**Alert type** — a named bundle of:

- **Name** — also used as the notification **title**.
- **Sound** — a Pushover built-in (e.g. `siren`, `bike`) **or** a custom sound
  name you uploaded to your Pushover app. Blank means "use your Pushover default".
- **Text** — a short line prepended to the notification body.

Example setup:

| Alert type | Sound | Text |
|---|---|---|
| `Urgent` | `siren` | "Urgent email just arrived." |
| `Work` | `bike` | "New work email." |
| `Boss` | `my_custom_chime` *(custom)* | "Message from the boss." |

…then in the watch list:

```
boss@acme.com   = Boss
acme.com        = Work        (domain + all subdomains)
alerts@bank.com = Urgent
newsletter@x.io               (no '=' -> uses the default alert type)
```

---

## Custom Pushover sounds

Pushover lets you upload your own notification sounds to your application
(pushover.net → **Your Applications** → *(your app)* → **sounds**). Each uploaded
sound gets a **name**, and Pushover plays it when you pass that name as the
`sound` parameter.

To use one here:

1. Upload the sound on the Pushover website and note the **name** it's given.
2. In Outlook → **Custom Sounds**, type that name and **Add** (use **Test** to
   confirm it plays). Alternatively, type the name straight into the **Sound**
   box on the **Alert Types** dialog (the box is an editable dropdown) and **Save**.
3. The name is remembered in `customsounds.txt` and appears in the **Sound**
   dropdown for every alert type thereafter.

Names are normalized to Pushover's accepted format automatically: lowercased,
spaces become `_`, and any other disallowed characters are dropped (so
`My Chime!` becomes `my_chime`). You can also edit `customsounds.txt` directly
(one name per line) via **Edit Files**.

> The add-in can't verify a custom name against your Pushover account; if the
> name doesn't match a sound actually uploaded to **that** application token, the
> send still succeeds but the device plays your default sound. Use **Test** to
> confirm — see below.

---

## Testing a sound or alert type

The **Alert Types** dialog has a **Test** button. It sends a **real** Pushover
notification using whatever is currently in the editor (name, sound, text)
without saving — the fastest way to confirm a custom sound name actually plays on
your device. On failure it shows the exact error returned by Pushover (also
written to `alerts.log`).

The ribbon's **Send Test** and the **Pushover Settings** dialog's **Send test**
button do the same thing for a generic/default notification, mainly to confirm
your token and user key are correct.

---

## How sender matching works

- A watch entry containing `@` matches that **exact address**.
- A watch entry with no `@` is a **domain** and matches the domain **and any
  subdomain** of it (`acme.com` also matches `mail.acme.com`).
- Matching is case-insensitive; a leading `@` on a domain entry is ignored.
- Alerts fire as mail arrives while Outlook is open. **Rescan Unread** sweeps
  unread mail already in the Inbox; read mail is skipped.
- Optionally set **TargetAddresses** (settings) to only alert on mail addressed
  **to** specific addresses (semicolon-separated) — useful for shared mailboxes.

---

## Data files

All stored in `%APPDATA%\OutlookPushoverAlerter\` and editable via **Edit Files**:

| File | Format |
|---|---|
| `watchlist.txt` | `address-or-domain = alert type name`, one per line (`#` = comment) |
| `alerttypes.txt` | Tab-separated `name<TAB>sound<TAB>text` (best edited via the dialog) |
| `customsounds.txt` | One custom Pushover sound name per line (`#` = comment) |
| `settings.txt` | `key=value` pairs (token, user key, default type, toggles) |
| `alerts.log` | Timestamped record of every send and failure |

The files are created on first run with header comments explaining each format.

---

## Settings reference

Defaults live at the top of `PushoverAlerter.cs`; most are also editable through
the dialogs or by editing `settings.txt`:

| Key | Default | Meaning |
|---|---|---|
| `PushoverAppToken` | `""` | Your Pushover **application/API token** (required) |
| `PushoverUserKey` | `""` | Your Pushover **user (or group) key** (required) |
| `Enabled` | `true` | Master switch for automatic alerting |
| `ShowConfirmations` | `true` | Pop-up confirmations for manual actions |
| `DefaultAlertType` | `""` | Alert type used when a watch entry has none |
| `TargetAddresses` | `""` | If set, only alert on mail sent to these (semicolon-separated) |
| `LogToFile` | `true` | Append every send/failure to `alerts.log` |

Pushover limits: the notification **message** is capped at 1024 characters and
the **title** at 250; the add-in truncates to fit.

---

## Architecture / source map

| File | Role |
|---|---|
| `OutlookPushoverAlerter.sln` | Solution — open this in Visual Studio |
| `OutlookPushoverAlerter.csproj` | Project (references, build config, VSTO flavor) |
| `ThisAddIn.cs` | Add-in entry point — hooks the Inbox `ItemAdd` event, wires the ribbon |
| `ThisAddIn.Designer.cs` / `.xml` | VSTO-generated plumbing (`Globals`, base class) |
| `PushoverAlerter.cs` | The engine: watch list, alert types, custom sounds, settings, Pushover send |
| `Ribbon1.cs` | Ribbon button callbacks |
| `Ribbon1.xml` | The "Pushover Alerts" ribbon tab layout (embedded resource) |
| `WatchListForm.cs` | Dialog: manage watched senders + their alert type |
| `AlertTypeForm.cs` | Dialog: manage alert types (name, sound, text) + live **Test** |
| `PushoverSettingsForm.cs` | Dialog: API token, user key, default type, test button |
| `install.ps1` / `uninstall.ps1` | Build + register / unregister the add-in |
| `OutlookPushoverAlerter_TemporaryKey.pfx` | Self-signed manifest-signing key |

**Send flow:** `ThisAddIn` Inbox `ItemAdd` → `PushoverAlerter.Process(mail)` →
resolve SMTP → `MatchWatch` → `ResolveAlertType` → `BuildMessage` →
`SendPushoverAsync` (POSTs `token`/`user`/`title`/`message`/`sound` to
`https://api.pushover.net/1/messages.json` over TLS 1.2).

---

## Build & install

**Prerequisites (one time):**

- **Visual Studio 2022** (Community is fine) with the **"Office/SharePoint
  development"** workload (includes the VSTO runtime and templates).
- .NET Framework 4.7.2+ (already present on Windows 10/11).
- A Pushover account, user key, and application/API token.

**Build & run from Visual Studio:**

1. Open **`OutlookPushoverAlerter.sln`**; accept any retarget / NuGet-restore prompt.
2. Choose the **Release** configuration → **Build → Build Solution** (Ctrl+Shift+B).
3. Press **F5** to launch Outlook with the add-in loaded for debugging.

**Build & register from PowerShell (Outlook closed):**

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

This builds Release and registers the add-in under
`HKCU\...\Outlook\Addins\OutlookPushoverAlerter`. To remove it, run
`uninstall.ps1`.

> See **BUILD-AND-INSTALL.md** for the fully detailed, step-by-step walkthrough,
> including notes on the signing key and resolving Office interop references.

---

## Troubleshooting

- **No ribbon tab after install:** **File → Options → Add-ins → COM Add-ins → Go…**
  and tick *OutlookPushoverAlerter*. If `LoadBehavior` got set to `2`, the add-in
  errored on load — check it isn't under **Disabled Items**.
- **Test fails with an HTTP error:** double-check the API **token** and **user
  key** (they're easy to swap). Pushover's error text is shown in the dialog and
  written to `alerts.log`.
- **Custom sound doesn't play (but the notification arrives):** the sound name
  doesn't match one uploaded to that application token. Re-check the name on the
  Pushover website and use **Alert Types → Test** to confirm.
- **Notifications arrive but Outlook feels laggy:** sends already run on a
  background thread, so this is usually unrelated — check your network.

More cases (namespace/`GetCustomUI`, Office interop references, signing key) are
covered in **BUILD-AND-INSTALL.md**.

---

## FAQ

**Does this work in the new Outlook / web / Mac?** No — VSTO COM add-ins are
classic Outlook for Windows only.

**Where are my settings stored?** In `%APPDATA%\OutlookPushoverAlerter\`. Your
Pushover token and user key live in `settings.txt` in plain text on your machine.

**Can one sender map to multiple sounds?** A watch entry maps to one alert type
(one sound). Create separate alert types for different sounds and assign whichever
you want per sender.

**Will a rescan re-alert on mail I've already seen?** No — **Rescan Unread** only
considers **unread** Inbox mail.
