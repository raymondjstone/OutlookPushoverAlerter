# Outlook Pushover Alerter

A classic-Outlook (Windows desktop) COM add-in that watches your Inbox and sends
a **Pushover** push notification whenever mail arrives from a sender you care
about. Each watched address or domain is mapped to an **alert type** that decides
the **message text** and the **Pushover sound** used.

Built as a VSTO C# add-in, same shape as `OutlookSpamPlugin`: it hooks the
Inbox `ItemAdd` event, processes each new mail through an engine
(`PushoverAlerter`), and is driven by a **"Pushover Alerts"** ribbon tab.

> **Platform:** Classic Outlook for Windows only. VSTO add-ins don't run in the
> new Outlook, Outlook on the web, or Outlook for Mac.

## What it does

When a new message lands in the Inbox, the add-in resolves the sender's SMTP
address, checks it against the watch list (exact address, or domain incl.
subdomains), and if it matches, sends a Pushover notification using the alert
type assigned to that sender. The notification title is the alert type name; the
body is the type's text followed by the actual sender and subject.

## Ribbon

| Button | Action |
|---|---|
| **Watch Sender** | Add the selected message's sender address to the watch list |
| **Watch Domain** | Add the selected sender's whole domain to the watch list |
| **Watch List** | Add/remove watched senders and assign each one an alert type |
| **Alert Types** | Create/edit alert types (name, sound, text) |
| **Pushover Settings** | Enter your API token + user key, pick the default type, send a test |
| **Edit Files** | Open the three data files in Notepad |
| **Send Test** | Fire a test notification to confirm credentials work |
| **Rescan Unread** | Re-check unread Inbox mail and alert on watched senders |
| **Alerting On/Off** | Master switch for automatic alerting |
| **Confirmations On/Off** | Show/suppress pop-ups for manual actions |

## Data files

Stored in `%APPDATA%\OutlookPushoverAlerter\`:

- `watchlist.txt` - `address-or-domain = alert type name` (one per line; `#` comment).
- `alerttypes.txt` - tab-separated `name<TAB>sound<TAB>text` (edit via the dialog).
- `settings.txt` - `key=value` (Pushover token/user key, default type, toggles).
- `alerts.log` - timestamped record of sends and failures.

See **BUILD-AND-INSTALL.md** for setup, build and install steps.
