using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookPushoverAlerter
{
    /// <summary>A named alert type: the text to push and the Pushover sound to use.</summary>
    public class AlertType
    {
        public string Name = "";
        public string Sound = "";   // empty => use the user's Pushover default sound
        public string Text = "";

        public AlertType() { }
        public AlertType(string name, string sound, string text)
        {
            Name = name ?? "";
            Sound = sound ?? "";
            Text = text ?? "";
        }
    }

    /// <summary>A watched sender: an email address or domain, mapped to an alert type name.</summary>
    public class WatchEntry
    {
        public string Pattern = "";    // full address (contains @) or bare domain
        public string AlertType = "";  // name of the AlertType to fire ("" => default)

        public WatchEntry() { }
        public WatchEntry(string pattern, string alertType)
        {
            Pattern = pattern ?? "";
            AlertType = alertType ?? "";
        }
    }

    /// <summary>
    /// All of the watch / alert-type / Pushover logic lives here. When mail
    /// arrives from a watched address or domain, a Pushover push notification is
    /// sent using that sender's assigned alert type (text + sound).
    /// </summary>
    public class PushoverAlerter
    {
        // -------------------- Storage locations --------------------

        // Where the data files live (per-user AppData by default).
        public string ListFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OutlookPushoverAlerter");

        public string WatchListFile     { get { return Path.Combine(ListFolder, "watchlist.txt"); } }
        public string AlertTypesFile    { get { return Path.Combine(ListFolder, "alerttypes.txt"); } }
        public string CustomSoundsFile  { get { return Path.Combine(ListFolder, "customsounds.txt"); } }
        public string SettingsFile      { get { return Path.Combine(ListFolder, "settings.txt"); } }
        public string LogFile           { get { return Path.Combine(ListFolder, "alerts.log"); } }

        // -------------------- Settings --------------------

        // Pushover credentials. AppToken = your application's API token/key.
        // UserKey = your user (or group) key. Both are required to send.
        public string PushoverAppToken = "";
        public string PushoverUserKey = "";

        // Master on/off switch for automatic alerting.
        public bool Enabled = true;

        // Show pop-up confirmations for manual actions (Test, add to watch list).
        public bool ShowConfirmations = true;

        // Alert type used when a watch entry has no type assigned (by name).
        public string DefaultAlertType = "";

        // OPTIONAL: only alert on mail sent TO these addresses (semicolon
        // separated). Leave empty to consider ALL incoming mail.
        public string TargetAddresses = "";

        // Append a line to alerts.log for every send / failure.
        public bool LogToFile = true;

        // Pushover's built-in notification sounds (for the alert-type editor).
        // Custom sounds you upload to your Pushover app are added by name through
        // the Alert Types dialog and persisted in customsounds.txt; see AllSounds().
        public static readonly string[] Sounds = new[]
        {
            "pushover", "bike", "bugle", "cashregister", "classical", "cosmic",
            "falling", "gamelan", "incoming", "intermission", "magic", "mechanical",
            "pianobar", "siren", "spacealarm", "tugboat", "alien", "climb",
            "persistent", "echo", "updown", "vibrate", "none"
        };

        private const string PR_SMTP =
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        private const string PushoverEndpoint = "https://api.pushover.net/1/messages.json";

        // -------------------- Core flow --------------------

        /// <summary>Handle one freshly-arrived mail item.</summary>
        public void Process(Outlook.MailItem mail)
        {
            if (!Enabled) return;
            if (!string.IsNullOrEmpty(TargetAddresses) && !SentToTarget(mail)) return;

            string addr = GetSenderSmtp(mail);
            if (string.IsNullOrEmpty(addr)) return;
            string domain = SplitDomain(addr);

            WatchEntry entry = MatchWatch(addr, domain);
            if (entry == null) return;

            AlertType type = ResolveAlertType(entry.AlertType);
            string subject = "";
            try { subject = mail.Subject ?? ""; } catch { }

            string title;
            string body;
            BuildMessage(type, addr, subject, out title, out body);

            // Send off the Outlook thread so a slow network never blocks Outlook.
            SendPushoverAsync(title, body, type.Sound);
        }

        /// <summary>
        /// Re-check the Inbox and alert on matching UNREAD messages. Read mail is
        /// skipped so a rescan never blasts notifications for old, already-seen
        /// messages. Returns how many alerts were sent.
        /// </summary>
        public int RescanInboxUnread()
        {
            Outlook.NameSpace ns = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            Outlook.Items items =
                ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items;

            int sent = 0;
            for (int i = 1; i <= items.Count; i++)
            {
                Outlook.MailItem mail = items[i] as Outlook.MailItem;
                if (mail == null) continue;
                try { if (!mail.UnRead) continue; } catch { continue; }
                if (!string.IsNullOrEmpty(TargetAddresses) && !SentToTarget(mail)) continue;

                string addr = GetSenderSmtp(mail);
                if (string.IsNullOrEmpty(addr)) continue;
                string domain = SplitDomain(addr);

                WatchEntry entry = MatchWatch(addr, domain);
                if (entry == null) continue;

                AlertType type = ResolveAlertType(entry.AlertType);
                string subject = "";
                try { subject = mail.Subject ?? ""; } catch { }

                string title, body;
                BuildMessage(type, addr, subject, out title, out body);

                string err;
                if (SendPushover(title, body, type.Sound, out err)) sent++;
            }
            return sent;
        }

        // -------------------- Message construction --------------------

        private void BuildMessage(AlertType type, string sender, string subject,
                                  out string title, out string body)
        {
            title = string.IsNullOrEmpty(type.Name) ? "Outlook Alert" : type.Name;
            if (title.Length > 250) title = title.Substring(0, 250);

            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(type.Text)) sb.AppendLine(type.Text.Trim());
            sb.AppendLine("From: " + sender);
            if (!string.IsNullOrEmpty(subject)) sb.AppendLine("Subject: " + subject.Trim());

            body = sb.ToString().TrimEnd();
            if (body.Length > 1024) body = body.Substring(0, 1024);
        }

        // -------------------- Pushover send --------------------

        /// <summary>Fire-and-forget send used by the automatic ItemAdd path.</summary>
        public void SendPushoverAsync(string title, string message, string sound)
        {
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                string err;
                SendPushover(title, message, sound, out err);
            });
        }

        /// <summary>
        /// Synchronous send. Returns true on success; on failure, <paramref name="error"/>
        /// holds a human-readable reason. Never throws.
        /// </summary>
        public bool SendPushover(string title, string message, string sound, out string error)
        {
            error = "";
            if (string.IsNullOrEmpty(PushoverAppToken) || string.IsNullOrEmpty(PushoverUserKey))
            {
                error = "Pushover API token and user key are not set. Open 'Pushover Settings' first.";
                Log("SEND FAILED (no credentials): " + title);
                return false;
            }

            try
            {
                // Pushover requires TLS 1.2.
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                var values = new NameValueCollection
                {
                    { "token", PushoverAppToken },
                    { "user",  PushoverUserKey },
                    { "title", title ?? "" },
                    { "message", string.IsNullOrEmpty(message) ? " " : message }
                };
                if (!string.IsNullOrEmpty(sound)) values.Add("sound", sound);

                using (var client = new WebClient())
                {
                    byte[] resp = client.UploadValues(PushoverEndpoint, "POST", values);
                    string body = Encoding.UTF8.GetString(resp);
                    Log("SENT: " + title + " | " + Flatten(message));
                    return true;
                }
            }
            catch (WebException wex)
            {
                string detail = wex.Message;
                try
                {
                    if (wex.Response != null)
                        using (var r = new StreamReader(wex.Response.GetResponseStream()))
                            detail = r.ReadToEnd();
                }
                catch { }
                error = "Pushover request failed: " + detail;
                Log("SEND FAILED: " + title + " | " + detail);
                return false;
            }
            catch (Exception ex)
            {
                error = "Pushover request failed: " + ex.Message;
                Log("SEND FAILED: " + title + " | " + ex.Message);
                return false;
            }
        }

        /// <summary>Send a test notification using the given alert type (or a default).</summary>
        public bool SendTest(AlertType type, out string error)
        {
            if (type == null) type = new AlertType("Outlook Alert", "pushover", "Test notification from OutlookPushoverAlerter.");
            string title, body;
            BuildMessage(type, "test@example.com", "Pushover test message", out title, out body);
            return SendPushover(title, body, type.Sound, out error);
        }

        // -------------------- Sender / recipient resolution --------------------

        public string GetSenderSmtp(Outlook.MailItem mail)
        {
            try
            {
                if (mail.SenderEmailType == "EX")
                {
                    Outlook.AddressEntry sender = mail.Sender;
                    if (sender != null)
                    {
                        Outlook.ExchangeUser ex = sender.GetExchangeUser();
                        if (ex != null && !string.IsNullOrEmpty(ex.PrimarySmtpAddress))
                            return ex.PrimarySmtpAddress.Trim().ToLower();

                        string v = sender.PropertyAccessor.GetProperty(PR_SMTP) as string;
                        if (!string.IsNullOrEmpty(v)) return v.Trim().ToLower();
                    }
                }
                return (mail.SenderEmailAddress ?? "").Trim().ToLower();
            }
            catch { return ""; }
        }

        public string SplitDomain(string addr)
        {
            int p = addr.IndexOf('@');
            return p >= 0 ? addr.Substring(p + 1).ToLower() : "";
        }

        private bool SentToTarget(Outlook.MailItem mail)
        {
            try
            {
                string[] targets = TargetAddresses.ToLower()
                    .Split(';').Select(t => t.Trim()).Where(t => t.Length > 0).ToArray();
                if (targets.Length == 0) return true;
                return SentToAny(mail, targets);
            }
            catch { return true; }
        }

        private bool SentToAny(Outlook.MailItem mail, string[] targets)
        {
            try
            {
                if (targets == null || targets.Length == 0) return false;

                string hay = "|" + (mail.To ?? "").ToLower() + "|"
                                 + (mail.CC ?? "").ToLower() + "|";

                foreach (Outlook.Recipient r in mail.Recipients)
                {
                    string ra = "";
                    try
                    {
                        Outlook.AddressEntry ae = r.AddressEntry;
                        if (ae != null)
                        {
                            Outlook.ExchangeUser ex = ae.GetExchangeUser();
                            if (ex != null) ra = ex.PrimarySmtpAddress;
                            if (string.IsNullOrEmpty(ra))
                                ra = ae.PropertyAccessor.GetProperty(PR_SMTP) as string;
                        }
                    }
                    catch { }
                    if (string.IsNullOrEmpty(ra)) ra = r.Address;
                    hay += "|" + (ra ?? "").Trim().ToLower() + "|";
                }

                return targets.Any(t => hay.Contains(t));
            }
            catch { return false; }
        }

        // -------------------- Watch list --------------------

        /// <summary>
        /// Returns the watch entry whose pattern matches the sender, or null.
        /// A pattern containing '@' matches the exact address; otherwise it is a
        /// domain and matches that domain and any subdomain of it.
        /// </summary>
        public WatchEntry MatchWatch(string addr, string domain)
        {
            foreach (WatchEntry e in ReadWatchEntries())
            {
                string p = e.Pattern;
                if (p.Length == 0) continue;

                if (p.Contains("@"))
                {
                    if (p == addr) return e;
                }
                else
                {
                    if (p == domain ||
                        (domain.Length > p.Length && domain.EndsWith("." + p)))
                        return e;
                }
            }
            return null;
        }

        /// <summary>Reads watchlist.txt into (pattern, alertType) pairs.</summary>
        public List<WatchEntry> ReadWatchEntries()
        {
            var list = new List<WatchEntry>();
            try
            {
                if (!File.Exists(WatchListFile)) return list;
                foreach (string raw in File.ReadAllLines(WatchListFile))
                {
                    string line = (raw ?? "").Trim();
                    if (line.Length == 0 || line[0] == '#' || line[0] == '\'') continue;

                    string pattern, type;
                    int eq = line.IndexOf('=');
                    if (eq >= 0)
                    {
                        pattern = line.Substring(0, eq).Trim();
                        type = line.Substring(eq + 1).Trim();
                    }
                    else
                    {
                        pattern = line;
                        type = "";
                    }

                    pattern = pattern.ToLower();
                    if (pattern.StartsWith("@")) pattern = pattern.Substring(1);
                    if (pattern.Length == 0) continue;

                    list.Add(new WatchEntry(pattern, type));
                }
            }
            catch { }
            return list;
        }

        /// <summary>Rewrites watchlist.txt from the given entries (de-duplicated by pattern).</summary>
        public void SaveWatchList(IEnumerable<WatchEntry> entries)
        {
            EnsureListFiles();

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var sb = new StringBuilder();
            sb.AppendLine("# Watch list - one sender per line.");
            sb.AppendLine("# Format:  address-or-domain = alert type name");
            sb.AppendLine("#   boss@acme.com   = Urgent");
            sb.AppendLine("#   acme.com        = Work       (domain + subdomains)");
            sb.AppendLine("#   newsletter@x.io             (no '='  -> uses the default alert type)");

            foreach (WatchEntry e in entries ?? Enumerable.Empty<WatchEntry>())
            {
                string p = (e.Pattern ?? "").Trim().ToLower();
                if (p.StartsWith("@")) p = p.Substring(1);
                if (p.Length == 0) continue;
                if (!seen.Add(p)) continue;

                string t = (e.AlertType ?? "").Trim();
                sb.AppendLine(t.Length > 0 ? (p + " = " + t) : p);
            }

            File.WriteAllText(WatchListFile, sb.ToString());
        }

        public void AppendWatch(string pattern, string alertType)
        {
            EnsureListFiles();
            pattern = (pattern ?? "").Trim().ToLower();
            if (pattern.StartsWith("@")) pattern = pattern.Substring(1);
            if (pattern.Length == 0) return;

            var entries = ReadWatchEntries();
            // Replace an existing entry with the same pattern, else add.
            entries.RemoveAll(e => string.Equals(e.Pattern, pattern, StringComparison.OrdinalIgnoreCase));
            entries.Add(new WatchEntry(pattern, alertType ?? ""));
            SaveWatchList(entries);
        }

        // -------------------- Alert types --------------------

        /// <summary>Reads alerttypes.txt (tab-separated: name, sound, text).</summary>
        public List<AlertType> ReadAlertTypes()
        {
            var list = new List<AlertType>();
            try
            {
                if (!File.Exists(AlertTypesFile)) return list;
                foreach (string raw in File.ReadAllLines(AlertTypesFile))
                {
                    string line = raw ?? "";
                    string trimmed = line.Trim();
                    if (trimmed.Length == 0 || trimmed[0] == '#') continue;

                    string[] parts = line.Split('\t');
                    string name = parts.Length > 0 ? parts[0].Trim() : "";
                    string sound = parts.Length > 1 ? parts[1].Trim() : "";
                    string text = parts.Length > 2 ? parts[2] : "";
                    if (name.Length == 0) continue;

                    list.Add(new AlertType(name, sound, text));
                }
            }
            catch { }
            return list;
        }

        /// <summary>Rewrites alerttypes.txt from the given types (de-duplicated by name).</summary>
        public void SaveAlertTypes(IEnumerable<AlertType> types)
        {
            EnsureListFiles();

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var sb = new StringBuilder();
            sb.AppendLine("# Alert types - one per line, TAB separated:  name<TAB>sound<TAB>text");
            sb.AppendLine("# 'sound' is one of Pushover's built-in sounds (blank = your default).");
            sb.AppendLine("# Edit these through the 'Alert Types' dialog rather than by hand.");

            foreach (AlertType t in types ?? Enumerable.Empty<AlertType>())
            {
                string name = (t.Name ?? "").Trim();
                if (name.Length == 0) continue;
                if (!seen.Add(name)) continue;

                string sound = (t.Sound ?? "").Trim();
                string text = Flatten(t.Text);   // strip tabs/newlines so the row stays intact
                sb.AppendLine(name + "\t" + sound + "\t" + text);
            }

            File.WriteAllText(AlertTypesFile, sb.ToString());
        }

        public AlertType GetAlertTypeByName(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            return ReadAlertTypes()
                .FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Picks the alert type to use: the named one if it exists, else the
        /// configured default, else the first defined type, else a built-in fallback.
        /// </summary>
        public AlertType ResolveAlertType(string name)
        {
            AlertType t = GetAlertTypeByName(name);
            if (t != null) return t;

            t = GetAlertTypeByName(DefaultAlertType);
            if (t != null) return t;

            var all = ReadAlertTypes();
            if (all.Count > 0) return all[0];

            return new AlertType("Outlook Alert", "pushover",
                "You have new mail from a watched sender.");
        }

        // -------------------- Custom sounds --------------------

        /// <summary>
        /// Reads customsounds.txt: one sound name per line. These are the names of
        /// sounds you've uploaded to your Pushover application (Pushover references
        /// custom sounds by the lowercase name they were given at upload time).
        /// </summary>
        public List<string> ReadCustomSounds()
        {
            var list = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (!File.Exists(CustomSoundsFile)) return list;
                foreach (string raw in File.ReadAllLines(CustomSoundsFile))
                {
                    string line = (raw ?? "").Trim();
                    if (line.Length == 0 || line[0] == '#') continue;

                    string name = NormalizeSoundName(line);
                    if (name.Length == 0) continue;
                    if (!seen.Add(name)) continue;

                    list.Add(name);
                }
            }
            catch { }
            return list;
        }

        /// <summary>Rewrites customsounds.txt from the given names (de-duplicated, built-ins skipped).</summary>
        public void SaveCustomSounds(IEnumerable<string> names)
        {
            EnsureListFiles();

            var builtIn = new HashSet<string>(Sounds, StringComparer.OrdinalIgnoreCase);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var sb = new StringBuilder();
            sb.AppendLine("# Custom Pushover sounds - one name per line.");
            sb.AppendLine("# Use the exact name a sound was given when uploaded to your Pushover app");
            sb.AppendLine("# (https://pushover.net -> Your Applications -> sounds). Lowercase, no spaces.");

            foreach (string raw in names ?? Enumerable.Empty<string>())
            {
                string name = NormalizeSoundName(raw);
                if (name.Length == 0) continue;
                if (builtIn.Contains(name)) continue;   // already covered by the built-in list
                if (!seen.Add(name)) continue;

                sb.AppendLine(name);
            }

            File.WriteAllText(CustomSoundsFile, sb.ToString());
        }

        /// <summary>Adds one custom sound name (if new and not a built-in). Returns the normalized name.</summary>
        public string AddCustomSound(string name)
        {
            name = NormalizeSoundName(name);
            if (name.Length == 0) return "";
            if (Sounds.Any(s => string.Equals(s, name, StringComparison.OrdinalIgnoreCase))) return name;

            var sounds = ReadCustomSounds();
            if (!sounds.Any(s => string.Equals(s, name, StringComparison.OrdinalIgnoreCase)))
            {
                sounds.Add(name);
                SaveCustomSounds(sounds);
            }
            return name;
        }

        /// <summary>The built-in sounds followed by any user-added custom sound names.</summary>
        public List<string> AllSounds()
        {
            var all = new List<string>(Sounds);
            var seen = new HashSet<string>(Sounds, StringComparer.OrdinalIgnoreCase);
            foreach (string s in ReadCustomSounds())
                if (seen.Add(s)) all.Add(s);
            return all;
        }

        /// <summary>
        /// Normalizes a sound name to what Pushover accepts: lowercase, with spaces
        /// collapsed to underscores and any other disallowed characters stripped.
        /// </summary>
        public static string NormalizeSoundName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "";
            var sb = new StringBuilder();
            foreach (char c in name.Trim().ToLowerInvariant())
            {
                if ((c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '-')
                    sb.Append(c);
                else if (c == ' ')
                    sb.Append('_');
                // anything else is dropped
            }
            return sb.ToString();
        }

        // -------------------- First-run setup --------------------

        public void EnsureListFiles()
        {
            try
            {
                Directory.CreateDirectory(ListFolder);

                if (!File.Exists(WatchListFile))
                    File.WriteAllText(WatchListFile,
                        "# Watch list - one sender per line." + Environment.NewLine +
                        "# Format:  address-or-domain = alert type name" + Environment.NewLine +
                        "#   boss@acme.com   = Urgent" + Environment.NewLine +
                        "#   acme.com        = Work       (domain + subdomains)" + Environment.NewLine +
                        "#   newsletter@x.io             (no '='  -> uses the default alert type)" + Environment.NewLine);

                if (!File.Exists(AlertTypesFile))
                    File.WriteAllText(AlertTypesFile,
                        "# Alert types - one per line, TAB separated:  name<TAB>sound<TAB>text" + Environment.NewLine +
                        "# 'sound' is one of Pushover's built-in sounds (blank = your default)." + Environment.NewLine +
                        "# Edit these through the 'Alert Types' dialog rather than by hand." + Environment.NewLine +
                        "Default\tpushover\tNew mail from a watched sender." + Environment.NewLine);

                if (!File.Exists(CustomSoundsFile))
                    File.WriteAllText(CustomSoundsFile,
                        "# Custom Pushover sounds - one name per line." + Environment.NewLine +
                        "# Use the exact name a sound was given when uploaded to your Pushover app" + Environment.NewLine +
                        "# (https://pushover.net -> Your Applications -> sounds). Lowercase, no spaces." + Environment.NewLine);
            }
            catch { }
        }

        // -------------------- Settings persistence --------------------

        private bool _settingsLoaded;

        public void EnsureSettingsLoaded()
        {
            if (_settingsLoaded) return;
            LoadSettings();
            _settingsLoaded = true;
        }

        public void LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsFile)) return;
                foreach (string raw in File.ReadAllLines(SettingsFile))
                {
                    string line = (raw ?? "").Trim();
                    if (line.Length == 0 || line[0] == '#') continue;

                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string key = line.Substring(0, eq).Trim().ToLower();
                    string val = line.Substring(eq + 1).Trim();

                    switch (key)
                    {
                        case "pushoverapptoken": PushoverAppToken = val; break;
                        case "pushoveruserkey":  PushoverUserKey = val; break;
                        case "enabled":          Enabled = ParseBool(val, Enabled); break;
                        case "showconfirmations":ShowConfirmations = ParseBool(val, ShowConfirmations); break;
                        case "defaultalerttype": DefaultAlertType = val; break;
                        case "targetaddresses":  TargetAddresses = val; break;
                        case "logtofile":        LogToFile = ParseBool(val, LogToFile); break;
                    }
                }
            }
            catch { }
        }

        public void SaveSettings()
        {
            try
            {
                Directory.CreateDirectory(ListFolder);
                var sb = new StringBuilder();
                sb.AppendLine("# OutlookPushoverAlerter settings");
                sb.AppendLine("PushoverAppToken=" + PushoverAppToken);
                sb.AppendLine("PushoverUserKey=" + PushoverUserKey);
                sb.AppendLine("Enabled=" + (Enabled ? "true" : "false"));
                sb.AppendLine("ShowConfirmations=" + (ShowConfirmations ? "true" : "false"));
                sb.AppendLine("DefaultAlertType=" + DefaultAlertType);
                sb.AppendLine("TargetAddresses=" + TargetAddresses);
                sb.AppendLine("LogToFile=" + (LogToFile ? "true" : "false"));
                File.WriteAllText(SettingsFile, sb.ToString());
                _settingsLoaded = true;
            }
            catch { }
        }

        // -------------------- Helpers --------------------

        private void Log(string msg)
        {
            if (!LogToFile) return;
            try
            {
                Directory.CreateDirectory(ListFolder);
                File.AppendAllText(LogFile,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + msg + Environment.NewLine);
            }
            catch { }
        }

        private static string Flatten(string s)
        {
            return (s ?? "").Replace("\t", " ").Replace("\r", " ").Replace("\n", " ").Trim();
        }

        private static bool ParseBool(string s, bool fallback)
        {
            s = (s ?? "").Trim().ToLower();
            if (s == "true" || s == "1" || s == "yes" || s == "on") return true;
            if (s == "false" || s == "0" || s == "no" || s == "off") return false;
            return fallback;
        }
    }
}
