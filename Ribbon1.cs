using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookPushoverAlerter
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;
        private Office.IRibbonUI ribbon;

        public Ribbon1(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonID)
        {
            // Loads the embedded Ribbon1.xml. The resource name is
            // "<DefaultNamespace>.Ribbon1.xml" -> keep the project's default
            // namespace as OutlookPushoverAlerter, or update the string below.
            return GetResourceText("OutlookPushoverAlerter.Ribbon1.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        private PushoverAlerter Alerter { get { return _addIn.Alerter; } }

        // ---- Watch selected ----

        public void OnWatchSender(Office.IRibbonControl c)
        {
            try { WatchSelected(false); }
            catch (Exception ex) { ShowError(ex); }
        }

        public void OnWatchDomain(Office.IRibbonControl c)
        {
            try { WatchSelected(true); }
            catch (Exception ex) { ShowError(ex); }
        }

        // ---- Manage dialogs ----

        public void OnEditWatchList(Office.IRibbonControl c)
        {
            try
            {
                Alerter.EnsureListFiles();
                using (var form = new WatchListForm(Alerter))
                    form.ShowDialog();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        public void OnEditAlertTypes(Office.IRibbonControl c)
        {
            try
            {
                Alerter.EnsureListFiles();
                using (var form = new AlertTypeForm(Alerter))
                    form.ShowDialog();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        public void OnEditSettings(Office.IRibbonControl c)
        {
            try
            {
                Alerter.EnsureSettingsLoaded();
                using (var form = new PushoverSettingsForm(Alerter))
                    form.ShowDialog();
            }
            catch (Exception ex) { ShowError(ex); }
        }

        public void OnEditFiles(Office.IRibbonControl c)
        {
            try
            {
                Alerter.EnsureListFiles();
                Alerter.SaveSettings(); // make sure settings.txt exists to open
                Process.Start("notepad.exe", "\"" + Alerter.WatchListFile + "\"");
                Process.Start("notepad.exe", "\"" + Alerter.AlertTypesFile + "\"");
                Process.Start("notepad.exe", "\"" + Alerter.SettingsFile + "\"");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        // ---- Actions ----

        public void OnTest(Office.IRibbonControl c)
        {
            try
            {
                Alerter.EnsureSettingsLoaded();
                AlertType type = Alerter.ResolveAlertType(Alerter.DefaultAlertType);
                string err;
                bool ok = Alerter.SendTest(type, out err);
                if (ok)
                    MessageBox.Show("Test notification sent via Pushover (sound: " +
                        (string.IsNullOrEmpty(type.Sound) ? "default" : type.Sound) + ").",
                        "Pushover Alerts");
                else
                    MessageBox.Show(err, "Pushover Alerts - Test failed",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex) { ShowError(ex); }
        }

        public void OnRescan(Office.IRibbonControl c)
        {
            try
            {
                Alerter.EnsureSettingsLoaded();
                int n = Alerter.RescanInboxUnread();
                Confirm("Rescan complete. Sent " + n + " alert(s) for unread mail from watched senders.");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        // ---- Enabled toggle ----

        public bool GetEnabledPressed(Office.IRibbonControl c)
        {
            Alerter.EnsureSettingsLoaded();
            return Alerter.Enabled;
        }

        public string GetEnabledLabel(Office.IRibbonControl c)
        {
            Alerter.EnsureSettingsLoaded();
            return Alerter.Enabled ? "Alerting: On" : "Alerting: Off";
        }

        public void OnToggleEnabled(Office.IRibbonControl c, bool pressed)
        {
            try
            {
                Alerter.Enabled = pressed;
                Alerter.SaveSettings();
                if (ribbon != null) ribbon.InvalidateControl("btnEnabled");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        // ---- Confirmations toggle ----

        public bool GetConfirmationsPressed(Office.IRibbonControl c)
        {
            Alerter.EnsureSettingsLoaded();
            return Alerter.ShowConfirmations;
        }

        public string GetConfirmationsLabel(Office.IRibbonControl c)
        {
            Alerter.EnsureSettingsLoaded();
            return Alerter.ShowConfirmations ? "Confirmations: On" : "Confirmations: Off";
        }

        public void OnToggleConfirmations(Office.IRibbonControl c, bool pressed)
        {
            try
            {
                Alerter.ShowConfirmations = pressed;
                Alerter.SaveSettings();
                if (ribbon != null) ribbon.InvalidateControl("btnConfirmations");
            }
            catch (Exception ex) { ShowError(ex); }
        }

        // ---- Shared helpers ----

        private void WatchSelected(bool useDomain)
        {
            Outlook.Explorer explorer = _addIn.Application.ActiveExplorer();
            Outlook.Selection sel = (explorer != null) ? explorer.Selection : null;
            if (sel == null || sel.Count == 0)
            {
                MessageBox.Show("Select one or more messages first.", "Pushover Alerts");
                return;
            }

            string added = "";
            foreach (object o in sel)
            {
                Outlook.MailItem mail = o as Outlook.MailItem;
                if (mail == null) continue;

                string addr = Alerter.GetSenderSmtp(mail);
                if (string.IsNullOrEmpty(addr)) continue;

                string val = useDomain ? Alerter.SplitDomain(addr) : addr;
                if (string.IsNullOrEmpty(val)) continue;

                // New entries use the default alert type (blank -> resolved at send time).
                Alerter.AppendWatch(val, Alerter.DefaultAlertType);
                added += val + Environment.NewLine;
            }

            Confirm(added.Length > 0
                ? "Added to watch list:\r\n\r\n" + added +
                  "\r\nOpen 'Watch List' to assign a specific alert type."
                : "Nothing added (could not read sender address).");
        }

        private void Confirm(string message)
        {
            if (Alerter.ShowConfirmations)
                MessageBox.Show(message, "Pushover Alerts");
        }

        private static void ShowError(Exception ex)
        {
            MessageBox.Show(ex.Message, "Pushover Alerts - Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // ---- Read the embedded XML resource ----

        private static string GetResourceText(string resourceName)
        {
            System.Reflection.Assembly asm =
                System.Reflection.Assembly.GetExecutingAssembly();
            using (System.IO.Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                using (System.IO.StreamReader reader = new System.IO.StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }
    }
}
