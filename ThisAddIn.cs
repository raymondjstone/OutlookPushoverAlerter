using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookPushoverAlerter
{
    public partial class ThisAddIn
    {
        // Holds the Inbox.Items collection so the ItemAdd event stays alive.
        private Outlook.Items _inboxItems;
        private bool _inboxHooked;

        // The alerting engine, shared with the ribbon buttons.
        public PushoverAlerter Alerter { get; private set; } = new PushoverAlerter();

        // Called by FinishInitialization() (guaranteed path) and also by
        // ThisAddIn_Startup if that event happens to fire. Idempotent.
        internal void HookInbox()
        {
            if (_inboxHooked) return;
            Alerter.EnsureListFiles();
            Alerter.EnsureSettingsLoaded();
            try
            {
                Outlook.NameSpace ns = this.Application.GetNamespace("MAPI");
                Outlook.MAPIFolder inbox =
                    ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                _inboxItems = inbox.Items;
                _inboxItems.ItemAdd += InboxItems_ItemAdd;
                _inboxHooked = true;
            }
            catch
            {
                // If hooking fails, the ribbon buttons + Rescan still work.
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            HookInbox(); // no-op if FinishInitialization already called it
        }

        // Fires for every new item that lands in the Inbox.
        private void InboxItems_ItemAdd(object item)
        {
            Outlook.MailItem mail = item as Outlook.MailItem;
            if (mail == null) return;
            try { Alerter.Process(mail); }
            catch { /* never let a single message crash Outlook */ }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_inboxItems != null)
                _inboxItems.ItemAdd -= InboxItems_ItemAdd;
            _inboxItems = null;
            _inboxHooked = false;
        }

        // Registers the ribbon defined in Ribbon1.cs / Ribbon1.xml.
        protected override Microsoft.Office.Core.IRibbonExtensibility
            CreateRibbonExtensibilityObject()
        {
            return new Ribbon1(this);
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
