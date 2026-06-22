using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OutlookPushoverAlerter
{
    /// <summary>
    /// Manager for the watch list. Each row is an email address or domain plus
    /// the alert type that fires when mail arrives from it. Changes are written
    /// back through <see cref="PushoverAlerter.SaveWatchList"/>.
    /// </summary>
    public class WatchListForm : Form
    {
        private readonly PushoverAlerter _alerter;
        private readonly List<WatchEntry> _entries;

        private ListView _list;
        private TextBox _pattern;
        private ComboBox _type;
        private Button _addUpdate;
        private Button _remove;
        private Button _close;

        public WatchListForm(PushoverAlerter alerter)
        {
            _alerter = alerter;
            _entries = _alerter.ReadWatchEntries();
            BuildUi();
            RefreshList();
        }

        private void BuildUi()
        {
            Text = "Watch List";
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(480, 380);
            Size = new Size(520, 440);
            ShowInTaskbar = false;

            var help = new Label
            {
                Dock = DockStyle.Top,
                Height = 50,
                Padding = new Padding(10, 8, 10, 0),
                Text = "Mail from any of these senders triggers a Pushover alert.\n" +
                       "Enter a full address (boss@acme.com) or a domain (acme.com), pick an " +
                       "alert type, then Add / Update."
            };

            _list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = true,
                HideSelection = false
            };
            _list.Columns.Add("Sender", 260);
            _list.Columns.Add("Alert type", 200);
            _list.SelectedIndexChanged += (s, e) => LoadSelectionIntoEditor();

            // Bottom editor: pattern + type combo + buttons.
            var bottom = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 96,
                ColumnCount = 3,
                RowCount = 2,
                Padding = new Padding(10, 6, 10, 8)
            };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _pattern = new TextBox { Dock = DockStyle.Fill };
            _pattern.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AddOrUpdate(); }
            };

            _type = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 160,
                Margin = new Padding(6, 0, 0, 0)
            };
            ReloadTypeChoices();

            _addUpdate = new Button { Text = "Add / Update", AutoSize = true, Margin = new Padding(6, 0, 0, 0) };
            _addUpdate.Click += (s, e) => AddOrUpdate();

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true
            };
            _close = new Button { Text = "Close", AutoSize = true };
            _close.Click += (s, e) => Close();
            _remove = new Button { Text = "Remove selected", AutoSize = true };
            _remove.Click += (s, e) => RemoveSelected();
            actions.Controls.Add(_close);
            actions.Controls.Add(_remove);

            bottom.Controls.Add(_pattern, 0, 0);
            bottom.Controls.Add(_type, 1, 0);
            bottom.Controls.Add(_addUpdate, 2, 0);
            bottom.Controls.Add(actions, 0, 1);
            bottom.SetColumnSpan(actions, 3);

            Controls.Add(_list);
            Controls.Add(bottom);
            Controls.Add(help);

            AcceptButton = _addUpdate;
            CancelButton = _close;
        }

        private const string DefaultChoice = "(default)";

        private void ReloadTypeChoices()
        {
            string current = _type.SelectedItem as string;
            _type.Items.Clear();
            _type.Items.Add(DefaultChoice);
            foreach (AlertType t in _alerter.ReadAlertTypes())
                _type.Items.Add(t.Name);
            _type.SelectedIndex = 0;
            if (current != null)
            {
                int idx = _type.Items.IndexOf(current);
                if (idx >= 0) _type.SelectedIndex = idx;
            }
        }

        private void RefreshList()
        {
            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (WatchEntry e in _entries)
            {
                var item = new ListViewItem(e.Pattern);
                item.SubItems.Add(string.IsNullOrEmpty(e.AlertType) ? DefaultChoice : e.AlertType);
                _list.Items.Add(item);
            }
            _list.EndUpdate();
        }

        private void LoadSelectionIntoEditor()
        {
            if (_list.SelectedItems.Count != 1) return;
            ListViewItem item = _list.SelectedItems[0];
            _pattern.Text = item.Text;

            string typeName = item.SubItems[1].Text;
            int idx = _type.Items.IndexOf(typeName);
            _type.SelectedIndex = idx >= 0 ? idx : 0;
        }

        private void AddOrUpdate()
        {
            string val = (_pattern.Text ?? "").Trim().ToLower();
            if (val.StartsWith("@")) val = val.Substring(1);
            if (val.Length == 0) return;

            bool looksLikeEmail = val.Contains("@") &&
                                  val.IndexOf('@') > 0 &&
                                  val.IndexOf('@') < val.Length - 1;
            bool looksLikeDomain = !val.Contains("@") && val.Contains(".");
            if (!looksLikeEmail && !looksLikeDomain)
            {
                MessageBox.Show(this,
                    "Enter a full email address (name@acme.com) or a domain (acme.com).",
                    "Watch List", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string type = _type.SelectedItem as string ?? DefaultChoice;
            if (type == DefaultChoice) type = "";

            WatchEntry existing = _entries
                .FirstOrDefault(x => string.Equals(x.Pattern, val, StringComparison.OrdinalIgnoreCase));
            if (existing != null) existing.AlertType = type;
            else _entries.Add(new WatchEntry(val, type));

            Save();
            RefreshList();
            _pattern.Clear();
            _pattern.Focus();
        }

        private void RemoveSelected()
        {
            if (_list.SelectedItems.Count == 0) return;
            var patterns = _list.SelectedItems.Cast<ListViewItem>()
                .Select(i => i.Text).ToList();
            _entries.RemoveAll(e => patterns.Contains(e.Pattern, StringComparer.OrdinalIgnoreCase));
            Save();
            RefreshList();
        }

        private void Save()
        {
            _alerter.SaveWatchList(_entries);
        }
    }
}
