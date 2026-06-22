using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OutlookPushoverAlerter
{
    /// <summary>
    /// Manager for custom Pushover sounds. Each row is the exact name of a sound
    /// uploaded to your Pushover application. Changes are written back through
    /// <see cref="PushoverAlerter.SaveCustomSounds"/>. The built-in sounds are
    /// shown for reference but cannot be edited or removed here.
    /// </summary>
    public class CustomSoundsForm : Form
    {
        private readonly PushoverAlerter _alerter;
        private readonly List<string> _sounds;

        private ListBox _list;
        private TextBox _name;
        private Button _add;
        private Button _test;
        private Button _remove;
        private Button _close;

        public CustomSoundsForm(PushoverAlerter alerter)
        {
            _alerter = alerter;
            _sounds = _alerter.ReadCustomSounds();
            BuildUi();
            RefreshList();
        }

        private void BuildUi()
        {
            Text = "Custom Sounds";
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(420, 340);
            Size = new Size(460, 400);
            ShowInTaskbar = false;

            var help = new Label
            {
                Dock = DockStyle.Top,
                Height = 58,
                Padding = new Padding(10, 8, 10, 0),
                Text = "Custom sounds you've uploaded to your Pushover app " +
                       "(pushover.net -> Your Applications -> sounds).\n" +
                       "Enter the exact name it was given at upload (lowercase, no spaces), " +
                       "then Add. These appear in the Alert Types sound list."
            };

            _list = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false
            };
            _list.DoubleClick += (s, e) => LoadSelectionIntoEditor();
            _list.SelectedIndexChanged += (s, e) => LoadSelectionIntoEditor();

            // Bottom editor: name + Add, plus the action row.
            var bottom = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 80,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(10, 6, 10, 8)
            };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _name = new TextBox { Dock = DockStyle.Fill };
            _name.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AddSound(); }
            };

            _add = new Button { Text = "Add", AutoSize = true, Margin = new Padding(6, 0, 0, 0) };
            _add.Click += (s, e) => AddSound();

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
            _test = new Button { Text = "Test", AutoSize = true };
            _test.Click += (s, e) => TestSelected();
            actions.Controls.Add(_close);
            actions.Controls.Add(_remove);
            actions.Controls.Add(_test);

            bottom.Controls.Add(_name, 0, 0);
            bottom.Controls.Add(_add, 1, 0);
            bottom.Controls.Add(actions, 0, 1);
            bottom.SetColumnSpan(actions, 2);

            Controls.Add(_list);
            Controls.Add(bottom);
            Controls.Add(help);

            AcceptButton = _add;
            CancelButton = _close;
        }

        private void RefreshList()
        {
            string keep = _list.SelectedItem as string;
            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (string s in _sounds) _list.Items.Add(s);
            _list.EndUpdate();
            if (keep != null)
            {
                int idx = _list.Items.IndexOf(keep);
                if (idx >= 0) _list.SelectedIndex = idx;
            }
        }

        private void LoadSelectionIntoEditor()
        {
            if (_list.SelectedItem is string s) _name.Text = s;
        }

        private void AddSound()
        {
            string name = PushoverAlerter.NormalizeSoundName(_name.Text ?? "");
            if (name.Length == 0)
            {
                MessageBox.Show(this,
                    "Enter the name of a sound uploaded to your Pushover app.",
                    "Custom Sounds", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (PushoverAlerter.Sounds.Any(b => string.Equals(b, name, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(this,
                    "'" + name + "' is a built-in Pushover sound and is already available.",
                    "Custom Sounds", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _name.Clear();
                _name.Focus();
                return;
            }

            if (!_sounds.Any(s => string.Equals(s, name, StringComparison.OrdinalIgnoreCase)))
                _sounds.Add(name);

            Save();
            RefreshList();
            int idx = _list.Items.IndexOf(name);
            if (idx >= 0) _list.SelectedIndex = idx;
            _name.Clear();
            _name.Focus();
        }

        /// <summary>
        /// Sends a real Pushover notification using the selected (or typed) sound so
        /// the name can be confirmed against your app/account.
        /// </summary>
        private void TestSelected()
        {
            string name = PushoverAlerter.NormalizeSoundName(_name.Text ?? "");
            if (name.Length == 0) name = _list.SelectedItem as string ?? "";
            if (name.Length == 0)
            {
                MessageBox.Show(this, "Select or type a sound to test.", "Custom Sounds",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var type = new AlertType("Custom Sound Test", name,
                "Testing the '" + name + "' Pushover sound.");

            Cursor.Current = Cursors.WaitCursor;
            string err;
            bool ok;
            try { ok = _alerter.SendTest(type, out err); }
            finally { Cursor.Current = Cursors.Default; }

            if (ok)
                MessageBox.Show(this,
                    "Test notification sent.\r\n\r\nSound: " + name +
                    "\r\n\r\nIf you don't hear this sound on your device, the name may not match a" +
                    " sound uploaded to your Pushover app, or your device is set to override it.",
                    "Custom Sounds", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(this,
                    "Test failed:\r\n\r\n" + err,
                    "Custom Sounds", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void RemoveSelected()
        {
            if (!(_list.SelectedItem is string name)) return;
            _sounds.RemoveAll(s => string.Equals(s, name, StringComparison.OrdinalIgnoreCase));
            Save();
            RefreshList();
            _name.Clear();
        }

        private void Save()
        {
            _alerter.SaveCustomSounds(_sounds);
        }
    }
}
