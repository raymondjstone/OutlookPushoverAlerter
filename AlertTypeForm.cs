using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OutlookPushoverAlerter
{
    /// <summary>
    /// Manager for alert types. Each type has a name, a Pushover sound, and the
    /// text to send. Changes are written back through
    /// <see cref="PushoverAlerter.SaveAlertTypes"/>.
    /// </summary>
    public class AlertTypeForm : Form
    {
        private readonly PushoverAlerter _alerter;
        private readonly List<AlertType> _types;

        private ListBox _list;
        private TextBox _name;
        private ComboBox _sound;
        private TextBox _text;
        private Button _newBtn;
        private Button _save;
        private Button _remove;
        private Button _close;

        public AlertTypeForm(PushoverAlerter alerter)
        {
            _alerter = alerter;
            _types = _alerter.ReadAlertTypes();
            BuildUi();
            RefreshList();
        }

        private void BuildUi()
        {
            Text = "Alert Types";
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(560, 380);
            Size = new Size(600, 420);
            ShowInTaskbar = false;

            // Left: list of type names.
            _list = new ListBox
            {
                Dock = DockStyle.Left,
                Width = 200,
                IntegralHeight = false
            };
            _list.SelectedIndexChanged += (s, e) => LoadSelection();

            // Right: editor panel.
            var editor = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 4,
                Padding = new Padding(10, 8, 10, 8)
            };
            editor.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            editor.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            editor.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            editor.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            editor.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            editor.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            _name = new TextBox { Dock = DockStyle.Fill };
            // Editable so a custom Pushover sound name can simply be typed in; the
            // dropdown lists the built-ins plus any custom names added so far.
            _sound = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDown
            };
            ReloadSounds();
            if (_sound.Items.Count > 0) _sound.SelectedIndex = 0;

            _text = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true
            };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true
            };
            _close = new Button { Text = "Close", AutoSize = true };
            _close.Click += (s, e) => Close();
            _remove = new Button { Text = "Remove", AutoSize = true };
            _remove.Click += (s, e) => RemoveSelected();
            _newBtn = new Button { Text = "New", AutoSize = true };
            _newBtn.Click += (s, e) => ClearEditor();
            _save = new Button { Text = "Save", AutoSize = true };
            _save.Click += (s, e) => SaveCurrent();
            buttons.Controls.Add(_close);
            buttons.Controls.Add(_remove);
            buttons.Controls.Add(_newBtn);
            buttons.Controls.Add(_save);

            editor.Controls.Add(new Label { Text = "Name:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0) }, 0, 0);
            editor.Controls.Add(_name, 1, 0);
            editor.Controls.Add(new Label { Text = "Sound:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0) }, 0, 1);
            editor.Controls.Add(_sound, 1, 1);
            editor.Controls.Add(new Label { Text = "Text:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0) }, 0, 2);
            editor.Controls.Add(_text, 1, 2);
            editor.Controls.Add(buttons, 1, 3);

            Controls.Add(editor);
            Controls.Add(_list);
        }

        private void RefreshList()
        {
            string keep = _list.SelectedItem as string;
            _list.Items.Clear();
            foreach (AlertType t in _types) _list.Items.Add(t.Name);
            if (keep != null)
            {
                int idx = _list.Items.IndexOf(keep);
                if (idx >= 0) _list.SelectedIndex = idx;
            }
        }

        private void LoadSelection()
        {
            if (_list.SelectedIndex < 0 || _list.SelectedIndex >= _types.Count) return;
            AlertType t = _types[_list.SelectedIndex];
            _name.Text = t.Name;
            SelectSound(t.Sound);
            _text.Text = t.Text;
        }

        private void ReloadSounds()
        {
            string keep = _sound.Text;
            _sound.Items.Clear();
            foreach (string s in _alerter.AllSounds()) _sound.Items.Add(s);
            if (!string.IsNullOrEmpty(keep)) _sound.Text = keep;
        }

        private void SelectSound(string sound)
        {
            // Editable combo: set the text directly so custom names that aren't in
            // the list still display. Blank means "use the Pushover default sound".
            _sound.Text = string.IsNullOrEmpty(sound) ? "pushover" : sound;
        }

        private void ClearEditor()
        {
            _list.ClearSelected();
            _name.Clear();
            SelectSound("pushover");
            _text.Clear();
            _name.Focus();
        }

        private void SaveCurrent()
        {
            string name = (_name.Text ?? "").Trim();
            if (name.Length == 0)
            {
                MessageBox.Show(this, "Give the alert type a name.", "Alert Types",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // The sound box is free-text: normalize it, and remember any new custom
            // name so it shows up in the dropdown next time.
            string sound = PushoverAlerter.NormalizeSoundName(_sound.Text ?? "");
            if (sound.Length > 0)
            {
                _alerter.AddCustomSound(sound);
                ReloadSounds();
                _sound.Text = sound;
            }
            string text = _text.Text ?? "";

            AlertType existing = _types
                .FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.Sound = sound;
                existing.Text = text;
            }
            else
            {
                _types.Add(new AlertType(name, sound, text));
            }

            _alerter.SaveAlertTypes(_types);
            RefreshList();
            int idx = _list.Items.IndexOf(name);
            if (idx >= 0) _list.SelectedIndex = idx;
        }

        private void RemoveSelected()
        {
            if (_list.SelectedIndex < 0 || _list.SelectedIndex >= _types.Count) return;
            _types.RemoveAt(_list.SelectedIndex);
            _alerter.SaveAlertTypes(_types);
            RefreshList();
            ClearEditor();
        }
    }
}
