using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookPushoverAlerter
{
    /// <summary>
    /// Editor for the Pushover credentials and the default alert type. The
    /// "Send test" button uses the values currently in the boxes so they can be
    /// verified before saving.
    /// </summary>
    public class PushoverSettingsForm : Form
    {
        private readonly PushoverAlerter _alerter;

        private TextBox _token;
        private TextBox _user;
        private ComboBox _defaultType;
        private Button _test;
        private Button _save;
        private Button _close;

        public PushoverSettingsForm(PushoverAlerter alerter)
        {
            _alerter = alerter;
            BuildUi();
            LoadValues();
        }

        private void BuildUi()
        {
            Text = "Pushover Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(460, 230);
            ShowInTaskbar = false;

            var help = new Label
            {
                Location = new Point(12, 10),
                Size = new Size(436, 34),
                Text = "Enter the API token of your Pushover application and your user key. " +
                       "Both are required to send notifications."
            };

            var lblToken = new Label { Location = new Point(12, 56), AutoSize = true, Text = "API token:" };
            _token = new TextBox { Location = new Point(120, 53), Size = new Size(328, 23) };

            var lblUser = new Label { Location = new Point(12, 90), AutoSize = true, Text = "User key:" };
            _user = new TextBox { Location = new Point(120, 87), Size = new Size(328, 23) };

            var lblDefault = new Label { Location = new Point(12, 124), AutoSize = true, Text = "Default type:" };
            _defaultType = new ComboBox
            {
                Location = new Point(120, 121),
                Size = new Size(220, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            _test = new Button { Location = new Point(12, 180), Size = new Size(110, 28), Text = "Send test" };
            _test.Click += (s, e) => SendTest();

            _save = new Button { Location = new Point(256, 180), Size = new Size(90, 28), Text = "Save" };
            _save.Click += (s, e) => { SaveValues(); Close(); };

            _close = new Button { Location = new Point(356, 180), Size = new Size(90, 28), Text = "Cancel" };
            _close.Click += (s, e) => Close();

            Controls.Add(help);
            Controls.Add(lblToken); Controls.Add(_token);
            Controls.Add(lblUser); Controls.Add(_user);
            Controls.Add(lblDefault); Controls.Add(_defaultType);
            Controls.Add(_test); Controls.Add(_save); Controls.Add(_close);

            AcceptButton = _save;
            CancelButton = _close;
        }

        private const string DefaultNone = "(none)";

        private void LoadValues()
        {
            _alerter.EnsureSettingsLoaded();
            _token.Text = _alerter.PushoverAppToken;
            _user.Text = _alerter.PushoverUserKey;

            _defaultType.Items.Clear();
            _defaultType.Items.Add(DefaultNone);
            foreach (AlertType t in _alerter.ReadAlertTypes())
                _defaultType.Items.Add(t.Name);

            int idx = string.IsNullOrEmpty(_alerter.DefaultAlertType)
                ? 0 : _defaultType.Items.IndexOf(_alerter.DefaultAlertType);
            _defaultType.SelectedIndex = idx >= 0 ? idx : 0;
        }

        private void ApplyToEngine()
        {
            _alerter.PushoverAppToken = (_token.Text ?? "").Trim();
            _alerter.PushoverUserKey = (_user.Text ?? "").Trim();
            string dt = _defaultType.SelectedItem as string ?? DefaultNone;
            _alerter.DefaultAlertType = (dt == DefaultNone) ? "" : dt;
        }

        private void SaveValues()
        {
            ApplyToEngine();
            _alerter.SaveSettings();
        }

        private void SendTest()
        {
            // Use whatever is currently typed, without committing it yet.
            ApplyToEngine();
            AlertType type = _alerter.ResolveAlertType(_alerter.DefaultAlertType);
            string err;
            bool ok = _alerter.SendTest(type, out err);
            if (ok)
                MessageBox.Show(this,
                    "Test notification sent (sound: " +
                    (string.IsNullOrEmpty(type.Sound) ? "default" : type.Sound) + ").",
                    "Pushover Settings");
            else
                MessageBox.Show(this, err, "Pushover Settings - Test failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
