using System;
using System.Drawing;
using System.Windows.Forms;

namespace RemoteServiceManager.Forms
{
    /// <summary>
    /// Export/Import için şifre girişi. Export'ta onay alanı da gösterilir.
    /// </summary>
    public class PasswordInputDialog : Form
    {
        private readonly TextBox _txtPassword;
        private readonly TextBox _txtConfirm;
        private readonly bool _requireConfirm;

        public string Password => _txtPassword.Text;

        public PasswordInputDialog(string title, string prompt, bool requireConfirm)
        {
            _requireConfirm = requireConfirm;
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(400, requireConfirm ? 170 : 130);
            Font = new Font("Segoe UI", 9F);

            var lbl = new Label
            {
                Text = prompt,
                AutoSize = false,
                Location = new Point(12, 12),
                Size = new Size(370, 36)
            };

            var lblPass = new Label
            {
                Text = "Şifre:",
                AutoSize = true,
                Location = new Point(12, 52)
            };

            _txtPassword = new TextBox
            {
                Location = new Point(100, 48),
                Size = new Size(280, 23),
                UseSystemPasswordChar = true
            };

            Controls.Add(lbl);
            Controls.Add(lblPass);
            Controls.Add(_txtPassword);

            int btnY = 90;
            if (requireConfirm)
            {
                var lblConfirm = new Label
                {
                    Text = "Tekrar:",
                    AutoSize = true,
                    Location = new Point(12, 82)
                };
                _txtConfirm = new TextBox
                {
                    Location = new Point(100, 78),
                    Size = new Size(280, 23),
                    UseSystemPasswordChar = true
                };
                Controls.Add(lblConfirm);
                Controls.Add(_txtConfirm);
                btnY = 120;
            }
            else
            {
                _txtConfirm = null;
            }

            var btnOk = new Button
            {
                Text = "Tamam",
                Location = new Point(224, btnY),
                Size = new Size(75, 28)
            };
            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(_txtPassword.Text))
                {
                    MessageBox.Show(this, "Şifre boş olamaz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (_requireConfirm && _txtPassword.Text != _txtConfirm.Text)
                {
                    MessageBox.Show(this, "Şifreler eşleşmiyor.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (_txtPassword.Text.Length < 4)
                {
                    MessageBox.Show(this, "Şifre en az 4 karakter olmalıdır.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                DialogResult = DialogResult.OK;
                Close();
            };

            var btnCancel = new Button
            {
                Text = "İptal",
                DialogResult = DialogResult.Cancel,
                Location = new Point(305, btnY),
                Size = new Size(75, 28)
            };

            AcceptButton = btnOk;
            CancelButton = btnCancel;
            Controls.Add(btnOk);
            Controls.Add(btnCancel);
            Shown += (s, e) => _txtPassword.Focus();
        }

        public static string Prompt(IWin32Window owner, string title, string prompt, bool requireConfirm)
        {
            using (var dlg = new PasswordInputDialog(title, prompt, requireConfirm))
            {
                if (dlg.ShowDialog(owner) == DialogResult.OK)
                    return dlg.Password;
            }
            return null;
        }
    }
}
