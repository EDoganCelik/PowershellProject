using System;
using System.Drawing;
using System.Windows.Forms;

namespace RemoteServiceManager.Forms
{
    public class NameInputDialog : Form
    {
        private readonly TextBox _txt;

        public string EnteredName => _txt.Text.Trim();

        public NameInputDialog(string title = "Kaydı İsimlendir", string prompt = "Bu kayıt için bir isim girin:")
        {
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(380, 130);
            Font = new Font("Segoe UI", 9F);

            var lbl = new Label
            {
                Text = prompt,
                AutoSize = true,
                Location = new Point(12, 14)
            };

            _txt = new TextBox
            {
                Location = new Point(12, 40),
                Size = new Size(350, 23)
            };

            var btnOk = new Button
            {
                Text = "Kaydet",
                DialogResult = DialogResult.OK,
                Location = new Point(206, 80),
                Size = new Size(75, 28)
            };

            var btnCancel = new Button
            {
                Text = "İptal",
                DialogResult = DialogResult.Cancel,
                Location = new Point(287, 80),
                Size = new Size(75, 28)
            };

            AcceptButton = btnOk;
            CancelButton = btnCancel;
            Controls.AddRange(new Control[] { lbl, _txt, btnOk, btnCancel });
            Shown += (s, e) => _txt.Focus();
        }

        public static string Prompt(IWin32Window owner = null)
        {
            using (var dlg = new NameInputDialog())
            {
                if (dlg.ShowDialog(owner) == DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.EnteredName))
                    return dlg.EnteredName;
            }
            return null;
        }
    }
}
