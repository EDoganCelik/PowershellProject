using System;
using System.Windows.Forms;

namespace GetKnownIssuesForm
{
    public partial class Form1 : Form
    {
        private KnownIssuesResult _lastResult;

        public Form1()
        {
            InitializeComponent();
            topPanel.BringToFront();
            browser.IsWebBrowserContextMenuEnabled = false;
            browser.WebBrowserShortcutsEnabled = false;
            browser.AllowWebBrowserDrop = false;
        }

        private void SetStatus(string text)
        {
            statusLabel.Text = text;
            Application.DoEvents();
        }

        private bool TryGetValidatedUrl(out string url)
        {
            url = txtUrl.Text.Trim();
            if (string.IsNullOrWhiteSpace(url))
            {
                MessageBox.Show("Lutfen bir URL girin.", "Uyari", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            try
            {
                url = SecurityHelper.ValidateAndNormalizeUrl(url);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Gecersiz URL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        private void btnFetch_Click(object sender, EventArgs e)
        {
            if (!TryGetValidatedUrl(out string url))
                return;

            btnFetch.Enabled = false;
            Cursor = Cursors.WaitCursor;
            SetStatus("Sayfa indiriliyor...");

            try
            {
                var result = KnownIssuesExtractor.Extract(url);
                _lastResult = result;

                SetStatus("Bulundu: " + result.Title + " -- onizleme yukleniyor...");
                browser.DocumentText = result.FullHtml;
                SetStatus("Tamam. '" + result.Title + "' icin Known Issues bolumu yuklendi.");
            }
            catch (Exception ex)
            {
                SetStatus("Hata olustu.");
                MessageBox.Show(ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                btnFetch.Enabled = true;
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (!TryGetValidatedUrl(out string url))
                return;

            Cursor = Cursors.WaitCursor;
            try
            {
                if (_lastResult == null || !string.Equals(_lastResult.Url, url, StringComparison.OrdinalIgnoreCase))
                {
                    SetStatus("Bolum konumu tespit ediliyor...");
                    var result = KnownIssuesExtractor.Extract(url);
                    _lastResult = result;
                    browser.DocumentText = result.FullHtml;
                }

                if (!string.IsNullOrEmpty(_lastResult.AnchorId))
                    SetStatus("Tarayicida aciliyor (dogrudan bolume atlanacak).");
                else
                    SetStatus("Sayfada 'id' bulunamadi, sayfa basindan aciliyor.");

                SecurityHelper.OpenUrlInBrowser(_lastResult.AnchorUrl);
            }
            catch (Exception ex)
            {
                SetStatus("Tarayici acilamadi.");
                MessageBox.Show(ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
    }
}
