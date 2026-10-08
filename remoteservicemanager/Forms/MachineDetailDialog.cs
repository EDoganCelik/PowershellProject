using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RemoteServiceManager.Models;
using RemoteServiceManager.Services;

namespace RemoteServiceManager.Forms
{
    public class MachineDetailDialog : Form
    {
        private readonly string _computerName;
        private readonly Func<IList<string>> _getServices;
        private readonly Func<string> _getFilePath;
        private readonly Func<string> _getRegPath;
        private readonly Func<string> _getRegValue;
        private readonly Func<MetricPreferences> _getServiceMetrics;
        private readonly Func<MetricPreferences> _getPatchMetrics;

        private readonly RemoteChecker _checker = new RemoteChecker();
        private readonly PatchChecker _patchChecker = new PatchChecker();
        private readonly ServiceFixer _fixer = new ServiceFixer();

        private HostResult _serviceResult;
        private HostResult _patchResult;

        private readonly TextBox _txtGeneral;
        private readonly TextBox _txtPatch;
        private readonly DataGridView _dgvServices;
        private readonly Label _lblStatus;
        private readonly Button _btnRefresh;
        private readonly Button _btnFix;
        private readonly ComboBox _cboCcm;

        /// <summary>Ana forma satır güncellemesi için.</summary>
        public event Action<HostResult, bool> ResultUpdated;

        public MachineDetailDialog(
            string computerName,
            HostResult existingService,
            HostResult existingPatch,
            Func<IList<string>> getServices,
            Func<string> getFilePath,
            Func<string> getRegPath,
            Func<string> getRegValue,
            Func<MetricPreferences> getServiceMetrics,
            Func<MetricPreferences> getPatchMetrics)
        {
            _computerName = computerName;
            _serviceResult = existingService;
            _patchResult = existingPatch;
            _getServices = getServices;
            _getFilePath = getFilePath;
            _getRegPath = getRegPath;
            _getRegValue = getRegValue;
            _getServiceMetrics = getServiceMetrics;
            _getPatchMetrics = getPatchMetrics;

            Text = computerName + " — Makine Detayı";
            Size = new Size(920, 640);
            MinimumSize = new Size(720, 500);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9F);

            var tabs = new TabControl { Dock = DockStyle.Fill };

            var tabGeneral = new TabPage("Genel Bilgi");
            _txtGeneral = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F),
                WordWrap = false
            };
            tabGeneral.Controls.Add(_txtGeneral);

            var tabPatch = new TabPage("Güncelleme / Patch");
            _txtPatch = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F),
                WordWrap = false
            };
            tabPatch.Controls.Add(_txtPatch);

            var tabSvc = new TabPage("Servisler");
            _dgvServices = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                BackgroundColor = Color.White
            };
            _dgvServices.Columns.Add("Name", "Servis");
            _dgvServices.Columns.Add("Status", "Durum");
            tabSvc.Controls.Add(_dgvServices);

            tabs.TabPages.Add(tabGeneral);
            tabs.TabPages.Add(tabPatch);
            tabs.TabPages.Add(tabSvc);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 56, Padding = new Padding(8) };
            _btnRefresh = new Button { Text = "Bilgileri Güncelle", Width = 140, Height = 32, Left = 8, Top = 12 };
            _btnFix = new Button { Text = "Durumu Düzelt", Width = 130, Height = 32, Left = 156, Top = 12 };
            var lblCcm = new Label { Text = "CCM Cycle:", AutoSize = true, Left = 300, Top = 18 };
            _cboCcm = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Left = 380,
                Top = 14,
                Width = 220
            };
            foreach (var kv in PatchChecker.CcmSchedules)
                _cboCcm.Items.Add(kv.Key);
            if (_cboCcm.Items.Count > 0) _cboCcm.SelectedIndex = 0;

            var btnCcm = new Button { Text = "CCM Gönder", Width = 110, Height = 32, Left = 610, Top = 12 };
            var btnClose = new Button { Text = "Kapat", Width = 90, Height = 32, Left = 800, Top = 12, Anchor = AnchorStyles.Top | AnchorStyles.Right, DialogResult = DialogResult.OK };

            _lblStatus = new Label
            {
                AutoSize = true,
                ForeColor = Color.FromArgb(20, 60, 120),
                Left = 8,
                Top = 0,
                Text = ""
            };

            _btnRefresh.Click += async (s, e) => await RefreshAsync();
            _btnFix.Click += async (s, e) => await FixAsync();
            btnCcm.Click += (s, e) => TriggerCcm();

            bottom.Controls.AddRange(new Control[] { _btnRefresh, _btnFix, lblCcm, _cboCcm, btnCcm, btnClose });

            var statusBar = new Panel { Dock = DockStyle.Bottom, Height = 24 };
            statusBar.Controls.Add(_lblStatus);

            Controls.Add(tabs);
            Controls.Add(bottom);
            Controls.Add(statusBar);

            RenderFromCache();
            AcceptButton = btnClose;
        }

        private void RenderFromCache()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Makine: " + _computerName);
            sb.AppendLine("Zaman : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine(new string('-', 80));

            var src = _serviceResult ?? _patchResult;
            if (src == null)
            {
                sb.AppendLine("(Henüz veri yok — Bilgileri Güncelle'ye basın)");
            }
            else
            {
                AppendLine(sb, "Online", src.OnlineStatus ?? src.IsOnline.ToString());
                AppendLine(sb, "OS Build", src.OsBuild);
                AppendLine(sb, "Uptime", src.Uptime);
                AppendLine(sb, "Last Boot", src.LastBootTime);
                AppendLine(sb, "Son Kapanma", src.LastShutdown);
                AppendLine(sb, "Oturum", src.LoggedInUser);
                AppendLine(sb, "Disk", src.DiskInfo);
                AppendLine(sb, "RAM", src.MemoryInfo);
                AppendLine(sb, "Ağ", src.NetworkInfo);
                AppendLine(sb, "Mavi Ekran", src.Bsod + (src.BsodFileCount != "N/A" ? " (dosya: " + src.BsodFileCount + ")" : ""));
                AppendLine(sb, "Bekleyen Update", src.PendingUpdates);
                AppendLine(sb, "Top Process", src.TopProcesses);
                AppendLine(sb, "Dosya", src.FilePathControl);
                AppendLine(sb, "Registry", src.RegistryValue);
                AppendLine(sb, "Hata", src.ErrorMessage);
                if (src.CustomValues != null)
                {
                    foreach (var kv in src.CustomValues)
                        AppendLine(sb, kv.Key, kv.Value);
                }
            }
            _txtGeneral.Text = sb.ToString();

            var patchSb = new StringBuilder();
            patchSb.AppendLine("Makine: " + _computerName);
            var patch = _patchResult ?? _serviceResult;
            if (patch != null)
            {
                patchSb.AppendLine("Özet : " + (string.IsNullOrEmpty(patch.UpdateSummary) || patch.UpdateSummary == "N/A"
                    ? patch.PendingUpdates : patch.UpdateSummary));
                patchSb.AppendLine("ccmexec : " + patch.CcmExecStatus);
                patchSb.AppendLine("wuauserv: " + patch.WuauservStatus);
                patchSb.AppendLine("bits    : " + patch.BitsStatus);
                patchSb.AppendLine("cryptsvc: " + patch.CryptSvcStatus);
                patchSb.AppendLine(new string('-', 80));
                if (patch.PatchDetailLines != null && patch.PatchDetailLines.Count > 0)
                    patchSb.AppendLine(string.Join(Environment.NewLine, patch.PatchDetailLines));
                else
                    patchSb.AppendLine("(Detay yok — Bilgileri Güncelle ile CCM sorgusu alın)");
            }
            else
            {
                patchSb.AppendLine("(Henüz patch verisi yok)");
            }
            _txtPatch.Text = patchSb.ToString();

            _dgvServices.Rows.Clear();
            var detail = _serviceResult?.ServiceDetail;
            if (detail != null)
            {
                foreach (var s in detail)
                    _dgvServices.Rows.Add(s.Name, s.Status);
            }
        }

        private static void AppendLine(StringBuilder sb, string label, string value)
        {
            if (string.IsNullOrWhiteSpace(value) || value == "N/A") return;
            sb.AppendLine(label.PadRight(16) + ": " + value);
        }

        private async Task RefreshAsync()
        {
            _btnRefresh.Enabled = false;
            _btnFix.Enabled = false;
            _lblStatus.Text = "Güncelleniyor: " + _computerName + " ...";
            try
            {
                var services = _getServices() ?? new List<string>();
                var svcPrefs = _getServiceMetrics() ?? MetricCatalog.DefaultServicePreferences();
                var patchPrefs = _getPatchMetrics() ?? MetricCatalog.DefaultPatchPreferences();

                // Patch detayları her zaman çekilsin (dialog için)
                EnsurePatchDetails(patchPrefs);

                var serviceTask = Task.Run(() => _checker.Check(
                    _computerName, services, _getFilePath(), _getRegPath(), _getRegValue(),
                    svcPrefs));
                var patchTask = Task.Run(() => _patchChecker.Check(_computerName, patchPrefs));

                _serviceResult = await serviceTask;
                _patchResult = await patchTask;

                // Patch detay satırlarını service sonucuna da taşı
                if (_patchResult?.PatchDetailLines != null && _patchResult.PatchDetailLines.Count > 0)
                {
                    _serviceResult.PatchDetailLines = _patchResult.PatchDetailLines;
                    if (string.IsNullOrEmpty(_serviceResult.PendingUpdates) || _serviceResult.PendingUpdates == "N/A")
                        _serviceResult.PendingUpdates = _patchResult.PendingUpdates;
                    if (string.IsNullOrEmpty(_serviceResult.UpdateSummary) || _serviceResult.UpdateSummary == "N/A")
                        _serviceResult.UpdateSummary = _patchResult.UpdateSummary;
                    if (string.IsNullOrEmpty(_serviceResult.Bsod) || _serviceResult.Bsod == "N/A")
                    {
                        _serviceResult.Bsod = _patchResult.Bsod;
                        _serviceResult.BsodFileCount = _patchResult.BsodFileCount;
                    }
                }

                RenderFromCache();
                ResultUpdated?.Invoke(_serviceResult, false);
                ResultUpdated?.Invoke(_patchResult, true);
                _lblStatus.Text = "Güncellendi: " + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _lblStatus.Text = "Hata.";
            }
            finally
            {
                _btnRefresh.Enabled = true;
                _btnFix.Enabled = true;
            }
        }

        private static void EnsurePatchDetails(MetricPreferences prefs)
        {
            if (prefs.EnabledBuiltInIds == null)
                prefs.EnabledBuiltInIds = new List<string>();
            if (!prefs.EnabledBuiltInIds.Contains(MetricCatalog.UpdateSummary))
                prefs.EnabledBuiltInIds.Add(MetricCatalog.UpdateSummary);
            if (!prefs.EnabledBuiltInIds.Contains(MetricCatalog.CcmExec))
                prefs.EnabledBuiltInIds.Add(MetricCatalog.CcmExec);
        }

        private async Task FixAsync()
        {
            if (_serviceResult?.ServiceDetail == null || _serviceResult.ServiceDetail.Count == 0)
            {
                MessageBox.Show(this, "Önce bilgileri güncelleyin veya servis listesi boş.", "Bilgi");
                return;
            }

            var toFix = _serviceResult.ServiceDetail
                .Where(s => s.Status != "Running" && s.Status != "NotFound")
                .ToList();
            if (toFix.Count == 0)
            {
                MessageBox.Show(this, "Düzeltilecek servis yok.", "Bilgi");
                return;
            }

            if (MessageBox.Show(this,
                    "Bu makinede " + toFix.Count + " servis başlatılacak. Devam?",
                    "Onay", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            _btnFix.Enabled = false;
            _lblStatus.Text = "Servisler düzeltiliyor...";
            try
            {
                var results = await Task.Run(() => _fixer.Fix(_computerName, _serviceResult.ServiceDetail));
                _serviceResult.ServiceDetail = results.Select(r => new ServiceStatusItem
                {
                    Name = r.ServiceName,
                    Status = r.SonrakiDurum
                }).ToList();
                _serviceResult.Running = string.Join(", ",
                    _serviceResult.ServiceDetail.Where(s => s.Status == "Running").Select(s => s.Name));
                _serviceResult.NotRunning = string.Join(", ",
                    _serviceResult.ServiceDetail.Where(s => s.Status != "Running")
                        .Select(s => s.Name + "(" + s.Status + ")"));

                // Patch servislerini de dene
                await Task.Run(() => _fixer.Fix(_computerName,
                    PatchChecker.PatchServices.Select(n => new ServiceStatusItem { Name = n, Status = "Stopped" }).ToList()));

                RenderFromCache();
                ResultUpdated?.Invoke(_serviceResult, false);
                _lblStatus.Text = "Düzeltme tamamlandı.";

                var summary = ServiceFixer.GetHostSummary(results);
                MessageBox.Show(this, summary, "Düzeltme Sonucu");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _btnFix.Enabled = true;
            }
        }

        private void TriggerCcm()
        {
            if (_cboCcm.SelectedItem == null) return;
            var name = _cboCcm.SelectedItem.ToString();
            string id;
            if (!PatchChecker.CcmSchedules.TryGetValue(name, out id)) return;

            _lblStatus.Text = "CCM gönderiliyor: " + name;
            var msg = _patchChecker.TriggerCcmSchedule(_computerName, id);
            _lblStatus.Text = "CCM: " + msg;
            MessageBox.Show(this, _computerName + ": " + msg, "CCM Cycle: " + name);
        }
    }
}
