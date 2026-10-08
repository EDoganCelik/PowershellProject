using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using RemoteServiceManager.Forms;
using RemoteServiceManager.Models;
using RemoteServiceManager.Services;

namespace RemoteServiceManager
{
    public partial class MainForm : Form
    {
        private readonly ConfigStore _store = new ConfigStore();
        private readonly RemoteChecker _checker = new RemoteChecker();
        private readonly PatchChecker _patchChecker = new PatchChecker();
        private readonly ServiceFixer _fixer = new ServiceFixer();

        private MetricPreferences _serviceMetrics = MetricCatalog.DefaultServicePreferences();
        private MetricPreferences _patchMetrics = MetricCatalog.DefaultPatchPreferences();

        private DataTable _serviceTable;
        private DataTable _patchTable;
        private readonly BindingSource _serviceBinding = new BindingSource();
        private readonly BindingSource _patchBinding = new BindingSource();
        private readonly Dictionary<string, HostResult> _serviceResults =
            new Dictionary<string, HostResult>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, HostResult> _patchResults =
            new Dictionary<string, HostResult>(StringComparer.OrdinalIgnoreCase);

        private CancellationTokenSource _cts;
        private bool _suppressMetricEvents;
        private bool _suppressPanelEvents;
        private const int DefaultMaxParallel = 10;

        // Service tab controls
        private CheckedListBox _lstServiceMetrics;
        private TextBox _txtServices, _txtFilePath, _txtRegPath, _txtRegValue, _txtServiceFilter;
        private DataGridView _dgvService;
        private Button _btnServiceCheck, _btnServiceFix, _btnServiceCancel, _btnServiceSave, _btnServiceUpdate, _btnServiceDelete, _btnServiceExport;
        private Button _btnApplyServiceCols;
        private ComboBox _cboHistory;

        // Patch tab controls
        private CheckedListBox _lstPatchMetrics;
        private TextBox _txtPatchFilter;
        private DataGridView _dgvPatch;
        private Button _btnPatchRefresh, _btnPatchRefreshSelected, _btnPatchFix, _btnPatchCancel, _btnPatchExport;
        private Button _btnApplyPatchCols;

        public MainForm()
        {
            InitializeComponent();
            BuildServiceTab();
            BuildPatchTab();
        }

        private void BuildServiceTab()
        {
            tabService.AutoScroll = true;

            var top = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 210,
                ColumnCount = 3,
                RowCount = 1,
                Padding = new Padding(0, 0, 0, 4)
            };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28F));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 44F));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28F));

            var left = new GroupBox { Text = "Toplanacak veriler (işaretsiz = sütun yok / sorgu yok)", Dock = DockStyle.Fill, Padding = new Padding(6) };
            _lstServiceMetrics = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                IntegralHeight = false
            };
            _lstServiceMetrics.ItemCheck += (s, e) =>
            {
                if (_suppressMetricEvents) return;
                BeginInvoke(new Action(ApplyServiceMetricSelection));
            };
            var metricBtns = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 32, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            _btnApplyServiceCols = new Button { Text = "Sütunları Uygula", Width = 120, Height = 26 };
            _btnApplyServiceCols.Click += (s, e) => { ApplyServiceMetricSelection(); RebuildServiceGrid(true); };
            metricBtns.Controls.Add(_btnApplyServiceCols);
            left.Controls.Add(_lstServiceMetrics);
            left.Controls.Add(metricBtns);

            var mid = new GroupBox { Text = "Parametreler", Dock = DockStyle.Fill, Padding = new Padding(8, 8, 8, 4) };
            mid.Controls.Add(MakeLabel("Servisler:", 10, 22));
            _txtServices = MakeAnchoredTextBox(mid, 120, 18, "WiaRpc, WinRM");
            mid.Controls.Add(MakeLabel("Dosya yolu:", 10, 50));
            _txtFilePath = MakeAnchoredTextBox(mid, 120, 46);
            mid.Controls.Add(MakeLabel("Registry yolu:", 10, 78));
            _txtRegPath = MakeAnchoredTextBox(mid, 120, 74);
            mid.Controls.Add(MakeLabel("Registry value:", 10, 106));
            _txtRegValue = MakeAnchoredTextBox(mid, 120, 102);
            mid.Controls.Add(MakeLabel("Filtre:", 10, 138));
            _txtServiceFilter = MakeAnchoredTextBox(mid, 120, 134);
            _txtServiceFilter.TextChanged += (s, e) => ApplyFilter(_serviceBinding, _txtServiceFilter.Text);
            mid.Controls.Add(MakeLabel("Profil:", 10, 168));
            _cboHistory = new ComboBox
            {
                Location = new Point(120, 164),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _cboHistory.SelectionChangeCommitted += (s, e) => LoadHistoryItem(_cboHistory.SelectedIndex);
            mid.Controls.AddRange(new Control[] { _txtServices, _txtFilePath, _txtRegPath, _txtRegValue, _txtServiceFilter, _cboHistory });
            mid.Resize += (s, e) => LayoutLabeledInputs(mid, 120);

            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 5,
                Padding = new Padding(4)
            };
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            for (var i = 0; i < 5; i++)
                right.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));

            _btnServiceCheck = MakeDockButton("Kontrol Et");
            _btnServiceFix = MakeDockButton("Durumu Düzelt");
            _btnServiceFix.Enabled = false;
            _btnServiceCancel = MakeDockButton("İptal");
            _btnServiceCancel.Enabled = false;
            _btnServiceSave = MakeDockButton("Profil Kaydet");
            _btnServiceUpdate = MakeDockButton("Profil Güncelle");
            _btnServiceDelete = MakeDockButton("Profil Sil");
            _btnServiceExport = MakeDockButton("Excel / CSV");
            var hint = new Label
            {
                Text = "Çift tık → makine detayı",
                Dock = DockStyle.Fill,
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.TopLeft
            };
            _btnServiceCheck.Click += async (s, e) => await RunServiceCheckAsync(ParseComputers());
            _btnServiceFix.Click += async (s, e) => await RunServiceFixAsync(null);
            _btnServiceCancel.Click += (s, e) => { try { _cts?.Cancel(); } catch { } };
            _btnServiceSave.Click += (s, e) => SaveProfile();
            _btnServiceUpdate.Click += (s, e) => UpdateProfile();
            _btnServiceDelete.Click += (s, e) => DeleteProfile();
            _btnServiceExport.Click += (s, e) => ExportGrid(_serviceBinding, _serviceTable, "ServisKontrol");
            right.Controls.Add(_btnServiceCheck, 0, 0);
            right.Controls.Add(_btnServiceFix, 1, 0);
            right.Controls.Add(_btnServiceCancel, 0, 1);
            right.Controls.Add(_btnServiceSave, 1, 1);
            right.Controls.Add(_btnServiceUpdate, 0, 2);
            right.Controls.Add(_btnServiceDelete, 1, 2);
            right.Controls.Add(_btnServiceExport, 0, 3);
            right.SetColumnSpan(hint, 2);
            right.Controls.Add(hint, 0, 4);

            top.Controls.Add(left, 0, 0);
            top.Controls.Add(mid, 1, 0);
            top.Controls.Add(right, 2, 0);

            _dgvService = CreateGrid();
            _dgvService.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                GridBuilder.ApplyRowColors(e, _dgvService.Rows[e.RowIndex].DataBoundItem as DataRowView, false);
            };
            _dgvService.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                OpenMachineDetail(GetComputerFromGrid(_dgvService, e.RowIndex));
            };

            tabService.Controls.Add(_dgvService);
            tabService.Controls.Add(top);
        }

        private void BuildPatchTab()
        {
            tabPatch.AutoScroll = true;

            var top = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 200,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(0, 0, 0, 4)
            };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38F));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62F));

            var left = new GroupBox { Text = "Patch metrikleri (işaretsiz = sütun/sorgu yok)", Dock = DockStyle.Fill, Padding = new Padding(6) };
            _lstPatchMetrics = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, IntegralHeight = false };
            _lstPatchMetrics.ItemCheck += (s, e) =>
            {
                if (_suppressMetricEvents) return;
                BeginInvoke(new Action(ApplyPatchMetricSelection));
            };
            var metricBtns = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 32 };
            _btnApplyPatchCols = new Button { Text = "Sütunları Uygula", Width = 120, Height = 26 };
            _btnApplyPatchCols.Click += (s, e) => { ApplyPatchMetricSelection(); RebuildPatchGrid(true); };
            metricBtns.Controls.Add(_btnApplyPatchCols);
            left.Controls.Add(_lstPatchMetrics);
            left.Controls.Add(metricBtns);

            var mid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 4,
                Padding = new Padding(4)
            };
            for (var i = 0; i < 3; i++)
                mid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
            mid.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            mid.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            mid.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            mid.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var filterRow = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 2, 0, 0) };
            var lblFilter = MakeLabel("Filtre:", 0, 6);
            _txtPatchFilter = new TextBox
            {
                Location = new Point(50, 2),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _txtPatchFilter.TextChanged += (s, e) => ApplyFilter(_patchBinding, _txtPatchFilter.Text);
            filterRow.Controls.Add(lblFilter);
            filterRow.Controls.Add(_txtPatchFilter);
            filterRow.Resize += (s, e) => _txtPatchFilter.Width = Math.Max(80, filterRow.ClientSize.Width - 54);
            mid.Controls.Add(filterRow, 0, 0);
            mid.SetColumnSpan(filterRow, 3);

            _btnPatchRefresh = MakeDockButton("Tümünü Yenile");
            _btnPatchRefreshSelected = MakeDockButton("Seçileni Yenile");
            _btnPatchFix = MakeDockButton("Seçili Servisleri Düzelt");
            _btnPatchCancel = MakeDockButton("İptal");
            _btnPatchCancel.Enabled = false;
            _btnPatchExport = MakeDockButton("Excel / CSV");
            var hint = new Label
            {
                Text = "Servis ekranından bağımsızdır.\nÇift tık → makine detayı + güncelleme listesi.\nCCM Minidump BSOD: >2 dosya = VAR",
                Dock = DockStyle.Fill,
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.TopLeft
            };
            _btnPatchRefresh.Click += async (s, e) => await RunPatchCheckAsync(ParseComputers(), true);
            _btnPatchRefreshSelected.Click += async (s, e) =>
            {
                var selected = GetSelectedComputers(_dgvPatch);
                if (selected.Count == 0)
                {
                    MessageBox.Show(this, "Tabloda satır seçin.", "Bilgi");
                    return;
                }
                await RunPatchCheckAsync(selected, false);
            };
            _btnPatchFix.Click += async (s, e) => await RunPatchServiceFixAsync();
            _btnPatchCancel.Click += (s, e) => { try { _cts?.Cancel(); } catch { } };
            _btnPatchExport.Click += (s, e) => ExportGrid(_patchBinding, _patchTable, "PatchTakip");
            mid.Controls.Add(_btnPatchRefresh, 0, 1);
            mid.Controls.Add(_btnPatchRefreshSelected, 1, 1);
            mid.Controls.Add(_btnPatchFix, 2, 1);
            mid.Controls.Add(_btnPatchCancel, 0, 2);
            mid.Controls.Add(_btnPatchExport, 1, 2);
            mid.SetColumnSpan(hint, 3);
            mid.Controls.Add(hint, 0, 3);

            top.Controls.Add(left, 0, 0);
            top.Controls.Add(mid, 1, 0);

            _dgvPatch = CreateGrid();
            _dgvPatch.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                GridBuilder.ApplyRowColors(e, _dgvPatch.Rows[e.RowIndex].DataBoundItem as DataRowView, true);
            };
            _dgvPatch.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                OpenMachineDetail(GetComputerFromGrid(_dgvPatch, e.RowIndex));
            };

            tabPatch.Controls.Add(_dgvPatch);
            tabPatch.Controls.Add(top);
        }

        private static Label MakeLabel(string text, int x, int y)
        {
            return new Label { Text = text, Location = new Point(x, y), AutoSize = true };
        }

        private static TextBox MakeAnchoredTextBox(Control parent, int left, int top, string text = null)
        {
            var tb = new TextBox
            {
                Location = new Point(left, top),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Width = Math.Max(80, parent.ClientSize.Width - left - 10)
            };
            if (text != null) tb.Text = text;
            return tb;
        }

        private static Button MakeDockButton(string text)
        {
            return new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                Margin = new Padding(2),
                MinimumSize = new Size(80, 28)
            };
        }

        private static void LayoutLabeledInputs(Control parent, int inputLeft)
        {
            var width = Math.Max(80, parent.ClientSize.Width - inputLeft - 10);
            foreach (Control c in parent.Controls)
            {
                if (c is TextBox || c is ComboBox)
                {
                    if ((c.Anchor & AnchorStyles.Right) == AnchorStyles.Right)
                        c.Width = width;
                }
            }
        }

        private static DataGridView CreateGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AllowUserToResizeColumns = true,
                AllowUserToOrderColumns = true,
                ScrollBars = ScrollBars.Both,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText,
                BackgroundColor = Color.White,
                StandardTab = true
            };
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                _store.LoadOrCreate();
                ApplyLoadedPayload();
                lblStatus.Text = "Hazır. " + _store.ConfigPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Yapılandırma: " + ex.Message, "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                try
                {
                    if (System.IO.File.Exists(_store.ConfigPath))
                        System.IO.File.Move(_store.ConfigPath, _store.ConfigPath + ".broken_" + DateTime.Now.ToString("yyyyMMddHHmmss"));
                    _store.LoadOrCreate();
                    ApplyLoadedPayload();
                }
                catch (Exception ex2)
                {
                    MessageBox.Show(this, ex2.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ApplyLoadedPayload()
        {
            if (_store.Payload.ServiceMetrics != null)
                _serviceMetrics = _store.Payload.ServiceMetrics;
            else if (_store.Payload.DefaultOptions != null)
                _serviceMetrics = MetricCatalog.FromLegacy(_store.Payload.DefaultOptions);

            if (_store.Payload.PatchMetrics != null)
                _patchMetrics = _store.Payload.PatchMetrics;

            // Dinamik özel metrik ekleme kaldırıldı; eski kayıtlar da gösterilmesin.
            if (_serviceMetrics.CustomMetrics != null)
                _serviceMetrics.CustomMetrics.Clear();
            if (_patchMetrics.CustomMetrics != null)
                _patchMetrics.CustomMetrics.Clear();

            EnsureCommonColumns(_serviceMetrics, true);
            EnsureCommonColumns(_patchMetrics, false);

            PopulateMetricList(_lstServiceMetrics, _serviceMetrics, MetricScope.Service);
            PopulateMetricList(_lstPatchMetrics, _patchMetrics, MetricScope.Patch);
            RebuildServiceGrid(false);
            RebuildPatchGrid(false);
            RefreshHistoryCombo(_store.Payload.LastSelectedKey);
            ApplyUiPreferences();
        }

        private void ApplyUiPreferences()
        {
            if (_store.Payload.Ui == null)
                _store.Payload.Ui = new UiPreferences();

            _suppressPanelEvents = true;
            try
            {
                chkShowService.Checked = _store.Payload.Ui.ShowServicePanel;
                chkShowPatch.Checked = _store.Payload.Ui.ShowPatchPanel;
            }
            finally
            {
                _suppressPanelEvents = false;
            }
            UpdatePanelVisibility();
        }

        private void chkShowPanel_CheckedChanged(object sender, EventArgs e)
        {
            if (_suppressPanelEvents) return;

            if (!chkShowService.Checked && !chkShowPatch.Checked)
            {
                _suppressPanelEvents = true;
                try
                {
                    // En az bir panel açık kalmalı
                    if (sender == chkShowService)
                        chkShowService.Checked = true;
                    else
                        chkShowPatch.Checked = true;
                }
                finally
                {
                    _suppressPanelEvents = false;
                }
                MessageBox.Show(this, "En az bir panel görünür olmalıdır.", "Bilgi",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            UpdatePanelVisibility();
            PersistUiPreferences();
        }

        private void UpdatePanelVisibility()
        {
            var showService = chkShowService.Checked;
            var showPatch = chkShowPatch.Checked;

            tabs.SuspendLayout();
            try
            {
                tabs.TabPages.Clear();
                if (showService)
                    tabs.TabPages.Add(tabService);
                if (showPatch)
                    tabs.TabPages.Add(tabPatch);
            }
            finally
            {
                tabs.ResumeLayout();
            }
        }

        private void PersistUiPreferences()
        {
            try
            {
                if (_store.Payload.Ui == null)
                    _store.Payload.Ui = new UiPreferences();
                _store.Payload.Ui.ShowServicePanel = chkShowService.Checked;
                _store.Payload.Ui.ShowPatchPanel = chkShowPatch.Checked;
                _store.Save();
            }
            catch { /* ignore */ }
        }

        private void mnuExport_Click(object sender, EventArgs e)
        {
            try
            {
                ApplyServiceMetricSelection();
                ApplyPatchMetricSelection();
                _store.Payload.ServiceMetrics = _serviceMetrics;
                _store.Payload.PatchMetrics = _patchMetrics;
                if (_store.Payload.Ui == null)
                    _store.Payload.Ui = new UiPreferences();
                _store.Payload.Ui.ShowServicePanel = chkShowService.Checked;
                _store.Payload.Ui.ShowPatchPanel = chkShowPatch.Checked;

                var password = PasswordInputDialog.Prompt(this,
                    "Dışa Aktar",
                    "Yapılandırmayı başka makinede açmak için bir şifre belirleyin.\nBu şifreyi import sırasında girmeniz gerekecek.",
                    requireConfirm: true);
                if (password == null) return;

                using (var sfd = new SaveFileDialog
                {
                    Filter = "RSM Taşınabilir Yapılandırma (*.rsmx)|*.rsmx|Tüm Dosyalar (*.*)|*.*",
                    DefaultExt = "rsmx",
                    FileName = "RemoteServiceManager_export_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".rsmx",
                    Title = "Yapılandırmayı Dışa Aktar"
                })
                {
                    if (sfd.ShowDialog(this) != DialogResult.OK) return;
                    _store.ExportPortable(sfd.FileName, password);
                    lblStatus.Text = "Dışa aktarıldı: " + sfd.FileName;
                    MessageBox.Show(this,
                        "Yapılandırma şifreli olarak dışa aktarıldı.\nHedef makinede Dosya → İçe Aktar ile aynı şifreyi kullanın.",
                        "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Dışa Aktarma Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void mnuImport_Click(object sender, EventArgs e)
        {
            try
            {
                using (var ofd = new OpenFileDialog
                {
                    Filter = "RSM Taşınabilir Yapılandırma (*.rsmx)|*.rsmx|Tüm Dosyalar (*.*)|*.*",
                    Title = "Yapılandırmayı İçe Aktar"
                })
                {
                    if (ofd.ShowDialog(this) != DialogResult.OK) return;

                    var password = PasswordInputDialog.Prompt(this,
                        "İçe Aktar",
                        "Export sırasında belirlenen şifreyi girin:",
                        requireConfirm: false);
                    if (password == null) return;

                    _store.ImportPortable(ofd.FileName, password);
                    ApplyLoadedPayload();
                    lblStatus.Text = "İçe aktarıldı. Bu makineye bağlandı: " + _store.ConfigPath;
                    MessageBox.Show(this,
                        "Yapılandırma başarıyla içe aktarıldı ve bu makineye bağlandı.",
                        "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "İçe Aktarma Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void EnsureCommonColumns(MetricPreferences prefs, bool service)
        {
            if (prefs.EnabledBuiltInIds == null)
                prefs.EnabledBuiltInIds = new List<string>();
            // Her iki ekranda da BSOD + bekleyen update bulunabilsin
            var must = service
                ? new[] { MetricCatalog.Bsod, MetricCatalog.PendingUpdates, MetricCatalog.IsOnline }
                : new[] { MetricCatalog.PatchBsod, MetricCatalog.UpdateSummary, MetricCatalog.PatchOnline };
            foreach (var id in must)
            {
                if (!prefs.EnabledBuiltInIds.Contains(id))
                    prefs.EnabledBuiltInIds.Add(id);
            }
        }

        private void PopulateMetricList(CheckedListBox list, MetricPreferences prefs, MetricScope scope)
        {
            _suppressMetricEvents = true;
            try
            {
                list.Items.Clear();
                var enabled = new HashSet<string>(prefs.EnabledBuiltInIds ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
                var builtins = scope == MetricScope.Patch
                    ? MetricCatalog.GetBuiltInPatchMetrics()
                    : MetricCatalog.GetBuiltInServiceMetrics();

                // Patch listesinde PendingUpdates alias'ını gösterme (UpdateSummary var)
                foreach (var m in builtins)
                {
                    if (scope == MetricScope.Patch && m.Id == MetricCatalog.PendingUpdates)
                        continue;
                    list.Items.Add(m, enabled.Contains(m.Id));
                }
                // Özel/dinamik metrik ekleme kapalı — listede gösterilmez.
            }
            finally
            {
                _suppressMetricEvents = false;
            }
        }

        private void ApplyServiceMetricSelection()
        {
            _serviceMetrics = ReadPrefsFromList(_lstServiceMetrics, _serviceMetrics, MetricScope.Service);
        }

        private void ApplyPatchMetricSelection()
        {
            _patchMetrics = ReadPrefsFromList(_lstPatchMetrics, _patchMetrics, MetricScope.Patch);
        }

        private MetricPreferences ReadPrefsFromList(CheckedListBox list, MetricPreferences existing, MetricScope scope)
        {
            var prefs = new MetricPreferences
            {
                EnabledBuiltInIds = new List<string>(),
                CustomMetrics = existing.CustomMetrics != null
                    ? existing.CustomMetrics.Select(CloneMetric).ToList()
                    : new List<MetricDefinition>()
            };

            for (int i = 0; i < list.Items.Count; i++)
            {
                var m = list.Items[i] as MetricDefinition;
                if (m == null) continue;
                bool on = list.GetItemChecked(i);
                if (m.IsCustom)
                {
                    var found = prefs.CustomMetrics.FirstOrDefault(x => x.Id == m.Id);
                    if (found != null) found.Enabled = on;
                }
                else if (on)
                {
                    prefs.EnabledBuiltInIds.Add(m.Id);
                }
            }
            return prefs;
        }

        private static MetricDefinition CloneMetric(MetricDefinition m)
        {
            return new MetricDefinition
            {
                Id = m.Id,
                DisplayName = m.DisplayName,
                ColumnHeader = m.ColumnHeader,
                Scope = m.Scope,
                Kind = m.Kind,
                Enabled = m.Enabled,
                Param1 = m.Param1,
                Param2 = m.Param2
            };
        }

        private void RebuildServiceGrid(bool keepRows)
        {
            ApplyServiceMetricSelection();
            var metrics = MetricCatalog.ResolveActive(_serviceMetrics, MetricScope.Service);
            var old = _serviceTable;
            _serviceTable = GridBuilder.CreateTable(metrics);
            if (keepRows && old != null)
            {
                foreach (var kv in _serviceResults)
                    GridBuilder.AddOrUpdateRow(_serviceTable, kv.Value, metrics);
            }
            else
            {
                _serviceResults.Clear();
            }
            _serviceBinding.DataSource = _serviceTable;
            GridBuilder.BindGrid(_dgvService, _serviceTable, metrics);
            _dgvService.DataSource = _serviceBinding;
        }

        private void RebuildPatchGrid(bool keepRows)
        {
            ApplyPatchMetricSelection();
            var metrics = MetricCatalog.ResolveActive(_patchMetrics, MetricScope.Patch);
            var old = _patchTable;
            _patchTable = GridBuilder.CreateTable(metrics);
            if (keepRows && old != null)
            {
                foreach (var kv in _patchResults)
                    GridBuilder.AddOrUpdateRow(_patchTable, kv.Value, metrics);
            }
            else
            {
                _patchResults.Clear();
            }
            _patchBinding.DataSource = _patchTable;
            GridBuilder.BindGrid(_dgvPatch, _patchTable, metrics);
            _dgvPatch.DataSource = _patchBinding;
        }

        private void PersistMetrics()
        {
            try
            {
                if (_serviceMetrics.CustomMetrics != null)
                    _serviceMetrics.CustomMetrics.Clear();
                if (_patchMetrics.CustomMetrics != null)
                    _patchMetrics.CustomMetrics.Clear();
                _store.Payload.ServiceMetrics = _serviceMetrics;
                _store.Payload.PatchMetrics = _patchMetrics;
                _store.Save();
            }
            catch { /* ignore */ }
        }

        private int GetMaxParallel()
        {
            var opts = _store.Payload?.DefaultOptions;
            if (opts != null && opts.MaxParallel > 0)
                return Math.Min(opts.MaxParallel, 10);
            return DefaultMaxParallel;
        }

        private List<string> ParseComputers()
        {
            return txtComputers.Text
                .Split(new[] { "\r\n", "\n", ",", ";" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => x.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<string> ParseServices()
        {
            return (_txtServices.Text ?? "")
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => x.Length > 0)
                .ToList();
        }

        private static void ApplyFilter(BindingSource src, string text)
        {
            var q = (text ?? "").Trim().Replace("'", "''");
            if (string.IsNullOrWhiteSpace(q))
                src.RemoveFilter();
            else
                src.Filter = "ComputerName LIKE '%" + q + "%' OR Hostname LIKE '%" + q + "%'";
        }

        private async Task RunServiceCheckAsync(List<string> computers)
        {
            if (computers == null || computers.Count == 0)
            {
                MessageBox.Show(this, "En az bir bilgisayar girin.", "Uyarı");
                return;
            }

            ApplyServiceMetricSelection();
            RebuildServiceGrid(false);
            var metrics = MetricCatalog.ResolveActive(_serviceMetrics, MetricScope.Service);
            var services = ParseServices();
            if (metrics.Any(m => m.Id == MetricCatalog.Running || m.Id == MetricCatalog.NotRunning) && services.Count == 0)
            {
                MessageBox.Show(this, "Servis kolonları açık; servis adı girin veya ilgili metrikleri kapatın.", "Uyarı");
                return;
            }

            GridBuilder.SeedPendingRows(_serviceTable, _serviceResults, computers, metrics);

            await RunParallelAsync(computers, async (cn, token) =>
            {
                HostResult result;
                try
                {
                    result = await Task.Run(() => _checker.Check(
                        cn, services, _txtFilePath.Text.Trim(), _txtRegPath.Text.Trim(), _txtRegValue.Text.Trim(),
                        _serviceMetrics), token);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    result = HostResult.CreateError(cn, ex.Message);
                }

                BeginInvoke(new Action(() =>
                {
                    _serviceResults[cn] = result;
                    GridBuilder.AddOrUpdateRow(_serviceTable, result, metrics);
                }));
            }, "servis kontrol");

            _btnServiceFix.Enabled = _serviceResults.Values.Any(r => r.IsOnline
                && r.ServiceDetail != null
                && r.ServiceDetail.Any(s => s.Status != "Running" && s.Status != "NotFound"));
        }

        private async Task RunPatchCheckAsync(List<string> computers, bool resetTable)
        {
            if (computers == null || computers.Count == 0)
            {
                MessageBox.Show(this, "En az bir bilgisayar girin.", "Uyarı");
                return;
            }

            ApplyPatchMetricSelection();
            var metrics = MetricCatalog.ResolveActive(_patchMetrics, MetricScope.Patch);
            if (resetTable || _patchTable == null)
            {
                RebuildPatchGrid(false);
                GridBuilder.SeedPendingRows(_patchTable, _patchResults, computers, metrics);
            }
            else
            {
                GridBuilder.SeedPendingRows(_patchTable, _patchResults, computers, metrics);
            }

            await RunParallelAsync(computers, async (cn, token) =>
            {
                HostResult result;
                try
                {
                    result = await Task.Run(() => _patchChecker.Check(cn, _patchMetrics), token);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    result = HostResult.CreateError(cn, ex.Message, patchMode: true);
                }

                BeginInvoke(new Action(() =>
                {
                    _patchResults[cn] = result;
                    GridBuilder.AddOrUpdateRow(_patchTable, result, metrics);
                }));
            }, "patch kontrol");
        }

        private async Task RunParallelAsync(List<string> computers, Func<string, CancellationToken, Task> work, string label)
        {
            _cts = new CancellationTokenSource();
            SetBusy(true);
            progressBar.Maximum = Math.Max(1, computers.Count);
            progressBar.Value = 0;
            var done = 0;
            var maxParallel = GetMaxParallel();
            var gate = new SemaphoreSlim(maxParallel);

            try
            {
                var tasks = computers.Select(async cn =>
                {
                    await gate.WaitAsync(_cts.Token).ConfigureAwait(false);
                    try
                    {
                        _cts.Token.ThrowIfCancellationRequested();
                        await work(cn, _cts.Token).ConfigureAwait(false);
                        BeginInvoke(new Action(() =>
                        {
                            done++;
                            progressBar.Value = Math.Min(done, progressBar.Maximum);
                            lblStatus.Text = done + " / " + computers.Count + " (" + label + ", max " + maxParallel + ")";
                        }));
                    }
                    finally { gate.Release(); }
                });
                await Task.WhenAll(tasks);
                lblStatus.Text = "Tamamlandı: " + label + " (" + computers.Count + ")";
            }
            catch (OperationCanceledException)
            {
                lblStatus.Text = "İptal edildi.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                gate.Dispose();
                SetBusy(false);
                _cts.Dispose();
                _cts = null;
            }
        }

        private async Task RunServiceFixAsync(string onlyComputer)
        {
            var targets = _serviceResults.Values.Where(r =>
                r.IsOnline
                && (onlyComputer == null || string.Equals(r.ComputerName, onlyComputer, StringComparison.OrdinalIgnoreCase))
                && r.ServiceDetail != null
                && r.ServiceDetail.Any(s => s.Status != "Running" && s.Status != "NotFound")).ToList();

            if (targets.Count == 0)
            {
                MessageBox.Show(this, "Düzeltilecek servis yok.", "Bilgi");
                return;
            }

            if (MessageBox.Show(this, targets.Count + " makinede düzeltme yapılsın mı?", "Onay",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            var metrics = MetricCatalog.ResolveActive(_serviceMetrics, MetricScope.Service);
            await RunParallelAsync(targets.Select(t => t.ComputerName).ToList(), async (cn, token) =>
            {
                HostResult host;
                if (!_serviceResults.TryGetValue(cn, out host)) return;
                var fix = await Task.Run(() => _fixer.Fix(cn, host.ServiceDetail), token);
                host.ServiceDetail = fix.Select(x => new ServiceStatusItem { Name = x.ServiceName, Status = x.SonrakiDurum }).ToList();
                host.Running = string.Join(", ", host.ServiceDetail.Where(s => s.Status == "Running").Select(s => s.Name));
                host.NotRunning = string.Join(", ", host.ServiceDetail.Where(s => s.Status != "Running")
                    .Select(s => s.Name + "(" + s.Status + ")"));
                BeginInvoke(new Action(() => GridBuilder.AddOrUpdateRow(_serviceTable, host, metrics)));
            }, "servis düzeltme");
        }

        private async Task RunPatchServiceFixAsync()
        {
            var selected = GetSelectedComputers(_dgvPatch);
            if (selected.Count == 0)
                selected = _patchResults.Keys.ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, "Makine yok.", "Bilgi");
                return;
            }

            if (MessageBox.Show(this,
                    selected.Count + " makinede patch servisleri (ccmexec, wuauserv, bits, cryptsvc) başlatılsın mı?",
                    "Onay", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            var metrics = MetricCatalog.ResolveActive(_patchMetrics, MetricScope.Patch);
            await RunParallelAsync(selected, async (cn, token) =>
            {
                var detail = PatchChecker.PatchServices
                    .Select(n => new ServiceStatusItem { Name = n, Status = "Stopped" }).ToList();
                await Task.Run(() => _fixer.Fix(cn, detail), token);
                var refreshed = await Task.Run(() => _patchChecker.Check(cn, _patchMetrics), token);
                BeginInvoke(new Action(() =>
                {
                    _patchResults[cn] = refreshed;
                    GridBuilder.AddOrUpdateRow(_patchTable, refreshed, metrics);
                }));
            }, "patch servis düzeltme");
        }

        private void OpenMachineDetail(string computer)
        {
            if (string.IsNullOrWhiteSpace(computer)) return;
            HostResult svc, patch;
            _serviceResults.TryGetValue(computer, out svc);
            _patchResults.TryGetValue(computer, out patch);

            using (var dlg = new MachineDetailDialog(
                computer, svc, patch,
                () => ParseServices(),
                () => _txtFilePath.Text.Trim(),
                () => _txtRegPath.Text.Trim(),
                () => _txtRegValue.Text.Trim(),
                () => { ApplyServiceMetricSelection(); return _serviceMetrics; },
                () => { ApplyPatchMetricSelection(); return _patchMetrics; }))
            {
                dlg.ResultUpdated += (result, isPatch) =>
                {
                    if (result == null) return;
                    if (isPatch)
                    {
                        _patchResults[result.ComputerName] = result;
                        var metrics = MetricCatalog.ResolveActive(_patchMetrics, MetricScope.Patch);
                        if (_patchTable == null) RebuildPatchGrid(false);
                        // Bilgisayarı sol listeye ekle
                        EnsureComputerInList(result.ComputerName);
                        GridBuilder.AddOrUpdateRow(_patchTable, result, metrics);
                    }
                    else
                    {
                        _serviceResults[result.ComputerName] = result;
                        var metrics = MetricCatalog.ResolveActive(_serviceMetrics, MetricScope.Service);
                        if (_serviceTable == null) RebuildServiceGrid(false);
                        EnsureComputerInList(result.ComputerName);
                        GridBuilder.AddOrUpdateRow(_serviceTable, result, metrics);
                    }
                };
                dlg.ShowDialog(this);
            }
        }

        private void EnsureComputerInList(string name)
        {
            var list = ParseComputers();
            if (list.Any(x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase))) return;
            if (!string.IsNullOrWhiteSpace(txtComputers.Text) && !txtComputers.Text.EndsWith("\n"))
                txtComputers.AppendText(Environment.NewLine);
            txtComputers.AppendText(name + Environment.NewLine);
        }

        private string GetComputerFromGrid(DataGridView dgv, int rowIndex)
        {
            var drv = dgv.Rows[rowIndex].DataBoundItem as DataRowView;
            if (drv == null) return null;
            if (drv.Row.Table.Columns.Contains("ComputerName"))
                return Convert.ToString(drv["ComputerName"]);
            return Convert.ToString(drv["Hostname"]);
        }

        private List<string> GetSelectedComputers(DataGridView dgv)
        {
            var list = new List<string>();
            foreach (DataGridViewRow row in dgv.SelectedRows)
            {
                var name = GetComputerFromGrid(dgv, row.Index);
                if (!string.IsNullOrWhiteSpace(name)) list.Add(name);
            }
            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private void ExportGrid(BindingSource binding, DataTable table, string prefix)
        {
            if (table == null || table.Rows.Count == 0)
            {
                MessageBox.Show(this, "Aktarılacak veri yok.", "Bilgi");
                return;
            }
            var export = table.Clone();
            foreach (var item in binding.List)
            {
                var drv = item as DataRowView;
                if (drv != null) export.ImportRow(drv.Row);
            }
            ExcelExporter.ExportDataTableToExcelAndOpen(export, prefix + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
        }

        private PresetRecord BuildProfileFromUi(string name, string timestamp)
        {
            ApplyServiceMetricSelection();
            ApplyPatchMetricSelection();
            if (_serviceMetrics.CustomMetrics != null)
                _serviceMetrics.CustomMetrics.Clear();
            if (_patchMetrics.CustomMetrics != null)
                _patchMetrics.CustomMetrics.Clear();

            return new PresetRecord
            {
                Name = name,
                Timestamp = timestamp,
                Services = ParseServices(),
                FilePath = _txtFilePath.Text.Trim(),
                RegPath = _txtRegPath.Text.Trim(),
                RegValueName = _txtRegValue.Text.Trim(),
                ServiceMetrics = _serviceMetrics,
                PatchMetrics = _patchMetrics
            };
        }

        private PresetRecord GetSelectedPreset()
        {
            if (_cboHistory.SelectedIndex < 0 || _cboHistory.SelectedIndex >= _cboHistory.Items.Count)
                return null;
            var item = _cboHistory.Items[_cboHistory.SelectedIndex] as HistoryItem;
            return item?.Data;
        }

        private void SaveProfile()
        {
            var name = NameInputDialog.Prompt(this);
            if (string.IsNullOrWhiteSpace(name)) return;

            if (_store.ProfileNameExists(name))
            {
                MessageBox.Show(this, "Bu profil ismi zaten mevcut.", "Uyarı",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var record = BuildProfileFromUi(name, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                _store.Payload.ServiceMetrics = _serviceMetrics;
                _store.Payload.PatchMetrics = _patchMetrics;
                var key = _store.SavePreset(record);
                RefreshHistoryCombo(key);
                lblStatus.Text = "Kaydedildi: " + name;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateProfile()
        {
            var selected = GetSelectedPreset();
            if (selected == null)
            {
                MessageBox.Show(this, "Güncellenecek bir profil seçin.", "Bilgi",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var displayName = string.IsNullOrWhiteSpace(selected.Name) ? selected.Timestamp : selected.Name;
            if (MessageBox.Show(this,
                    "\"" + displayName + "\" profili mevcut ayarlarla güncellensin mi?",
                    "Profil Güncelle",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            try
            {
                var name = string.IsNullOrWhiteSpace(selected.Name) ? selected.Timestamp : selected.Name;
                var record = BuildProfileFromUi(name, selected.Timestamp);
                _store.Payload.ServiceMetrics = _serviceMetrics;
                _store.Payload.PatchMetrics = _patchMetrics;
                var key = _store.UpdatePreset(selected.Timestamp, record);
                RefreshHistoryCombo(key);
                lblStatus.Text = "Güncellendi: " + name;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteProfile()
        {
            var selected = GetSelectedPreset();
            if (selected == null)
            {
                MessageBox.Show(this, "Silinecek bir profil seçin.", "Bilgi",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var displayName = string.IsNullOrWhiteSpace(selected.Name) ? selected.Timestamp : selected.Name;
            if (MessageBox.Show(this,
                    "\"" + displayName + "\" profili silinsin mi?\nBu işlem geri alınamaz.",
                    "Profil Sil",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            try
            {
                _store.DeletePreset(selected.Timestamp);
                RefreshHistoryCombo(_store.Payload.LastSelectedKey);
                if (_cboHistory.Items.Count == 0)
                    lblStatus.Text = "Profil silindi. Kayıtlı profil kalmadı.";
                else
                    lblStatus.Text = "Profil silindi: " + displayName;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshHistoryCombo(string selectKey)
        {
            _cboHistory.Items.Clear();
            var presets = (_store.Payload.Presets ?? new List<PresetRecord>())
                .OrderByDescending(p => p.Timestamp).ToList();
            foreach (var p in presets)
            {
                var display = string.IsNullOrWhiteSpace(p.Name) ? p.Timestamp : p.Name + "  (" + p.Timestamp + ")";
                _cboHistory.Items.Add(new HistoryItem { Display = display, Data = p });
            }
            _cboHistory.DisplayMember = "Display";

            if (string.IsNullOrEmpty(selectKey))
                selectKey = _store.Payload.LastSelectedKey;
            var idx = 0;
            for (int i = 0; i < presets.Count; i++)
            {
                if (presets[i].Timestamp == selectKey) { idx = i; break; }
            }
            if (_cboHistory.Items.Count > 0)
            {
                _cboHistory.SelectedIndex = Math.Min(idx, _cboHistory.Items.Count - 1);
                LoadHistoryItem(_cboHistory.SelectedIndex);
            }
        }

        private void LoadHistoryItem(int index)
        {
            if (index < 0 || index >= _cboHistory.Items.Count) return;
            var item = _cboHistory.Items[index] as HistoryItem;
            if (item?.Data == null) return;
            var data = item.Data;
            _txtServices.Text = data.Services != null ? string.Join(", ", data.Services) : "";
            _txtFilePath.Text = data.FilePath ?? "";
            _txtRegPath.Text = data.RegPath ?? "";
            _txtRegValue.Text = data.RegValueName ?? "";

            if (data.ServiceMetrics != null) _serviceMetrics = data.ServiceMetrics;
            else if (data.Options != null) _serviceMetrics = MetricCatalog.FromLegacy(data.Options);
            if (data.PatchMetrics != null) _patchMetrics = data.PatchMetrics;

            if (_serviceMetrics.CustomMetrics != null)
                _serviceMetrics.CustomMetrics.Clear();
            if (_patchMetrics.CustomMetrics != null)
                _patchMetrics.CustomMetrics.Clear();

            PopulateMetricList(_lstServiceMetrics, _serviceMetrics, MetricScope.Service);
            PopulateMetricList(_lstPatchMetrics, _patchMetrics, MetricScope.Patch);
            RebuildServiceGrid(true);
            RebuildPatchGrid(true);

            try { _store.SetLastSelected(data.Timestamp); } catch { }
            lblStatus.Text = "Profil: " + (string.IsNullOrWhiteSpace(data.Name) ? data.Timestamp : data.Name);
        }

        private void SetBusy(bool busy)
        {
            _btnServiceCheck.Enabled = !busy;
            _btnServiceFix.Enabled = !busy && _serviceResults.Count > 0;
            _btnServiceSave.Enabled = !busy;
            _btnServiceUpdate.Enabled = !busy;
            _btnServiceDelete.Enabled = !busy;
            _btnServiceExport.Enabled = !busy;
            _btnServiceCancel.Enabled = busy;
            _btnPatchRefresh.Enabled = !busy;
            _btnPatchRefreshSelected.Enabled = !busy;
            _btnPatchFix.Enabled = !busy;
            _btnPatchExport.Enabled = !busy;
            _btnPatchCancel.Enabled = busy;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try { _cts?.Cancel(); } catch { }
            try
            {
                ApplyServiceMetricSelection();
                ApplyPatchMetricSelection();
                if (_serviceMetrics.CustomMetrics != null)
                    _serviceMetrics.CustomMetrics.Clear();
                if (_patchMetrics.CustomMetrics != null)
                    _patchMetrics.CustomMetrics.Clear();
                _store.Payload.ServiceMetrics = _serviceMetrics;
                _store.Payload.PatchMetrics = _patchMetrics;
                if (_store.Payload.Ui == null)
                    _store.Payload.Ui = new UiPreferences();
                _store.Payload.Ui.ShowServicePanel = chkShowService.Checked;
                _store.Payload.Ui.ShowPatchPanel = chkShowPatch.Checked;
                _store.Save();
            }
            catch { }
        }

        private class HistoryItem
        {
            public string Display { get; set; }
            public PresetRecord Data { get; set; }
            public override string ToString() { return Display; }
        }
    }
}
