using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Forms;
using RemoteServiceManager.Models;

namespace RemoteServiceManager.Services
{
    public static class GridBuilder
    {
        // Uzak makinelerden gelen veri güvenilmez kabul edilir.
        // Aşırı uzun değerler UI'ı kilitleyebileceği için hücre başına üst sınır uygulanır.
        private const int MaxCellLength = 8192;

        // Satır renkleri (her hücre biçimlendirmesinde yeniden oluşturulmasın diye statik)
        private static readonly Color PendingColor = Color.FromArgb(230, 240, 255);
        private static readonly Color OfflineColor = Color.LightCoral;
        private static readonly Color ErrorColor = Color.FromArgb(255, 220, 180);
        private static readonly Color BsodColor = Color.FromArgb(255, 200, 200);
        private static readonly Color WarnColor = Color.LightYellow;
        private static readonly Color OkColor = Color.LightGreen;

        // AddOrUpdateRow için O(1) satır arama indeksi (tablo başına, tablo ile birlikte GC'lenir)
        private sealed class RowIndex
        {
            public readonly Dictionary<string, DataRow> Map =
                new Dictionary<string, DataRow>(StringComparer.OrdinalIgnoreCase);
            public int SyncedRowCount = -1;
            public DataColumn KeyColumn;
        }

        private static readonly ConditionalWeakTable<DataTable, RowIndex> Indexes =
            new ConditionalWeakTable<DataTable, RowIndex>();

        public static DataTable CreateTable(IList<MetricDefinition> metrics, bool includeComputerName = true)
        {
            var table = new DataTable();
            if (includeComputerName)
                table.Columns.Add("ComputerName", typeof(string));
            table.Columns.Add("Hostname", typeof(string));

            var added = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ComputerName", "Hostname" };
            if (metrics != null)
            {
                foreach (var m in metrics)
                {
                    if (m == null) continue;
                    var key = m.ResultKey;
                    if (string.IsNullOrWhiteSpace(key)) continue;
                    if (added.Contains(key)) continue;
                    table.Columns.Add(key, typeof(string));
                    added.Add(key);
                }
            }
            return table;
        }

        public static void BindGrid(DataGridView dgv, DataTable table, IList<MetricDefinition> metrics)
        {
            if (dgv == null) return;

            EnableDoubleBuffering(dgv);

            dgv.SuspendLayout();
            try
            {
                dgv.AutoGenerateColumns = false;
                dgv.DataSource = null;
                dgv.Columns.Clear();

                var columns = new List<DataGridViewColumn>
                {
                    CreateScrollableColumn("Hostname", "Hostname", "Hostname")
                };

                if (metrics != null && table != null)
                {
                    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Hostname" };
                    foreach (var m in metrics)
                    {
                        if (m == null) continue;
                        var key = m.ResultKey;
                        if (string.IsNullOrWhiteSpace(key)) continue;
                        if (seen.Contains(key)) continue;
                        if (!table.Columns.Contains(key)) continue;
                        seen.Add(key);
                        var header = string.IsNullOrWhiteSpace(m.ColumnHeader) ? m.DisplayName : m.ColumnHeader;
                        columns.Add(CreateScrollableColumn(key, key, header));
                    }
                }

                dgv.Columns.AddRange(columns.ToArray());

                // Gizli ComputerName (satır eşleme için DataTable'da kalır; grid'de yok)
                dgv.DataSource = table;
                ApplyInitialColumnWidths(dgv);
            }
            finally
            {
                dgv.ResumeLayout();
            }
        }

        /// <summary>
        /// Çift tamponlama: renkli satırlarda yeniden çizim titremesini ve CPU kullanımını azaltır.
        /// RDP/Terminal Services oturumlarında kapalı bırakılır.
        /// </summary>
        private static void EnableDoubleBuffering(DataGridView dgv)
        {
            try
            {
                if (SystemInformation.TerminalServerSession) return;
                var prop = typeof(DataGridView).GetProperty("DoubleBuffered",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                if (prop != null) prop.SetValue(dgv, true, null);
            }
            catch
            {
                // Performans iyileştirmesi; başarısız olması işlevi etkilemez.
            }
        }

        private static DataGridViewTextBoxColumn CreateScrollableColumn(string name, string dataProperty, string header)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                DataPropertyName = dataProperty,
                HeaderText = header,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                MinimumWidth = 70,
                Width = 130,
                SortMode = DataGridViewColumnSortMode.Automatic,
                Resizable = DataGridViewTriState.True
            };
        }

        /// <summary>
        /// İçeriğe göre başlangıç genişliği verir; kullanıcı sütun kenarından genişletebilir.
        /// Toplam genişlik grid'i aşınca yatay kaydırma çubuğu görünür.
        /// </summary>
        private static void ApplyInitialColumnWidths(DataGridView dgv, int defaultWidth = 130, int minWidth = 70, int maxWidth = 360)
        {
            if (dgv == null || dgv.Columns.Count == 0) return;

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                col.MinimumWidth = minWidth;
                col.Width = defaultWidth;
                col.Resizable = DataGridViewTriState.True;
            }

            if (dgv.Rows.Count == 0) return;

            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader);
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.Width > maxWidth) col.Width = maxWidth;
                if (col.Width < minWidth) col.Width = minWidth;
            }
        }

        public static void SeedPendingRows(DataTable table, IDictionary<string, HostResult> cache,
            IEnumerable<string> computerNames, IList<MetricDefinition> metrics)
        {
            if (table == null || computerNames == null) return;

            // Toplu ekleme: bağlı grid'e satır başına bildirim gitmesini engeller (büyük listelerde belirgin hız kazancı).
            table.BeginLoadData();
            try
            {
                foreach (var cn in computerNames)
                {
                    if (string.IsNullOrWhiteSpace(cn)) continue;
                    var pending = HostResult.CreatePending(cn);
                    if (cache != null)
                        cache[cn] = pending;
                    AddOrUpdateRow(table, pending, metrics);
                }
            }
            finally
            {
                table.EndLoadData();
            }
        }

        public static void AddOrUpdateRow(DataTable table, HostResult result, IList<MetricDefinition> metrics)
        {
            if (table == null || result == null) return;

            var nameCol = table.Columns["ComputerName"];
            var hostCol = table.Columns["Hostname"];

            // Uzak makineden gelen değerler temizlenir (kontrol karakterleri, aşırı uzunluk)
            var computerName = Sanitize(result.ComputerName);
            var hostname = Sanitize(result.Hostname ?? result.ComputerName);

            // Eşleme anahtarı: ComputerName; yoksa Hostname (eskiden bu durumda istisna oluşuyordu)
            DataColumn keyCol = nameCol ?? hostCol;
            string key = nameCol != null ? computerName : hostname;

            DataRow existing = keyCol == null ? null : FindRow(table, keyCol, key);

            var row = existing ?? table.NewRow();

            // Mevcut satırda tek seferde güncelle: her sütun için ayrı değişiklik bildirimi/yeniden çizim olmasın
            bool editing = existing != null;
            if (editing) row.BeginEdit();
            try
            {
                if (nameCol != null)
                    SetIfChanged(row, nameCol, computerName);
                if (hostCol != null)
                    SetIfChanged(row, hostCol, hostname);

                if (metrics != null)
                {
                    foreach (var m in metrics)
                    {
                        if (m == null) continue;
                        var mkey = m.ResultKey;
                        if (string.IsNullOrWhiteSpace(mkey)) continue;
                        var col = table.Columns[mkey];
                        if (col == null) continue;
                        SetIfChanged(row, col, Sanitize(Convert.ToString(result.GetValue(mkey))));
                    }
                }

                var onlineCol = table.Columns[MetricCatalog.IsOnline];
                if (onlineCol != null)
                    SetIfChanged(row, onlineCol, Sanitize(Convert.ToString(result.GetValue(MetricCatalog.IsOnline))));

                if (editing) row.EndEdit();
            }
            catch
            {
                if (editing) row.CancelEdit();
                throw;
            }

            if (existing == null)
            {
                table.Rows.Add(row);
                if (keyCol != null)
                    RegisterRow(table, keyCol, key, row);
            }
        }

        public static void ApplyRowColors(DataGridViewCellFormattingEventArgs e, DataRowView rowView, bool patchMode)
        {
            if (e == null || rowView == null) return;

            // Bu metot her hücre için çağrılır; sütun aramaları tek seferde (DataColumn) yapılır.
            var row = rowView.Row;
            var cols = row.Table.Columns;
            Color back;

            if (patchMode)
            {
                var onlineCol = cols["PatchOnline"] ?? cols["IsOnline"];
                var online = onlineCol != null ? Convert.ToString(row[onlineCol]) : "";

                if (string.Equals(online, HostResult.PendingText, StringComparison.OrdinalIgnoreCase))
                {
                    back = PendingColor;
                }
                else if (string.Equals(online, "False", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(online, "Offline", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(online, "DNS'te yok", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(online, "Hayır", StringComparison.OrdinalIgnoreCase))
                {
                    back = OfflineColor;
                }
                else
                {
                    var summaryCol = cols["UpdateSummary"];
                    var bsodCol = cols["PatchBsod"];

                    if (summaryCol != null
                        && Convert.ToString(row[summaryCol]).IndexOf("Bekleyen update:", StringComparison.OrdinalIgnoreCase) >= 0)
                        back = WarnColor;
                    else if (bsodCol != null
                        && string.Equals(Convert.ToString(row[bsodCol]), "VAR", StringComparison.OrdinalIgnoreCase))
                        back = BsodColor;
                    else
                        back = OkColor;
                }
            }
            else
            {
                var onlineCol = cols["IsOnline"];
                var notRunningCol = cols["NotRunning"];
                var errCol = cols["ErrorMessage"];
                var bsodCol = cols["Bsod"];

                var isOnline = onlineCol != null ? Convert.ToString(row[onlineCol]) : "True";
                var notRunning = notRunningCol != null ? Convert.ToString(row[notRunningCol]) : "";

                if (string.Equals(isOnline, HostResult.PendingText, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(notRunning, HostResult.PendingText, StringComparison.OrdinalIgnoreCase))
                {
                    back = PendingColor;
                }
                else if (string.Equals(isOnline, "False", StringComparison.OrdinalIgnoreCase))
                {
                    back = OfflineColor;
                }
                else
                {
                    var err = errCol != null ? Convert.ToString(row[errCol]) : "";
                    if (!string.IsNullOrWhiteSpace(err))
                    {
                        back = ErrorColor;
                    }
                    else
                    {
                        var bsod = bsodCol != null ? Convert.ToString(row[bsodCol]) : "";
                        if (string.Equals(bsod, "VAR", StringComparison.OrdinalIgnoreCase))
                            back = BsodColor;
                        else if (!string.IsNullOrWhiteSpace(notRunning))
                            back = WarnColor;
                        else
                            back = OkColor;
                    }
                }
            }

            e.CellStyle.BackColor = back;
        }

        // ------------------------------------------------------------------
        // Yardımcılar
        // ------------------------------------------------------------------

        /// <summary>
        /// Uzak makineden gelen metni temizler: null -> boş, kontrol karakterleri (\r \n \t hariç) silinir,
        /// MaxCellLength üstü kısaltılır.
        /// </summary>
        private static string Sanitize(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;

            bool dirty = value.Length > MaxCellLength;
            if (!dirty)
            {
                for (int i = 0; i < value.Length; i++)
                {
                    char c = value[i];
                    if (char.IsControl(c) && c != '\r' && c != '\n' && c != '\t') { dirty = true; break; }
                }
            }
            if (!dirty) return value; // hızlı yol: ek ayırma yok

            int len = Math.Min(value.Length, MaxCellLength);
            var sb = new StringBuilder(len + 3);
            for (int i = 0; i < len; i++)
            {
                char c = value[i];
                if (char.IsControl(c) && c != '\r' && c != '\n' && c != '\t') continue;
                sb.Append(c);
            }

            if (value.Length > MaxCellLength)
            {
                if (sb.Length > 0 && char.IsHighSurrogate(sb[sb.Length - 1])) sb.Length--;
                sb.Append("...");
            }
            return sb.ToString();
        }

        private static void SetIfChanged(DataRow row, DataColumn col, string value)
        {
            var current = row[col] as string;
            if (current != null && string.Equals(current, value, StringComparison.Ordinal)) return;
            row[col] = value;
        }

        private static bool IsLive(DataTable table, DataRow r, DataColumn keyCol, string key)
        {
            return r.RowState != DataRowState.Detached
                && r.RowState != DataRowState.Deleted
                && ReferenceEquals(r.Table, table)
                && string.Equals(Convert.ToString(r[keyCol]), key, StringComparison.OrdinalIgnoreCase);
        }

        private static void RebuildIndex(RowIndex idx, DataTable table, DataColumn keyCol)
        {
            idx.Map.Clear();
            foreach (DataRow r in table.Rows)
            {
                if (r.RowState == DataRowState.Deleted || r.RowState == DataRowState.Detached) continue;
                var k = Convert.ToString(r[keyCol]);
                if (!idx.Map.ContainsKey(k)) // eski davranış: ilk eşleşen satır
                    idx.Map[k] = r;
            }
            idx.KeyColumn = keyCol;
            idx.SyncedRowCount = table.Rows.Count;
        }

        /// <summary>
        /// Eski O(n) döngü yerine sözlük araması. Tablo dışarıdan değiştirilmişse (satır sayısı farkı,
        /// silinmiş/ayrılmış satır) indeks otomatik yeniden kurulur.
        /// </summary>
        private static DataRow FindRow(DataTable table, DataColumn keyCol, string key)
        {
            var idx = Indexes.GetValue(table, t => new RowIndex());
            lock (idx)
            {
                DataRow r;
                bool found = idx.Map.TryGetValue(key, out r);
                if (found && IsLive(table, r, keyCol, key)) return r;

                bool stale = found; // bulundu ama artık geçerli değil
                if (stale || idx.SyncedRowCount != table.Rows.Count || !ReferenceEquals(idx.KeyColumn, keyCol))
                {
                    RebuildIndex(idx, table, keyCol);
                    if (idx.Map.TryGetValue(key, out r) && IsLive(table, r, keyCol, key)) return r;
                }
                return null;
            }
        }

        private static void RegisterRow(DataTable table, DataColumn keyCol, string key, DataRow row)
        {
            var idx = Indexes.GetValue(table, t => new RowIndex());
            lock (idx)
            {
                idx.Map[key] = row;
                idx.KeyColumn = keyCol;
                idx.SyncedRowCount = table.Rows.Count;
            }
        }
    }
}