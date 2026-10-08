using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace RemoteServiceManager.Services
{
    public static class ExcelExporter
    {
        // Excel sınırları
        private const int MaxExcelRows = 1048576;
        private const int MaxExcelColumns = 16384;
        private const int MaxExcelCellChars = 32767;

        // COM'a tek seferde yazılacak satır sayısı (çok büyük dizileri tek çağrıda marshal etmemek için)
        private const int ComChunkRows = 5000;

        private static readonly char[] CsvSpecialChars = { ';', '"', '\n', '\r' };

        public static void ExportDataTableToExcelAndOpen(DataTable table, string defaultFileName)
        {
            if (table == null || table.Rows.Count == 0)
            {
                MessageBox.Show("Aktarılacak veri yok.", "Bilgi");
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Dosyası (*.xlsx)|*.xlsx|CSV Dosyası (*.csv)|*.csv";
                sfd.FileName = defaultFileName + ".xlsx";
                if (sfd.ShowDialog() != DialogResult.OK)
                    return;

                if (sfd.FilterIndex == 2 || sfd.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    ExportCsv(table, sfd.FileName);
                    MessageBox.Show("Dışa aktarıldı: " + sfd.FileName, "Bilgi");
                    return;
                }

                try
                {
                    ExportViaExcelCom(table, sfd.FileName);
                }
                catch (Exception ex)
                {
                    // Güvenlik: yedek CSV, kullanıcının onaylamadığı mevcut bir dosyanın üzerine sessizce yazmasın.
                    var csvFallback = GetNonCollidingPath(Path.ChangeExtension(sfd.FileName, ".csv"));
                    ExportCsv(table, csvFallback);
                    MessageBox.Show(
                        "Excel bulunamadı veya aktarım hatası. CSV olarak kaydedildi:\n"
                        + csvFallback + "\n\n" + ex.Message, "Uyarı");
                }
            }
        }

        public static void ExportSummariesToExcel(IList<Models.HostFixSummary> data, string defaultFileName)
        {
            var dt = new DataTable();
            dt.Columns.Add("Hostname");
            dt.Columns.Add("Status");

            dt.BeginLoadData();
            try
            {
                foreach (var d in data ?? Enumerable.Empty<Models.HostFixSummary>())
                {
                    if (d == null) continue;
                    dt.Rows.Add(d.Hostname, d.Status);
                }
            }
            finally
            {
                dt.EndLoadData();
            }

            ExportDataTableToExcelAndOpen(dt, defaultFileName);
        }

        private static void ExportCsv(DataTable table, string path)
        {
            var columns = table.Columns;
            int colCount = columns.Count;

            // Performans: tüm içeriği önce StringBuilder'da biriktirmek yerine doğrudan dosyaya akıt.
            // UTF8Encoding(true) = Encoding.UTF8 ile aynı çıktı (BOM'lu; Excel'de Türkçe karakterler için gerekli).
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65536))
            using (var sw = new StreamWriter(fs, new UTF8Encoding(true), 65536))
            {
                var line = new StringBuilder(256);

                for (int c = 0; c < colCount; c++)
                {
                    if (c > 0) line.Append(';');
                    line.Append(EscapeCsv(NeutralizeFormula(columns[c].ColumnName)));
                }
                sw.WriteLine(line.ToString());

                foreach (DataRow row in table.Rows)
                {
                    line.Clear();
                    for (int c = 0; c < colCount; c++)
                    {
                        if (c > 0) line.Append(';');
                        line.Append(EscapeCsv(NeutralizeFormula(Convert.ToString(row[c]))));
                    }
                    sw.WriteLine(line.ToString());
                }
            }
        }

        private static string EscapeCsv(string value)
        {
            value = value ?? string.Empty;
            if (value.IndexOfAny(CsvSpecialChars) >= 0)
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
        }

        /// <summary>
        /// Güvenlik (CSV/formül enjeksiyonu): uzak makinelerden gelen değerler (servis adı, registry değeri,
        /// kullanıcı adı, süreç adı vb.) '=', '+', '-', '@', TAB veya CR ile başlıyorsa Excel bunu formül
        /// olarak çalıştırabilir. Başına tek tırnak eklenerek metin olarak saklanması sağlanır.
        /// Salt sayısal değerler (örn. -5, +90) formül değildir; olduğu gibi bırakılır.
        /// </summary>
        private static string NeutralizeFormula(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value ?? string.Empty;

            char first = value[0];
            if (first != '=' && first != '+' && first != '-' && first != '@' && first != '\t' && first != '\r')
                return value;

            if (first == '+' || first == '-')
            {
                double d;
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out d)
                    || double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out d))
                    return value;
            }

            return "'" + value;
        }

        // Excel hücresine yazılacak güvenli metin: formül etkisiz + hücre karakter sınırı.
        private static string CellText(string value)
        {
            var s = NeutralizeFormula(value);
            return s.Length > MaxExcelCellChars ? s.Substring(0, MaxExcelCellChars) : s;
        }

        private static string GetNonCollidingPath(string path)
        {
            if (!File.Exists(path))
                return path;

            var dir = Path.GetDirectoryName(path) ?? string.Empty;
            var name = Path.GetFileNameWithoutExtension(path);
            var ext = Path.GetExtension(path);

            for (int i = 1; i < 1000; i++)
            {
                var candidate = Path.Combine(dir, name + " (" + i + ")" + ext);
                if (!File.Exists(candidate))
                    return candidate;
            }
            return path;
        }

        private static void ExportViaExcelCom(DataTable table, string path)
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("Microsoft Excel yüklü değil.");

            int colCount = table.Columns.Count;
            int rowCount = table.Rows.Count;

            // Excel'i hiç başlatmadan sınır kontrolü: aşılırsa çağıran taraf CSV'ye düşer.
            if (rowCount + 1 > MaxExcelRows || colCount > MaxExcelColumns)
                throw new InvalidOperationException("Veri Excel satır/sütun sınırlarını aşıyor.");

            object excel = Activator.CreateInstance(excelType);
            object workbooks = null, workbook = null, sheets = null, sheet = null;
            try
            {
                Set(excel, "Visible", false);
                Set(excel, "DisplayAlerts", false); // görünmezken bir iletişim kutusunda takılı kalmasın

                workbooks = Get(excel, "Workbooks");
                workbook = Call(workbooks, "Add");
                sheets = Get(workbook, "Worksheets");
                sheet = Get(sheets, "Item", 1);

                if (colCount > 0)
                {
                    // Performans: hücre hücre (satır*sütun*2 reflection çağrısı) yerine blok halinde Value2 ataması.
                    var header = new object[1, colCount];
                    for (int c = 0; c < colCount; c++)
                        header[0, c] = CellText(table.Columns[c].ColumnName);
                    WriteBlock(sheet, 1, header);

                    for (int start = 0; start < rowCount; start += ComChunkRows)
                    {
                        int take = Math.Min(ComChunkRows, rowCount - start);
                        var block = new object[take, colCount];
                        for (int r = 0; r < take; r++)
                        {
                            var dataRow = table.Rows[start + r];
                            for (int c = 0; c < colCount; c++)
                                block[r, c] = CellText(Convert.ToString(dataRow[c]));
                        }
                        WriteBlock(sheet, start + 2, block);
                    }
                }

                Call(workbook, "SaveAs", path);

                Set(excel, "DisplayAlerts", true);
                Set(excel, "Visible", true);
            }
            catch
            {
                try { if (workbook != null) Call(workbook, "Close", false); } catch { /* ignore */ }
                try { Call(excel, "Quit"); } catch { /* ignore */ }
                throw;
            }
            finally
            {
                // COM nesnelerini serbest bırak (Excel kullanıcı için açık kalır; yalnızca bizim referanslarımız bırakılır).
                ReleaseCom(sheet);
                ReleaseCom(sheets);
                ReleaseCom(workbook);
                ReleaseCom(workbooks);
                ReleaseCom(excel);
            }
        }

        // data[0,0], startRow satırı / 1. sütuna gelecek şekilde tek çağrıyla yazılır.
        private static void WriteBlock(object sheet, int startRow, object[,] data)
        {
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            if (rows == 0 || cols == 0) return;

            object first = null, last = null, range = null;
            try
            {
                first = Get(sheet, "Cells", startRow, 1);
                last = Get(sheet, "Cells", startRow + rows - 1, cols);
                range = Get(sheet, "Range", first, last);
                Set(range, "Value2", data);
            }
            finally
            {
                ReleaseCom(range);
                ReleaseCom(last);
                ReleaseCom(first);
            }
        }

        private static object Get(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target,
                args != null && args.Length > 0 ? args : null);
        }

        private static void Set(object target, string name, object value)
        {
            target.GetType().InvokeMember(name, BindingFlags.SetProperty, null, target, new object[] { value });
        }

        private static object Call(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target,
                args != null && args.Length > 0 ? args : null);
        }

        private static void ReleaseCom(object obj)
        {
            if (obj == null) return;
            try
            {
                if (Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch { /* ignore */ }
        }
    }
}