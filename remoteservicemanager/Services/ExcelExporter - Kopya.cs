using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RemoteServiceManager.Services
{
    public static class ExcelExporter
    {
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
                    var csvFallback = Path.ChangeExtension(sfd.FileName, ".csv");
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
            foreach (var d in data ?? Enumerable.Empty<Models.HostFixSummary>())
                dt.Rows.Add(d.Hostname, d.Status);
            ExportDataTableToExcelAndOpen(dt, defaultFileName);
        }

        private static void ExportCsv(DataTable table, string path)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.Join(";", table.Columns.Cast<DataColumn>().Select(c => EscapeCsv(c.ColumnName))));
            foreach (DataRow row in table.Rows)
            {
                sb.AppendLine(string.Join(";", table.Columns.Cast<DataColumn>()
                    .Select(c => EscapeCsv(Convert.ToString(row[c])))));
            }
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        private static string EscapeCsv(string value)
        {
            value = value ?? string.Empty;
            if (value.Contains(";") || value.Contains("\"") || value.Contains("\n"))
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
        }

        private static void ExportViaExcelCom(DataTable table, string path)
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("Microsoft Excel yüklü değil.");

            object excel = Activator.CreateInstance(excelType);
            try
            {
                excelType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, excel, new object[] { false });
                var workbooks = excelType.InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, excel, null);
                var workbook = workbooks.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, workbooks, null);
                var sheets = workbook.GetType().InvokeMember("Worksheets", System.Reflection.BindingFlags.GetProperty, null, workbook, null);
                var sheet = sheets.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, sheets, new object[] { 1 });

                for (int c = 0; c < table.Columns.Count; c++)
                    SetCell(sheet, 1, c + 1, table.Columns[c].ColumnName);

                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Columns.Count; c++)
                        SetCell(sheet, r + 2, c + 1, Convert.ToString(table.Rows[r][c]));
                }

                workbook.GetType().InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod, null, workbook,
                    new object[] { path });
                excelType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, excel, new object[] { true });
            }
            catch
            {
                try
                {
                    excelType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, excel, null);
                }
                catch { /* ignore */ }
                throw;
            }
        }

        private static void SetCell(object sheet, int row, int col, string value)
        {
            var cells = sheet.GetType().InvokeMember("Cells", System.Reflection.BindingFlags.GetProperty, null, sheet,
                new object[] { row, col });
            cells.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, cells,
                new object[] { value ?? string.Empty });
        }
    }
}
