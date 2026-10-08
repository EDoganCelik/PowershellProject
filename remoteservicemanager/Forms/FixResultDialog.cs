using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RemoteServiceManager.Models;
using RemoteServiceManager.Services;

namespace RemoteServiceManager.Forms
{
    public class FixResultDialog : Form
    {
        public FixResultDialog(IList<HostFixSummary> data)
        {
            Text = "Düzeltme Sonuç Raporu";
            Size = new Size(900, 500);
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(700, 400);
            Font = new Font("Segoe UI", 9F);

            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = true,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText,
                BackgroundColor = Color.White
            };
            dgv.Columns.Add("Hostname", "Hostname");
            dgv.Columns.Add("Status", "Status");
            foreach (var d in data)
                dgv.Rows.Add(d.Hostname, d.Status);

            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 48,
                Padding = new Padding(8),
                FlowDirection = FlowDirection.LeftToRight
            };

            var btnCopy = new Button { Text = "Panoya Kopyala", Width = 140, Height = 30 };
            var btnCsv = new Button { Text = "CSV / Excel Aktar", Width = 140, Height = 30 };
            var btnClose = new Button { Text = "Kapat", Width = 100, Height = 30, DialogResult = DialogResult.OK };

            btnCopy.Click += (s, e) =>
            {
                var sb = new StringBuilder();
                sb.AppendLine("Hostname\tStatus");
                foreach (var d in data)
                    sb.AppendLine(d.Hostname + "\t" + d.Status);
                Clipboard.SetText(sb.ToString());
                MessageBox.Show(this, "Tablo panoya kopyalandı.", "Bilgi");
            };

            btnCsv.Click += (s, e) =>
                ExcelExporter.ExportSummariesToExcel(data, "ServisDuzeltmeRaporu_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));

            panel.Controls.AddRange(new Control[] { btnCopy, btnCsv, btnClose });
            Controls.Add(dgv);
            Controls.Add(panel);
            AcceptButton = btnClose;
        }
    }
}
