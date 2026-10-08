namespace RemoteServiceManager
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.menuNav = new System.Windows.Forms.MenuStrip();
            this.mnuDosya = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuExport = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuImport = new System.Windows.Forms.ToolStripMenuItem();
            this.pnlNav = new System.Windows.Forms.Panel();
            this.chkShowService = new System.Windows.Forms.CheckBox();
            this.chkShowPatch = new System.Windows.Forms.CheckBox();
            this.lblNavPanels = new System.Windows.Forms.Label();
            this.splitMain = new System.Windows.Forms.SplitContainer();
            this.pnlComputerList = new System.Windows.Forms.Panel();
            this.lblComputers = new System.Windows.Forms.Label();
            this.txtComputers = new System.Windows.Forms.TextBox();
            this.pnlShared = new System.Windows.Forms.Panel();
            this.lblStatus = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.tabs = new System.Windows.Forms.TabControl();
            this.tabService = new System.Windows.Forms.TabPage();
            this.tabPatch = new System.Windows.Forms.TabPage();
            this.menuNav.SuspendLayout();
            this.pnlNav.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitMain)).BeginInit();
            this.splitMain.Panel1.SuspendLayout();
            this.splitMain.Panel2.SuspendLayout();
            this.splitMain.SuspendLayout();
            this.pnlShared.SuspendLayout();
            this.tabs.SuspendLayout();
            this.SuspendLayout();
            //
            // menuNav
            //
            this.menuNav.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                this.mnuDosya});
            this.menuNav.Location = new System.Drawing.Point(0, 0);
            this.menuNav.Name = "menuNav";
            this.menuNav.Size = new System.Drawing.Size(1360, 24);
            this.menuNav.TabIndex = 1;
            this.menuNav.Text = "menuNav";
            //
            // mnuDosya
            //
            this.mnuDosya.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
                this.mnuExport,
                this.mnuImport});
            this.mnuDosya.Name = "mnuDosya";
            this.mnuDosya.Text = "Dosya";
            //
            // mnuExport
            //
            this.mnuExport.Name = "mnuExport";
            this.mnuExport.Text = "Yapılandırmayı Dışa Aktar...";
            this.mnuExport.Click += new System.EventHandler(this.mnuExport_Click);
            //
            // mnuImport
            //
            this.mnuImport.Name = "mnuImport";
            this.mnuImport.Text = "Yapılandırmayı İçe Aktar...";
            this.mnuImport.Click += new System.EventHandler(this.mnuImport_Click);
            //
            // pnlNav
            //
            this.pnlNav.BackColor = System.Drawing.Color.FromArgb(245, 247, 250);
            this.pnlNav.Controls.Add(this.lblNavPanels);
            this.pnlNav.Controls.Add(this.chkShowService);
            this.pnlNav.Controls.Add(this.chkShowPatch);
            this.pnlNav.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlNav.Height = 32;
            this.pnlNav.Name = "pnlNav";
            //
            // lblNavPanels
            //
            this.lblNavPanels.AutoSize = true;
            this.lblNavPanels.Location = new System.Drawing.Point(8, 8);
            this.lblNavPanels.Name = "lblNavPanels";
            this.lblNavPanels.Text = "Paneller:";
            //
            // chkShowService
            //
            this.chkShowService.AutoSize = true;
            this.chkShowService.Checked = true;
            this.chkShowService.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkShowService.Location = new System.Drawing.Point(70, 7);
            this.chkShowService.Name = "chkShowService";
            this.chkShowService.Text = "Servis / Sistem";
            this.chkShowService.CheckedChanged += new System.EventHandler(this.chkShowPanel_CheckedChanged);
            //
            // chkShowPatch
            //
            this.chkShowPatch.AutoSize = true;
            this.chkShowPatch.Checked = true;
            this.chkShowPatch.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkShowPatch.Location = new System.Drawing.Point(200, 7);
            this.chkShowPatch.Name = "chkShowPatch";
            this.chkShowPatch.Text = "Patch Takip / Güncelleme";
            this.chkShowPatch.CheckedChanged += new System.EventHandler(this.chkShowPanel_CheckedChanged);
            //
            // splitMain
            //
            this.splitMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitMain.Location = new System.Drawing.Point(0, 56);
            this.splitMain.Name = "splitMain";
            this.splitMain.Panel1MinSize = 120;
            this.splitMain.Panel2MinSize = 400;
            this.splitMain.Size = new System.Drawing.Size(1360, 764);
            this.splitMain.SplitterDistance = 200;
            this.splitMain.TabIndex = 0;
            //
            // pnlComputerList
            //
            this.pnlComputerList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlComputerList.Padding = new System.Windows.Forms.Padding(8);
            this.pnlComputerList.Controls.Add(this.txtComputers);
            this.pnlComputerList.Controls.Add(this.lblComputers);
            //
            // lblComputers
            //
            this.lblComputers.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblComputers.Name = "lblComputers";
            this.lblComputers.Padding = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.lblComputers.Text = "Bilgisayarlar (satır satır)";
            //
            // txtComputers
            //
            this.txtComputers.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtComputers.Multiline = true;
            this.txtComputers.Name = "txtComputers";
            this.txtComputers.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtComputers.TabIndex = 0;
            this.splitMain.Panel1.Controls.Add(this.pnlComputerList);
            //
            // pnlShared
            //
            this.pnlShared.Controls.Add(this.lblStatus);
            this.pnlShared.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlShared.Height = 28;
            this.pnlShared.Name = "pnlShared";
            this.pnlShared.Padding = new System.Windows.Forms.Padding(8, 4, 8, 4);
            //
            // lblStatus
            //
            this.lblStatus.AutoEllipsis = true;
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(20, 60, 120);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Text = "Hazır.";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // progressBar
            //
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar.Height = 14;
            this.progressBar.Name = "progressBar";
            //
            // tabs
            //
            this.tabs.Controls.Add(this.tabService);
            this.tabs.Controls.Add(this.tabPatch);
            this.tabs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabs.Name = "tabs";
            this.tabService.Text = "Servis / Sistem";
            this.tabService.Padding = new System.Windows.Forms.Padding(4);
            this.tabPatch.Text = "Patch Takip";
            this.tabPatch.Padding = new System.Windows.Forms.Padding(4);
            //
            // Panel2 (Dock sırası: önce eklenen en son yerleşir — alt/üst paneller önce)
            //
            this.splitMain.Panel2.Controls.Add(this.tabs);
            this.splitMain.Panel2.Controls.Add(this.pnlShared);
            this.splitMain.Panel2.Controls.Add(this.progressBar);
            //
            // MainForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1360, 820);
            this.Controls.Add(this.splitMain);
            this.Controls.Add(this.pnlNav);
            this.Controls.Add(this.menuNav);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MainMenuStrip = this.menuNav;
            this.MinimumSize = new System.Drawing.Size(1100, 700);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Remote Service Manager";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.menuNav.ResumeLayout(false);
            this.menuNav.PerformLayout();
            this.pnlNav.ResumeLayout(false);
            this.pnlNav.PerformLayout();
            this.splitMain.Panel1.ResumeLayout(false);
            this.splitMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitMain)).EndInit();
            this.splitMain.ResumeLayout(false);
            this.pnlShared.ResumeLayout(false);
            this.tabs.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.MenuStrip menuNav;
        private System.Windows.Forms.ToolStripMenuItem mnuDosya;
        private System.Windows.Forms.ToolStripMenuItem mnuExport;
        private System.Windows.Forms.ToolStripMenuItem mnuImport;
        private System.Windows.Forms.Panel pnlNav;
        private System.Windows.Forms.Label lblNavPanels;
        private System.Windows.Forms.CheckBox chkShowService;
        private System.Windows.Forms.CheckBox chkShowPatch;
        private System.Windows.Forms.SplitContainer splitMain;
        private System.Windows.Forms.Panel pnlComputerList;
        private System.Windows.Forms.Label lblComputers;
        private System.Windows.Forms.TextBox txtComputers;
        private System.Windows.Forms.Panel pnlShared;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.TabControl tabs;
        private System.Windows.Forms.TabPage tabService;
        private System.Windows.Forms.TabPage tabPatch;
    }
}
