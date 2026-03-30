namespace TeamCalendar
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            pnlHeader = new Panel();
            lblTitle = new Label();
            menuStrip = new MenuStrip();
            mnuFile = new ToolStripMenuItem();
            mnuFileExport = new ToolStripMenuItem();
            mnuFileSep1 = new ToolStripSeparator();
            mnuFileExit = new ToolStripMenuItem();
            mnuHelp = new ToolStripMenuItem();
            mnuHelpUsage = new ToolStripMenuItem();
            mnuHelpSep1 = new ToolStripSeparator();
            mnuHelpAbout = new ToolStripMenuItem();
            pnlToolbar = new Panel();
            lblStartDate = new Label();
            dtpStart = new DateTimePicker();
            lblDateSeparator = new Label();
            dtpEnd = new DateTimePicker();
            btnLoad = new Button();
            btnExport = new Button();
            chkDebugLog = new CheckBox();
            pnlSummary = new FlowLayoutPanel();
            pnlChart = new Panel();
            chkIncludeTentative = new CheckBox();
            pnlUserInput = new Panel();
            lblUserEmails = new Label();
            txtUserEmails = new TextBox();
            chkIncludeSelf = new CheckBox();
            pnlGridWrapper = new Panel();
            splitMain = new SplitContainer();
            dgvAppointments = new DataGridView();
            txtLog = new TextBox();
            statusStrip = new StatusStrip();
            lblStatus = new ToolStripStatusLabel();

            pnlHeader.SuspendLayout();
            pnlToolbar.SuspendLayout();
            pnlUserInput.SuspendLayout();
            pnlGridWrapper.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitMain).BeginInit();
            splitMain.Panel1.SuspendLayout();
            splitMain.Panel2.SuspendLayout();
            splitMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvAppointments).BeginInit();
            statusStrip.SuspendLayout();
            SuspendLayout();

            //
            // pnlHeader
            //
            pnlHeader.BackColor = Color.FromArgb(0, 120, 212);
            pnlHeader.Controls.Add(lblTitle);
            pnlHeader.Dock = DockStyle.Top;
            pnlHeader.Location = new Point(0, 0);
            pnlHeader.Size = new Size(1100, 52);

            //
            // lblTitle
            //
            lblTitle.AutoSize = true;
            lblTitle.Font = new Font("Segoe UI", 15F, FontStyle.Bold);
            lblTitle.ForeColor = Color.White;
            lblTitle.Location = new Point(20, 12);
            lblTitle.Text = "\U0001f4c5  Team Calendar";

            //
            // pnlToolbar
            //
            pnlToolbar.BackColor = Color.White;
            pnlToolbar.Controls.Add(lblStartDate);
            pnlToolbar.Controls.Add(dtpStart);
            pnlToolbar.Controls.Add(lblDateSeparator);
            pnlToolbar.Controls.Add(dtpEnd);
            pnlToolbar.Controls.Add(btnLoad);
            pnlToolbar.Controls.Add(btnExport);
            pnlToolbar.Controls.Add(chkDebugLog);
            pnlToolbar.Dock = DockStyle.Top;
            pnlToolbar.Location = new Point(0, 52);
            pnlToolbar.Padding = new Padding(20, 0, 20, 0);
            pnlToolbar.Size = new Size(1100, 56);

            //
            // menuStrip
            //
            menuStrip.BackColor = Color.FromArgb(0, 110, 200);
            menuStrip.Font = new Font("Segoe UI", 9F);
            menuStrip.ForeColor = Color.White;
            menuStrip.Items.AddRange(new ToolStripItem[] { mnuFile, mnuHelp });
            menuStrip.Location = new Point(0, 52);
            menuStrip.Padding = new Padding(8, 2, 0, 2);
            menuStrip.Size = new Size(1100, 24);
            menuStrip.Renderer = new MenuStripRenderer();

            //
            // mnuFile
            //
            mnuFile.DropDownItems.AddRange(new ToolStripItem[] { mnuFileExport, mnuFileSep1, mnuFileExit });
            mnuFile.ForeColor = Color.White;
            mnuFile.Text = "ファイル(&F)";

            //
            // mnuFileExport
            //
            mnuFileExport.Text = "📊  Excel出力 (承認済み)(&E)";
            mnuFileExport.ShortcutKeys = Keys.Control | Keys.E;
            mnuFileExport.Click += btnExport_Click;

            //
            // mnuFileSep1
            //

            //
            // mnuFileExit
            //
            mnuFileExit.Text = "終了(&X)";
            mnuFileExit.ShortcutKeys = Keys.Alt | Keys.F4;
            mnuFileExit.Click += (s, e) => Close();

            //
            // mnuHelp
            //
            mnuHelp.DropDownItems.AddRange(new ToolStripItem[] { mnuHelpUsage, mnuHelpSep1, mnuHelpAbout });
            mnuHelp.ForeColor = Color.White;
            mnuHelp.Text = "ヘルプ(&H)";

            //
            // mnuHelpUsage
            //
            mnuHelpUsage.Text = "📖  使い方(&U)";
            mnuHelpUsage.ShortcutKeys = Keys.F1;
            mnuHelpUsage.Click += mnuHelpUsage_Click;

            //
            // mnuHelpSep1
            //

            //
            // mnuHelpAbout
            //
            mnuHelpAbout.Text = "ℹ️  バージョン情報(&A)";
            mnuHelpAbout.Click += mnuHelpAbout_Click;

            //
            // lblStartDate
            //
            lblStartDate.AutoSize = true;
            lblStartDate.Font = new Font("Segoe UI", 9F);
            lblStartDate.ForeColor = Color.FromArgb(96, 96, 96);
            lblStartDate.Location = new Point(20, 19);
            lblStartDate.Text = "期間";

            //
            // dtpStart
            //
            dtpStart.CalendarFont = new Font("Segoe UI", 9F);
            dtpStart.Font = new Font("Segoe UI", 9.5F);
            dtpStart.Format = DateTimePickerFormat.Short;
            dtpStart.Location = new Point(58, 15);
            dtpStart.Size = new Size(130, 25);

            //
            // lblDateSeparator
            //
            lblDateSeparator.AutoSize = true;
            lblDateSeparator.Font = new Font("Segoe UI", 10F);
            lblDateSeparator.ForeColor = Color.FromArgb(96, 96, 96);
            lblDateSeparator.Location = new Point(194, 18);
            lblDateSeparator.Text = "〜";

            //
            // dtpEnd
            //
            dtpEnd.CalendarFont = new Font("Segoe UI", 9F);
            dtpEnd.Font = new Font("Segoe UI", 9.5F);
            dtpEnd.Format = DateTimePickerFormat.Short;
            dtpEnd.Location = new Point(220, 15);
            dtpEnd.Size = new Size(130, 25);

            //
            // btnLoad
            //
            btnLoad.BackColor = Color.FromArgb(0, 120, 212);
            btnLoad.Cursor = Cursors.Hand;
            btnLoad.FlatAppearance.BorderSize = 0;
            btnLoad.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 90, 158);
            btnLoad.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 69, 120);
            btnLoad.FlatStyle = FlatStyle.Flat;
            btnLoad.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnLoad.ForeColor = Color.White;
            btnLoad.Location = new Point(380, 12);
            btnLoad.Size = new Size(120, 32);
            btnLoad.Text = "▶  予定を取得";
            btnLoad.UseVisualStyleBackColor = false;
            btnLoad.Click += btnLoad_Click;

            //
            // btnExport
            //
            btnExport.BackColor = Color.FromArgb(16, 124, 16);
            btnExport.Cursor = Cursors.Hand;
            btnExport.FlatAppearance.BorderSize = 0;
            btnExport.FlatAppearance.MouseOverBackColor = Color.FromArgb(12, 100, 12);
            btnExport.FlatAppearance.MouseDownBackColor = Color.FromArgb(8, 76, 8);
            btnExport.FlatStyle = FlatStyle.Flat;
            btnExport.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnExport.ForeColor = Color.White;
            btnExport.Location = new Point(515, 12);
            btnExport.Size = new Size(170, 32);
            btnExport.Text = "📊  Excel出力 (承認済み)";
            btnExport.UseVisualStyleBackColor = false;
            btnExport.Click += btnExport_Click;

            //
            // chkDebugLog
            //
            chkDebugLog.AutoSize = true;
            chkDebugLog.Font = new Font("Segoe UI", 8.5F);
            chkDebugLog.ForeColor = Color.FromArgb(96, 96, 96);
            chkDebugLog.Location = new Point(710, 19);
            chkDebugLog.Text = "🔍 デバッグログ";
            chkDebugLog.UseVisualStyleBackColor = true;
            chkDebugLog.CheckedChanged += chkDebugLog_CheckedChanged;

            //
            // pnlUserInput
            //
            pnlUserInput.BackColor = Color.White;
            pnlUserInput.Controls.Add(lblUserEmails);
            pnlUserInput.Controls.Add(txtUserEmails);
            pnlUserInput.Controls.Add(chkIncludeSelf);
            pnlUserInput.Dock = DockStyle.Top;
            pnlUserInput.Location = new Point(0, 108);
            pnlUserInput.Padding = new Padding(20, 0, 20, 0);
            pnlUserInput.Size = new Size(1100, 44);

            //
            // lblUserEmails
            //
            lblUserEmails.AutoSize = true;
            lblUserEmails.Font = new Font("Segoe UI", 9F);
            lblUserEmails.ForeColor = Color.FromArgb(96, 96, 96);
            lblUserEmails.Location = new Point(20, 13);
            lblUserEmails.Text = "\U0001f465 対象ユーザー";

            //
            // txtUserEmails
            //
            txtUserEmails.BorderStyle = BorderStyle.FixedSingle;
            txtUserEmails.Font = new Font("Segoe UI", 9.5F);
            txtUserEmails.Location = new Point(140, 9);
            txtUserEmails.Size = new Size(420, 25);

            //
            // chkIncludeSelf
            //
            chkIncludeSelf.AutoSize = true;
            chkIncludeSelf.Checked = true;
            chkIncludeSelf.CheckState = CheckState.Checked;
            chkIncludeSelf.Font = new Font("Segoe UI", 9F);
            chkIncludeSelf.ForeColor = Color.FromArgb(96, 96, 96);
            chkIncludeSelf.Location = new Point(580, 12);
            chkIncludeSelf.Text = "自分の予定を含める";
            chkIncludeSelf.UseVisualStyleBackColor = true;

            //
            // pnlSummary
            //
            pnlSummary.AutoSize = false;
            pnlSummary.BackColor = Color.FromArgb(243, 243, 243);
            pnlSummary.Dock = DockStyle.Top;
            pnlSummary.Location = new Point(0, 152);
            pnlSummary.Padding = new Padding(16, 10, 16, 10);
            pnlSummary.Size = new Size(1100, 90);
            pnlSummary.WrapContents = false;

            //
            // chkIncludeTentative
            //
            chkIncludeTentative.AutoSize = true;
            chkIncludeTentative.Font = new Font("Segoe UI", 8.5F);
            chkIncludeTentative.ForeColor = Color.FromArgb(96, 96, 96);
            chkIncludeTentative.Location = new Point(340, 8);
            chkIncludeTentative.Text = "⏳ 任意 (仮) も含める";
            chkIncludeTentative.UseVisualStyleBackColor = true;
            chkIncludeTentative.BackColor = Color.White;
            chkIncludeTentative.CheckedChanged += chkIncludeTentative_CheckedChanged;

            //
            // pnlChart
            //
            pnlChart.BackColor = Color.FromArgb(243, 243, 243);
            pnlChart.Controls.Add(chkIncludeTentative);
            pnlChart.Dock = DockStyle.Top;
            pnlChart.Size = new Size(1100, 210);

            //
            // pnlGridWrapper
            //
            pnlGridWrapper.BackColor = Color.FromArgb(243, 243, 243);
            pnlGridWrapper.Controls.Add(splitMain);
            pnlGridWrapper.Dock = DockStyle.Fill;
            pnlGridWrapper.Padding = new Padding(16, 4, 16, 8);

            //
            // splitMain
            //
            splitMain.BackColor = Color.White;
            splitMain.Dock = DockStyle.Fill;
            splitMain.Orientation = Orientation.Horizontal;
            splitMain.SplitterDistance = 370;
            splitMain.SplitterWidth = 6;
            splitMain.Panel1.Controls.Add(dgvAppointments);
            splitMain.Panel1.BackColor = Color.White;
            splitMain.Panel2.Controls.Add(txtLog);
            splitMain.Panel2Collapsed = true;

            //
            // dgvAppointments
            //
            dgvAppointments.AllowUserToAddRows = false;
            dgvAppointments.AllowUserToDeleteRows = false;
            dgvAppointments.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvAppointments.BackgroundColor = Color.White;
            dgvAppointments.BorderStyle = BorderStyle.None;
            dgvAppointments.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgvAppointments.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgvAppointments.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dgvAppointments.ColumnHeadersHeight = 40;
            dgvAppointments.Dock = DockStyle.Fill;
            dgvAppointments.EnableHeadersVisualStyles = false;
            dgvAppointments.GridColor = Color.FromArgb(235, 235, 235);
            dgvAppointments.ReadOnly = true;
            dgvAppointments.RowHeadersVisible = false;
            dgvAppointments.RowTemplate.Height = 32;
            dgvAppointments.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dgvAppointments.CellDoubleClick += dgvAppointments_CellDoubleClick;

            //
            // txtLog
            //
            txtLog.BackColor = Color.FromArgb(30, 30, 30);
            txtLog.BorderStyle = BorderStyle.None;
            txtLog.Dock = DockStyle.Fill;
            txtLog.Font = new Font("Cascadia Code", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtLog.ForeColor = Color.FromArgb(204, 204, 204);
            txtLog.Multiline = true;
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Both;
            txtLog.WordWrap = false;

            //
            // statusStrip
            //
            statusStrip.BackColor = Color.White;
            statusStrip.Items.AddRange(new ToolStripItem[] { lblStatus });
            statusStrip.Location = new Point(0, 628);
            statusStrip.Size = new Size(1100, 24);
            statusStrip.SizingGrip = false;

            //
            // lblStatus
            //
            lblStatus.Font = new Font("Segoe UI", 8.5F);
            lblStatus.ForeColor = Color.FromArgb(96, 96, 96);
            lblStatus.Text = "Outlookの予定を取得するには「▶ 予定を取得」ボタンを押してください";

            //
            // Form1
            //
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(243, 243, 243);
            ClientSize = new Size(1100, 820);
            Controls.Add(pnlGridWrapper);
            Controls.Add(pnlChart);
            Controls.Add(pnlSummary);
            Controls.Add(pnlUserInput);
            Controls.Add(pnlToolbar);
            Controls.Add(menuStrip);
            Controls.Add(pnlHeader);
            Controls.Add(statusStrip);
            Font = new Font("Segoe UI", 9F);
            MainMenuStrip = menuStrip;
            MinimumSize = new Size(900, 600);
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Team Calendar";

            pnlHeader.ResumeLayout(false);
            pnlHeader.PerformLayout();
            pnlToolbar.ResumeLayout(false);
            pnlToolbar.PerformLayout();
            pnlUserInput.ResumeLayout(false);
            pnlUserInput.PerformLayout();
            pnlGridWrapper.ResumeLayout(false);
            splitMain.Panel1.ResumeLayout(false);
            splitMain.Panel2.ResumeLayout(false);
            splitMain.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)splitMain).EndInit();
            splitMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvAppointments).EndInit();
            statusStrip.ResumeLayout(false);
            statusStrip.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Panel pnlHeader;
        private Label lblTitle;
        private MenuStrip menuStrip;
        private ToolStripMenuItem mnuFile;
        private ToolStripMenuItem mnuFileExport;
        private ToolStripSeparator mnuFileSep1;
        private ToolStripMenuItem mnuFileExit;
        private ToolStripMenuItem mnuHelp;
        private ToolStripMenuItem mnuHelpUsage;
        private ToolStripSeparator mnuHelpSep1;
        private ToolStripMenuItem mnuHelpAbout;
        private Panel pnlToolbar;
        private Label lblStartDate;
        private DateTimePicker dtpStart;
        private Label lblDateSeparator;
        private DateTimePicker dtpEnd;
        private Button btnLoad;
        private Button btnExport;
        private CheckBox chkDebugLog;
        private Panel pnlUserInput;
        private Label lblUserEmails;
        private TextBox txtUserEmails;
        private CheckBox chkIncludeSelf;
        private FlowLayoutPanel pnlSummary;
        private Panel pnlChart;
        private CheckBox chkIncludeTentative;
        private Panel pnlGridWrapper;
        private SplitContainer splitMain;
        private DataGridView dgvAppointments;
        private TextBox txtLog;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel lblStatus;
    }
}
