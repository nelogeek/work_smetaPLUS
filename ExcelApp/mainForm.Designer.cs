namespace ExcelAPP
{
    partial class MainForm
    {

        private System.ComponentModel.IContainer components = null;


        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows


        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.labelNameFolder = new System.Windows.Forms.Label();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.labelProgressStage = new System.Windows.Forms.Label();
            this.btnBuild = new System.Windows.Forms.Button();
            this.StartNumberNumeric = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.CountPagePZNumeric = new System.Windows.Forms.NumericUpDown();
            this.TwoSidedPrintCheckBox = new System.Windows.Forms.CheckBox();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.SplitBookContentCheckBox = new System.Windows.Forms.CheckBox();
            this.infoTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AutoPageBreakerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AutoBooksPartPassCheckBox = new System.Windows.Forms.ToolStripMenuItem();
            this.PagesInPartBookLabel = new System.Windows.Forms.Label();
            this.partsBookCheckBox = new System.Windows.Forms.CheckBox();
            this.pagesInPartBookNumeric = new System.Windows.Forms.NumericUpDown();
            this.dividerPassPagesCount = new System.Windows.Forms.NumericUpDown();
            this.dividerPagesCountLabel = new System.Windows.Forms.Label();
            this.buildProgressBar = new System.Windows.Forms.ProgressBar();
            this.btnReBuild = new System.Windows.Forms.Button();
            this.backgroundWorker2 = new System.ComponentModel.BackgroundWorker();
            this.cbxType = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.RdPdToggle = new ExcelApp.Controls.ToggleButton();
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CountPagePZNumeric)).BeginInit();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pagesInPartBookNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dividerPassPagesCount)).BeginInit();
            this.SuspendLayout();
            // 
            // labelNameFolder
            // 
            this.labelNameFolder.AutoSize = true;
            this.labelNameFolder.Location = new System.Drawing.Point(17, 38);
            this.labelNameFolder.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.labelNameFolder.MaximumSize = new System.Drawing.Size(885, 20);
            this.labelNameFolder.MinimumSize = new System.Drawing.Size(21, 20);
            this.labelNameFolder.Name = "labelNameFolder";
            this.labelNameFolder.Size = new System.Drawing.Size(21, 20);
            this.labelNameFolder.TabIndex = 1;
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Location = new System.Drawing.Point(562, 362);
            this.btnSelectFolder.Margin = new System.Windows.Forms.Padding(4);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(180, 28);
            this.btnSelectFolder.TabIndex = 3;
            this.btnSelectFolder.Text = "Выбрать папку";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.BtnSelectFolder_Click);
            // 
            // labelProgressStage
            // 
            this.labelProgressStage.AutoSize = true;
            this.labelProgressStage.Location = new System.Drawing.Point(14, 403);
            this.labelProgressStage.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.labelProgressStage.MaximumSize = new System.Drawing.Size(885, 20);
            this.labelProgressStage.MinimumSize = new System.Drawing.Size(21, 20);
            this.labelProgressStage.Name = "labelProgressStage";
            this.labelProgressStage.Size = new System.Drawing.Size(21, 20);
            this.labelProgressStage.TabIndex = 7;
            // 
            // btnBuild
            // 
            this.btnBuild.Location = new System.Drawing.Point(766, 362);
            this.btnBuild.Margin = new System.Windows.Forms.Padding(4);
            this.btnBuild.Name = "btnBuild";
            this.btnBuild.Size = new System.Drawing.Size(129, 28);
            this.btnBuild.TabIndex = 5;
            this.btnBuild.Text = "Собрать книгу";
            this.btnBuild.UseVisualStyleBackColor = true;
            this.btnBuild.Click += new System.EventHandler(this.BtnBuild_Click);
            // 
            // StartNumberNumeric
            // 
            this.StartNumberNumeric.Location = new System.Drawing.Point(243, 66);
            this.StartNumberNumeric.Margin = new System.Windows.Forms.Padding(4, 4, 13, 12);
            this.StartNumberNumeric.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.StartNumberNumeric.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.StartNumberNumeric.Name = "StartNumberNumeric";
            this.StartNumberNumeric.Size = new System.Drawing.Size(76, 22);
            this.StartNumberNumeric.TabIndex = 10;
            this.StartNumberNumeric.Value = new decimal(new int[] {
            3,
            0,
            0,
            0});
            this.StartNumberNumeric.ValueChanged += new System.EventHandler(this.StartNumberNumeric_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 68);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(143, 16);
            this.label1.TabIndex = 11;
            this.label1.Text = "Начать нумерацию с";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 106);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 16);
            this.label4.TabIndex = 17;
            this.label4.Text = "Страниц в ПЗ";
            // 
            // CountPagePZNumeric
            // 
            this.CountPagePZNumeric.Location = new System.Drawing.Point(243, 104);
            this.CountPagePZNumeric.Margin = new System.Windows.Forms.Padding(4, 4, 13, 12);
            this.CountPagePZNumeric.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.CountPagePZNumeric.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.CountPagePZNumeric.Name = "CountPagePZNumeric";
            this.CountPagePZNumeric.Size = new System.Drawing.Size(76, 22);
            this.CountPagePZNumeric.TabIndex = 16;
            this.CountPagePZNumeric.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.CountPagePZNumeric.ValueChanged += new System.EventHandler(this.CountPagePZNumeric_ValueChanged);
            // 
            // TwoSidedPrintCheckBox
            // 
            this.TwoSidedPrintCheckBox.AutoSize = true;
            this.TwoSidedPrintCheckBox.Checked = true;
            this.TwoSidedPrintCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TwoSidedPrintCheckBox.Location = new System.Drawing.Point(20, 215);
            this.TwoSidedPrintCheckBox.Margin = new System.Windows.Forms.Padding(4);
            this.TwoSidedPrintCheckBox.Name = "TwoSidedPrintCheckBox";
            this.TwoSidedPrintCheckBox.Size = new System.Drawing.Size(163, 20);
            this.TwoSidedPrintCheckBox.TabIndex = 19;
            this.TwoSidedPrintCheckBox.Text = "Двустронняя печать";
            this.TwoSidedPrintCheckBox.UseMnemonic = false;
            this.TwoSidedPrintCheckBox.UseVisualStyleBackColor = true;
            this.TwoSidedPrintCheckBox.CheckedChanged += new System.EventHandler(this.TwoSidedPrintCheckBox_CheckedChanged);
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BackgroundWorker_DoWork);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.BackgroundWorker_ProgressChanged);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BackgroundWorker_RunWorkerCompleted);
            // 
            // SplitBookContentCheckBox
            // 
            this.SplitBookContentCheckBox.AutoSize = true;
            this.SplitBookContentCheckBox.Checked = true;
            this.SplitBookContentCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.SplitBookContentCheckBox.Enabled = false;
            this.SplitBookContentCheckBox.Location = new System.Drawing.Point(20, 241);
            this.SplitBookContentCheckBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.SplitBookContentCheckBox.Name = "SplitBookContentCheckBox";
            this.SplitBookContentCheckBox.Size = new System.Drawing.Size(176, 20);
            this.SplitBookContentCheckBox.TabIndex = 20;
            this.SplitBookContentCheckBox.Text = "Содержание отдельно";
            this.SplitBookContentCheckBox.UseVisualStyleBackColor = true;
            this.SplitBookContentCheckBox.CheckedChanged += new System.EventHandler(this.SplitBookContentCheckBox_CheckedChanged);
            // 
            // infoTextBox
            // 
            this.infoTextBox.BackColor = System.Drawing.SystemColors.Menu;
            this.infoTextBox.Location = new System.Drawing.Point(335, 66);
            this.infoTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.infoTextBox.Multiline = true;
            this.infoTextBox.Name = "infoTextBox";
            this.infoTextBox.ReadOnly = true;
            this.infoTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.infoTextBox.Size = new System.Drawing.Size(562, 285);
            this.infoTextBox.TabIndex = 21;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(17, 328);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 18);
            this.label2.TabIndex = 25;
            this.label2.Text = "ПД";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(110, 328);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 18);
            this.label3.TabIndex = 26;
            this.label3.Text = "РД";
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(219)))), ((int)(((byte)(219)))), ((int)(((byte)(219)))));
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.settingsToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(5, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(909, 30);
            this.menuStrip1.TabIndex = 28;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AutoPageBreakerToolStripMenuItem,
            this.AutoBooksPartPassCheckBox});
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(98, 26);
            this.settingsToolStripMenuItem.Text = "Настройки";
            // 
            // AutoPageBreakerToolStripMenuItem
            // 
            this.AutoPageBreakerToolStripMenuItem.Checked = true;
            this.AutoPageBreakerToolStripMenuItem.CheckOnClick = true;
            this.AutoPageBreakerToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AutoPageBreakerToolStripMenuItem.Name = "AutoPageBreakerToolStripMenuItem";
            this.AutoPageBreakerToolStripMenuItem.Size = new System.Drawing.Size(541, 26);
            this.AutoPageBreakerToolStripMenuItem.Text = "Авторазделение страниц";
            // 
            // AutoBooksPartPassCheckBox
            // 
            this.AutoBooksPartPassCheckBox.Checked = true;
            this.AutoBooksPartPassCheckBox.CheckOnClick = true;
            this.AutoBooksPartPassCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AutoBooksPartPassCheckBox.Name = "AutoBooksPartPassCheckBox";
            this.AutoBooksPartPassCheckBox.Size = new System.Drawing.Size(541, 26);
            this.AutoBooksPartPassCheckBox.Text = "Автоматичесий подбор допуска страниц к разделению на части";
            this.AutoBooksPartPassCheckBox.Click += new System.EventHandler(this.AutoBooksPartPassCheckBox_Click);
            // 
            // PagesInPartBookLabel
            // 
            this.PagesInPartBookLabel.AutoSize = true;
            this.PagesInPartBookLabel.Location = new System.Drawing.Point(17, 144);
            this.PagesInPartBookLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.PagesInPartBookLabel.Name = "PagesInPartBookLabel";
            this.PagesInPartBookLabel.Size = new System.Drawing.Size(155, 16);
            this.PagesInPartBookLabel.TabIndex = 30;
            this.PagesInPartBookLabel.Text = "Страниц в части книги";
            // 
            // partsBookCheckBox
            // 
            this.partsBookCheckBox.AutoSize = true;
            this.partsBookCheckBox.Checked = true;
            this.partsBookCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.partsBookCheckBox.Location = new System.Drawing.Point(20, 265);
            this.partsBookCheckBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.partsBookCheckBox.Name = "partsBookCheckBox";
            this.partsBookCheckBox.Size = new System.Drawing.Size(170, 20);
            this.partsBookCheckBox.TabIndex = 31;
            this.partsBookCheckBox.Text = "Разделение на части";
            this.partsBookCheckBox.UseVisualStyleBackColor = true;
            this.partsBookCheckBox.CheckedChanged += new System.EventHandler(this.partsBookCheckBox_CheckedChanged);
            // 
            // pagesInPartBookNumeric
            // 
            this.pagesInPartBookNumeric.Location = new System.Drawing.Point(243, 142);
            this.pagesInPartBookNumeric.Margin = new System.Windows.Forms.Padding(4, 4, 13, 12);
            this.pagesInPartBookNumeric.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.pagesInPartBookNumeric.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.pagesInPartBookNumeric.Name = "pagesInPartBookNumeric";
            this.pagesInPartBookNumeric.Size = new System.Drawing.Size(76, 22);
            this.pagesInPartBookNumeric.TabIndex = 32;
            this.pagesInPartBookNumeric.Value = new decimal(new int[] {
            400,
            0,
            0,
            0});
            this.pagesInPartBookNumeric.ValueChanged += new System.EventHandler(this.pagesInPartBookNumeric_ValueChanged);
            // 
            // dividerPassPagesCount
            // 
            this.dividerPassPagesCount.Enabled = false;
            this.dividerPassPagesCount.Location = new System.Drawing.Point(243, 180);
            this.dividerPassPagesCount.Margin = new System.Windows.Forms.Padding(4, 4, 13, 12);
            this.dividerPassPagesCount.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.dividerPassPagesCount.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.dividerPassPagesCount.Name = "dividerPassPagesCount";
            this.dividerPassPagesCount.Size = new System.Drawing.Size(76, 22);
            this.dividerPassPagesCount.TabIndex = 34;
            this.dividerPassPagesCount.Value = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.dividerPassPagesCount.ValueChanged += new System.EventHandler(this.dividerPassPagesCount_ValueChanged);
            // 
            // dividerPagesCountLabel
            // 
            this.dividerPagesCountLabel.AutoSize = true;
            this.dividerPagesCountLabel.Location = new System.Drawing.Point(17, 182);
            this.dividerPagesCountLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.dividerPagesCountLabel.Name = "dividerPagesCountLabel";
            this.dividerPagesCountLabel.Size = new System.Drawing.Size(178, 16);
            this.dividerPagesCountLabel.TabIndex = 35;
            this.dividerPagesCountLabel.Text = "Допуск страниц для части";
            // 
            // buildProgressBar
            // 
            this.buildProgressBar.Location = new System.Drawing.Point(0, 426);
            this.buildProgressBar.Name = "buildProgressBar";
            this.buildProgressBar.Size = new System.Drawing.Size(909, 10);
            this.buildProgressBar.TabIndex = 37;
            this.buildProgressBar.Visible = false;
            // 
            // btnReBuild
            // 
            this.btnReBuild.Enabled = false;
            this.btnReBuild.Location = new System.Drawing.Point(20, 362);
            this.btnReBuild.Name = "btnReBuild";
            this.btnReBuild.Size = new System.Drawing.Size(163, 28);
            this.btnReBuild.TabIndex = 38;
            this.btnReBuild.Text = "Пересобрать книгу";
            this.btnReBuild.UseVisualStyleBackColor = true;
            this.btnReBuild.Click += new System.EventHandler(this.btnReBuild_Click);
            // 
            // backgroundWorker2
            // 
            this.backgroundWorker2.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BackgroundWorker2_DoWork);
            this.backgroundWorker2.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.BackgroundWorker2_ProgressChanged);
            this.backgroundWorker2.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BackgroundWorker2_RunWorkerCompleted);
            // 
            // cbxType
            // 
            this.cbxType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxType.FormattingEnabled = true;
            this.cbxType.Items.AddRange(new object[] {
            "Лукойл",
            "Газпром"});
            this.cbxType.Location = new System.Drawing.Point(82, 293);
            this.cbxType.Name = "cbxType";
            this.cbxType.Size = new System.Drawing.Size(140, 24);
            this.cbxType.TabIndex = 39;
            this.cbxType.SelectedIndexChanged += new System.EventHandler(this.cbxType_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(17, 296);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 16);
            this.label5.TabIndex = 40;
            this.label5.Text = "Объект:";
            // 
            // RdPdToggle
            // 
            this.RdPdToggle.AutoSize = true;
            this.RdPdToggle.Location = new System.Drawing.Point(55, 328);
            this.RdPdToggle.Margin = new System.Windows.Forms.Padding(4);
            this.RdPdToggle.MinimumSize = new System.Drawing.Size(48, 21);
            this.RdPdToggle.Name = "RdPdToggle";
            this.RdPdToggle.OffBackColor = System.Drawing.Color.Gray;
            this.RdPdToggle.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.RdPdToggle.OnBackColor = System.Drawing.Color.RoyalBlue;
            this.RdPdToggle.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.RdPdToggle.Size = new System.Drawing.Size(48, 21);
            this.RdPdToggle.TabIndex = 23;
            this.RdPdToggle.UseVisualStyleBackColor = true;
            this.RdPdToggle.CheckedChanged += new System.EventHandler(this.RdPdToggle_CheckedChanged);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.ClientSize = new System.Drawing.Size(909, 434);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cbxType);
            this.Controls.Add(this.btnReBuild);
            this.Controls.Add(this.buildProgressBar);
            this.Controls.Add(this.dividerPagesCountLabel);
            this.Controls.Add(this.dividerPassPagesCount);
            this.Controls.Add(this.pagesInPartBookNumeric);
            this.Controls.Add(this.partsBookCheckBox);
            this.Controls.Add(this.PagesInPartBookLabel);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.RdPdToggle);
            this.Controls.Add(this.infoTextBox);
            this.Controls.Add(this.SplitBookContentCheckBox);
            this.Controls.Add(this.TwoSidedPrintCheckBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CountPagePZNumeric);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StartNumberNumeric);
            this.Controls.Add(this.labelProgressStage);
            this.Controls.Add(this.btnBuild);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.labelNameFolder);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Smeta++";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CountPagePZNumeric)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pagesInPartBookNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dividerPassPagesCount)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        public System.ComponentModel.BackgroundWorker backgroundWorker;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.Label labelProgressStage;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.Label PagesInPartBookLabel;
        private System.Windows.Forms.Label dividerPagesCountLabel;
        public System.Windows.Forms.Label labelNameFolder;
        public System.Windows.Forms.TextBox infoTextBox;
        public System.Windows.Forms.Button btnSelectFolder;
        public System.Windows.Forms.Button btnBuild;
        public System.Windows.Forms.NumericUpDown StartNumberNumeric;
        public System.Windows.Forms.NumericUpDown CountPagePZNumeric;
        public System.Windows.Forms.CheckBox TwoSidedPrintCheckBox;
        public System.Windows.Forms.CheckBox SplitBookContentCheckBox;
        public ExcelApp.Controls.ToggleButton RdPdToggle;
        public System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        public System.Windows.Forms.ToolStripMenuItem AutoPageBreakerToolStripMenuItem;
        public System.Windows.Forms.CheckBox partsBookCheckBox;
        public System.Windows.Forms.NumericUpDown pagesInPartBookNumeric;
        public System.Windows.Forms.ToolStripMenuItem AutoBooksPartPassCheckBox;
        public System.Windows.Forms.NumericUpDown dividerPassPagesCount;
        private System.Windows.Forms.ProgressBar buildProgressBar;
        public System.ComponentModel.BackgroundWorker backgroundWorker2;
        public System.Windows.Forms.Button btnReBuild;
        public System.Windows.Forms.ComboBox cbxType;
        private System.Windows.Forms.Label label5;
    }
}