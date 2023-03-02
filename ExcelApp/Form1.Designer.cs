namespace ExcelAPP
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.labelNameFolder = new System.Windows.Forms.Label();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.labelCompleted = new System.Windows.Forms.Label();
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
            this.AutoPageBreakeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.PagesInPartBookLabel = new System.Windows.Forms.Label();
            this.partsBookCheckBox = new System.Windows.Forms.CheckBox();
            this.pagesInPartBookNumeric = new System.Windows.Forms.NumericUpDown();
            this.RdPdToggle = new ExcelApp.Controls.ToggleButton();
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CountPagePZNumeric)).BeginInit();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pagesInPartBookNumeric)).BeginInit();
            this.SuspendLayout();
            // 
            // labelNameFolder
            // 
            this.labelNameFolder.AutoSize = true;
            this.labelNameFolder.Location = new System.Drawing.Point(17, 41);
            this.labelNameFolder.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.labelNameFolder.MaximumSize = new System.Drawing.Size(821, 20);
            this.labelNameFolder.MinimumSize = new System.Drawing.Size(21, 20);
            this.labelNameFolder.Name = "labelNameFolder";
            this.labelNameFolder.Size = new System.Drawing.Size(21, 20);
            this.labelNameFolder.TabIndex = 1;
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Location = new System.Drawing.Point(491, 351);
            this.btnSelectFolder.Margin = new System.Windows.Forms.Padding(4);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(180, 28);
            this.btnSelectFolder.TabIndex = 3;
            this.btnSelectFolder.Text = "Выбрать папку";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.BtnSelectFolder_Click);
            // 
            // labelCompleted
            // 
            this.labelCompleted.AutoSize = true;
            this.labelCompleted.Location = new System.Drawing.Point(14, 385);
            this.labelCompleted.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.labelCompleted.MaximumSize = new System.Drawing.Size(821, 20);
            this.labelCompleted.MinimumSize = new System.Drawing.Size(21, 20);
            this.labelCompleted.Name = "labelCompleted";
            this.labelCompleted.Size = new System.Drawing.Size(21, 20);
            this.labelCompleted.TabIndex = 7;
            // 
            // btnBuild
            // 
            this.btnBuild.Location = new System.Drawing.Point(711, 351);
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
            this.StartNumberNumeric.Location = new System.Drawing.Point(201, 70);
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
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 73);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(143, 16);
            this.label1.TabIndex = 11;
            this.label1.Text = "Начать нумерацию с";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 110);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 16);
            this.label4.TabIndex = 17;
            this.label4.Text = "Страниц в ПЗ";
            // 
            // CountPagePZNumeric
            // 
            this.CountPagePZNumeric.Location = new System.Drawing.Point(201, 108);
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
            // 
            // TwoSidedPrintCheckBox
            // 
            this.TwoSidedPrintCheckBox.AutoSize = true;
            this.TwoSidedPrintCheckBox.Checked = true;
            this.TwoSidedPrintCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TwoSidedPrintCheckBox.Location = new System.Drawing.Point(20, 212);
            this.TwoSidedPrintCheckBox.Margin = new System.Windows.Forms.Padding(4);
            this.TwoSidedPrintCheckBox.Name = "TwoSidedPrintCheckBox";
            this.TwoSidedPrintCheckBox.Size = new System.Drawing.Size(163, 20);
            this.TwoSidedPrintCheckBox.TabIndex = 19;
            this.TwoSidedPrintCheckBox.Text = "Двустронняя печать";
            this.TwoSidedPrintCheckBox.UseMnemonic = false;
            this.TwoSidedPrintCheckBox.UseVisualStyleBackColor = true;
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
            this.SplitBookContentCheckBox.Location = new System.Drawing.Point(20, 238);
            this.SplitBookContentCheckBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.SplitBookContentCheckBox.Name = "SplitBookContentCheckBox";
            this.SplitBookContentCheckBox.Size = new System.Drawing.Size(176, 20);
            this.SplitBookContentCheckBox.TabIndex = 20;
            this.SplitBookContentCheckBox.Text = "Содержание отдельно";
            this.SplitBookContentCheckBox.UseVisualStyleBackColor = true;
            // 
            // infoTextBox
            // 
            this.infoTextBox.BackColor = System.Drawing.SystemColors.Menu;
            this.infoTextBox.Location = new System.Drawing.Point(293, 70);
            this.infoTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.infoTextBox.Multiline = true;
            this.infoTextBox.Name = "infoTextBox";
            this.infoTextBox.ReadOnly = true;
            this.infoTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.infoTextBox.Size = new System.Drawing.Size(547, 260);
            this.infoTextBox.TabIndex = 21;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(60, 160);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 20);
            this.label2.TabIndex = 25;
            this.label2.Text = "ПД";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(169, 160);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(33, 20);
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
            this.menuStrip1.Size = new System.Drawing.Size(864, 28);
            this.menuStrip1.TabIndex = 28;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AutoPageBreakeToolStripMenuItem});
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(98, 24);
            this.settingsToolStripMenuItem.Text = "Настройки";
            // 
            // AutoPageBreakeToolStripMenuItem
            // 
            this.AutoPageBreakeToolStripMenuItem.Checked = true;
            this.AutoPageBreakeToolStripMenuItem.CheckOnClick = true;
            this.AutoPageBreakeToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AutoPageBreakeToolStripMenuItem.Name = "AutoPageBreakeToolStripMenuItem";
            this.AutoPageBreakeToolStripMenuItem.Size = new System.Drawing.Size(268, 26);
            this.AutoPageBreakeToolStripMenuItem.Text = "Авторазделение страниц";
            // 
            // PagesInPartBookLabel
            // 
            this.PagesInPartBookLabel.AutoSize = true;
            this.PagesInPartBookLabel.Location = new System.Drawing.Point(17, 290);
            this.PagesInPartBookLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.PagesInPartBookLabel.Name = "PagesInPartBookLabel";
            this.PagesInPartBookLabel.Size = new System.Drawing.Size(155, 16);
            this.PagesInPartBookLabel.TabIndex = 30;
            this.PagesInPartBookLabel.Text = "Страниц в части книги";
            // 
            // partsBookCheckBox
            // 
            this.partsBookCheckBox.AutoSize = true;
            this.partsBookCheckBox.Location = new System.Drawing.Point(20, 262);
            this.partsBookCheckBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.partsBookCheckBox.Name = "partsBookCheckBox";
            this.partsBookCheckBox.Size = new System.Drawing.Size(170, 20);
            this.partsBookCheckBox.TabIndex = 31;
            this.partsBookCheckBox.Text = "Разделение на части";
            this.partsBookCheckBox.UseVisualStyleBackColor = true;
            // 
            // pagesInPartBookNumeric
            // 
            this.pagesInPartBookNumeric.Location = new System.Drawing.Point(201, 290);
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
            // 
            // RdPdToggle
            // 
            this.RdPdToggle.AutoSize = true;
            this.RdPdToggle.Location = new System.Drawing.Point(103, 158);
            this.RdPdToggle.Margin = new System.Windows.Forms.Padding(4);
            this.RdPdToggle.MinimumSize = new System.Drawing.Size(60, 27);
            this.RdPdToggle.Name = "RdPdToggle";
            this.RdPdToggle.OffBackColor = System.Drawing.Color.Gray;
            this.RdPdToggle.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.RdPdToggle.OnBackColor = System.Drawing.Color.RoyalBlue;
            this.RdPdToggle.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.RdPdToggle.Size = new System.Drawing.Size(60, 27);
            this.RdPdToggle.TabIndex = 23;
            this.RdPdToggle.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.ClientSize = new System.Drawing.Size(864, 423);
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
            this.Controls.Add(this.labelCompleted);
            this.Controls.Add(this.btnBuild);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.labelNameFolder);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Smeta++";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CountPagePZNumeric)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pagesInPartBookNumeric)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label labelNameFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.Button btnBuild;
        private System.Windows.Forms.NumericUpDown StartNumberNumeric;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown CountPagePZNumeric;
        private System.Windows.Forms.CheckBox TwoSidedPrintCheckBox;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
        private System.Windows.Forms.CheckBox SplitBookContentCheckBox;
        private System.Windows.Forms.TextBox infoTextBox;
        private ExcelApp.Controls.ToggleButton RdPdToggle;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.Label labelCompleted;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem AutoPageBreakeToolStripMenuItem;
        private System.Windows.Forms.Label PagesInPartBookLabel;
        private System.Windows.Forms.CheckBox partsBookCheckBox;
        private System.Windows.Forms.NumericUpDown pagesInPartBookNumeric;
    }
}