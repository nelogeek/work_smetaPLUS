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
            this.label3 = new System.Windows.Forms.Label();
            this.afterTitleNumeric = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.CountPagePZNumeric = new System.Windows.Forms.NumericUpDown();
            this.TwoSidedPrintCheckBox = new System.Windows.Forms.CheckBox();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.SplitBookContentCheckBox = new System.Windows.Forms.CheckBox();
            this.infoTextBox = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.toggleButton1 = new ExcelApp.Controls.ToggleButton();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.afterTitleNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CountPagePZNumeric)).BeginInit();
            this.SuspendLayout();
            // 
            // labelNameFolder
            // 
            this.labelNameFolder.AutoSize = true;
            this.labelNameFolder.Location = new System.Drawing.Point(13, 19);
            this.labelNameFolder.Name = "labelNameFolder";
            this.labelNameFolder.Size = new System.Drawing.Size(0, 15);
            this.labelNameFolder.TabIndex = 1;
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Location = new System.Drawing.Point(370, 285);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(146, 23);
            this.btnSelectFolder.TabIndex = 3;
            this.btnSelectFolder.Text = "Выбрать папку";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.BtnSelectFolder_Click);
            // 
            // labelCompleted
            // 
            this.labelCompleted.AutoSize = true;
            this.labelCompleted.Location = new System.Drawing.Point(13, 290);
            this.labelCompleted.Name = "labelCompleted";
            this.labelCompleted.Size = new System.Drawing.Size(0, 15);
            this.labelCompleted.TabIndex = 7;
            // 
            // btnBuild
            // 
            this.btnBuild.Location = new System.Drawing.Point(555, 285);
            this.btnBuild.Name = "btnBuild";
            this.btnBuild.Size = new System.Drawing.Size(75, 23);
            this.btnBuild.TabIndex = 5;
            this.btnBuild.Text = "Собрать";
            this.btnBuild.UseVisualStyleBackColor = true;
            this.btnBuild.Click += new System.EventHandler(this.BtnBuild_Click);
            // 
            // StartNumberNumeric
            // 
            this.StartNumberNumeric.Location = new System.Drawing.Point(151, 57);
            this.StartNumberNumeric.Margin = new System.Windows.Forms.Padding(3, 3, 10, 10);
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
            this.StartNumberNumeric.Size = new System.Drawing.Size(57, 20);
            this.StartNumberNumeric.TabIndex = 10;
            this.StartNumberNumeric.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 59);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(151, 15);
            this.label1.TabIndex = 11;
            this.label1.Text = "Номер первой страницы";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 89);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(143, 15);
            this.label3.TabIndex = 15;
            this.label3.Text = "Стр. после содержания";
            // 
            // afterTitleNumeric
            // 
            this.afterTitleNumeric.Enabled = false;
            this.afterTitleNumeric.Location = new System.Drawing.Point(151, 88);
            this.afterTitleNumeric.Margin = new System.Windows.Forms.Padding(3, 3, 10, 10);
            this.afterTitleNumeric.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.afterTitleNumeric.Name = "afterTitleNumeric";
            this.afterTitleNumeric.Size = new System.Drawing.Size(57, 20);
            this.afterTitleNumeric.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(14, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(87, 15);
            this.label4.TabIndex = 17;
            this.label4.Text = "Страниц в ПЗ";
            // 
            // CountPagePZNumeric
            // 
            this.CountPagePZNumeric.Location = new System.Drawing.Point(151, 119);
            this.CountPagePZNumeric.Margin = new System.Windows.Forms.Padding(3, 3, 10, 10);
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
            this.CountPagePZNumeric.Size = new System.Drawing.Size(57, 20);
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
            this.TwoSidedPrintCheckBox.Location = new System.Drawing.Point(16, 211);
            this.TwoSidedPrintCheckBox.Name = "TwoSidedPrintCheckBox";
            this.TwoSidedPrintCheckBox.Size = new System.Drawing.Size(149, 19);
            this.TwoSidedPrintCheckBox.TabIndex = 19;
            this.TwoSidedPrintCheckBox.Text = "Двустронняя печать";
            this.TwoSidedPrintCheckBox.UseMnemonic = false;
            this.TwoSidedPrintCheckBox.UseVisualStyleBackColor = true;
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // SplitBookContentCheckBox
            // 
            this.SplitBookContentCheckBox.AutoSize = true;
            this.SplitBookContentCheckBox.Checked = true;
            this.SplitBookContentCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.SplitBookContentCheckBox.Location = new System.Drawing.Point(16, 234);
            this.SplitBookContentCheckBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SplitBookContentCheckBox.Name = "SplitBookContentCheckBox";
            this.SplitBookContentCheckBox.Size = new System.Drawing.Size(161, 19);
            this.SplitBookContentCheckBox.TabIndex = 20;
            this.SplitBookContentCheckBox.Text = "Содержание отдельно";
            this.SplitBookContentCheckBox.UseVisualStyleBackColor = true;
            // 
            // infoTextBox
            // 
            this.infoTextBox.BackColor = System.Drawing.SystemColors.Menu;
            this.infoTextBox.Location = new System.Drawing.Point(220, 57);
            this.infoTextBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.infoTextBox.Multiline = true;
            this.infoTextBox.Name = "infoTextBox";
            this.infoTextBox.ReadOnly = true;
            this.infoTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.infoTextBox.Size = new System.Drawing.Size(411, 212);
            this.infoTextBox.TabIndex = 21;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(346, 47);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(8, 8);
            this.button1.TabIndex = 22;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // toggleButton1
            // 
            this.toggleButton1.AutoSize = true;
            this.toggleButton1.Location = new System.Drawing.Point(31, 330);
            this.toggleButton1.MinimumSize = new System.Drawing.Size(45, 22);
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.OffBackColor = System.Drawing.Color.Gray;
            this.toggleButton1.OffToggleColor = System.Drawing.Color.Gainsboro;
            this.toggleButton1.OnBackColor = System.Drawing.Color.MediumSlateBlue;
            this.toggleButton1.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.toggleButton1.Size = new System.Drawing.Size(45, 22);
            this.toggleButton1.TabIndex = 23;
            this.toggleButton1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(136, 351);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 24;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(856, 398);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.toggleButton1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.infoTextBox);
            this.Controls.Add(this.SplitBookContentCheckBox);
            this.Controls.Add(this.TwoSidedPrintCheckBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CountPagePZNumeric);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.afterTitleNumeric);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StartNumberNumeric);
            this.Controls.Add(this.labelCompleted);
            this.Controls.Add(this.btnBuild);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.labelNameFolder);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Smeta++";
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.afterTitleNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CountPagePZNumeric)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label labelNameFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.Button btnBuild;
        private System.Windows.Forms.Label labelCompleted;
        private System.Windows.Forms.NumericUpDown StartNumberNumeric;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.NumericUpDown afterTitleNumeric;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown CountPagePZNumeric;
        private System.Windows.Forms.CheckBox TwoSidedPrintCheckBox;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
        private System.Windows.Forms.CheckBox SplitBookContentCheckBox;
        private System.Windows.Forms.TextBox infoTextBox;
        private System.Windows.Forms.Button button1;
        private ExcelApp.Controls.ToggleButton toggleButton1;
        private System.Windows.Forms.Button button2;
    }
}
