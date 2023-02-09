﻿namespace ExcelAPP
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
            this.numberPagesInBook = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.afterTitleNumeric = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.CountPagePZNumeric = new System.Windows.Forms.NumericUpDown();
            this.TwoSidedPrintCheckBox = new System.Windows.Forms.CheckBox();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.SplitBookContentCheckBox = new System.Windows.Forms.CheckBox();
            this.infoTextBox = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numberPagesInBook)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.afterTitleNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CountPagePZNumeric)).BeginInit();
            this.SuspendLayout();
            // 
            // labelNameFolder
            // 
            this.labelNameFolder.AutoSize = true;
            this.labelNameFolder.Location = new System.Drawing.Point(17, 23);
            this.labelNameFolder.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelNameFolder.Name = "labelNameFolder";
            this.labelNameFolder.Size = new System.Drawing.Size(0, 16);
            this.labelNameFolder.TabIndex = 1;
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Location = new System.Drawing.Point(493, 351);
            this.btnSelectFolder.Margin = new System.Windows.Forms.Padding(4);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(195, 28);
            this.btnSelectFolder.TabIndex = 3;
            this.btnSelectFolder.Text = "Выбрать папку";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.BtnSelectFolder_Click);
            // 
            // labelCompleted
            // 
            this.labelCompleted.AutoSize = true;
            this.labelCompleted.Location = new System.Drawing.Point(17, 357);
            this.labelCompleted.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelCompleted.Name = "labelCompleted";
            this.labelCompleted.Size = new System.Drawing.Size(0, 16);
            this.labelCompleted.TabIndex = 7;
            // 
            // btnBuild
            // 
            this.btnBuild.Location = new System.Drawing.Point(740, 351);
            this.btnBuild.Margin = new System.Windows.Forms.Padding(4);
            this.btnBuild.Name = "btnBuild";
            this.btnBuild.Size = new System.Drawing.Size(100, 28);
            this.btnBuild.TabIndex = 5;
            this.btnBuild.Text = "Собрать";
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
            this.StartNumberNumeric.Name = "StartNumberNumeric";
            this.StartNumberNumeric.Size = new System.Drawing.Size(76, 22);
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
            this.label1.Location = new System.Drawing.Point(17, 73);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(167, 16);
            this.label1.TabIndex = 11;
            this.label1.Text = "Номер первой страницы";
            // 
            // numberPagesInBook
            // 
            this.numberPagesInBook.Location = new System.Drawing.Point(201, 111);
            this.numberPagesInBook.Margin = new System.Windows.Forms.Padding(4, 4, 13, 12);
            this.numberPagesInBook.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.numberPagesInBook.Name = "numberPagesInBook";
            this.numberPagesInBook.Size = new System.Drawing.Size(76, 22);
            this.numberPagesInBook.TabIndex = 12;
            this.numberPagesInBook.Value = new decimal(new int[] {
            500,
            0,
            0,
            0});
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 113);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(114, 16);
            this.label2.TabIndex = 13;
            this.label2.Text = "Страниц в книге";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 154);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(158, 16);
            this.label3.TabIndex = 15;
            this.label3.Text = "Стр. после содержания";
            // 
            // afterTitleNumeric
            // 
            this.afterTitleNumeric.Location = new System.Drawing.Point(201, 151);
            this.afterTitleNumeric.Margin = new System.Windows.Forms.Padding(4, 4, 13, 12);
            this.afterTitleNumeric.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.afterTitleNumeric.Name = "afterTitleNumeric";
            this.afterTitleNumeric.Size = new System.Drawing.Size(76, 22);
            this.afterTitleNumeric.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 194);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 16);
            this.label4.TabIndex = 17;
            this.label4.Text = "Страниц в ПЗ";
            // 
            // CountPagePZNumeric
            // 
            this.CountPagePZNumeric.Location = new System.Drawing.Point(201, 192);
            this.CountPagePZNumeric.Margin = new System.Windows.Forms.Padding(4, 4, 13, 12);
            this.CountPagePZNumeric.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.CountPagePZNumeric.Name = "CountPagePZNumeric";
            this.CountPagePZNumeric.Size = new System.Drawing.Size(76, 22);
            this.CountPagePZNumeric.TabIndex = 16;
            // 
            // TwoSidedPrintCheckBox
            // 
            this.TwoSidedPrintCheckBox.AutoSize = true;
            this.TwoSidedPrintCheckBox.Location = new System.Drawing.Point(21, 260);
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
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // SplitBookContentCheckBox
            // 
            this.SplitBookContentCheckBox.AutoSize = true;
            this.SplitBookContentCheckBox.Checked = true;
            this.SplitBookContentCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.SplitBookContentCheckBox.Location = new System.Drawing.Point(21, 288);
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
            this.infoTextBox.Multiline = true;
            this.infoTextBox.Name = "infoTextBox";
            this.infoTextBox.ReadOnly = true;
            this.infoTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.infoTextBox.Size = new System.Drawing.Size(547, 260);
            this.infoTextBox.TabIndex = 21;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(856, 398);
            this.Controls.Add(this.infoTextBox);
            this.Controls.Add(this.SplitBookContentCheckBox);
            this.Controls.Add(this.TwoSidedPrintCheckBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CountPagePZNumeric);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.afterTitleNumeric);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.numberPagesInBook);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StartNumberNumeric);
            this.Controls.Add(this.labelCompleted);
            this.Controls.Add(this.btnBuild);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.labelNameFolder);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Smeta++";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numberPagesInBook)).EndInit();
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
        private System.Windows.Forms.NumericUpDown numberPagesInBook;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.NumericUpDown afterTitleNumeric;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown CountPagePZNumeric;
        private System.Windows.Forms.CheckBox TwoSidedPrintCheckBox;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
        private System.Windows.Forms.CheckBox SplitBookContentCheckBox;
        private System.Windows.Forms.TextBox infoTextBox;
    }
}
