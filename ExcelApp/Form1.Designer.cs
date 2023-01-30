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
            this.labelNameFolder = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.labelCompleted = new System.Windows.Forms.Label();
            this.btnBuild = new System.Windows.Forms.Button();
            this.infoTextBox = new System.Windows.Forms.TextBox();
            this.StartNumberTextBox = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.numericUpDown2 = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.CountPagePZ = new System.Windows.Forms.NumericUpDown();
            this.TwoSidedPrintCheckBox = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberTextBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CountPagePZ)).BeginInit();
            this.SuspendLayout();
            // 
            // labelNameFolder
            // 
            this.labelNameFolder.AutoSize = true;
            this.labelNameFolder.Location = new System.Drawing.Point(13, 19);
            this.labelNameFolder.Name = "labelNameFolder";
            this.labelNameFolder.Size = new System.Drawing.Size(0, 13);
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
            this.labelCompleted.Size = new System.Drawing.Size(0, 13);
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
            // infoTextBox
            // 
            this.infoTextBox.Location = new System.Drawing.Point(221, 41);
            this.infoTextBox.Multiline = true;
            this.infoTextBox.Name = "infoTextBox";
            this.infoTextBox.ReadOnly = true;
            this.infoTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.infoTextBox.Size = new System.Drawing.Size(409, 216);
            this.infoTextBox.TabIndex = 6;
            // 
            // StartNumberTextBox
            // 
            this.StartNumberTextBox.Location = new System.Drawing.Point(151, 57);
            this.StartNumberTextBox.Margin = new System.Windows.Forms.Padding(3, 3, 10, 10);
            this.StartNumberTextBox.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.StartNumberTextBox.Name = "StartNumberTextBox";
            this.StartNumberTextBox.Size = new System.Drawing.Size(57, 20);
            this.StartNumberTextBox.TabIndex = 10;
            this.StartNumberTextBox.Value = new decimal(new int[] {
            3,
            0,
            0,
            0});
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 59);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(132, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Номер первой страницы";
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(151, 90);
            this.numericUpDown1.Margin = new System.Windows.Forms.Padding(3, 3, 10, 10);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(57, 20);
            this.numericUpDown1.TabIndex = 12;
            this.numericUpDown1.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 92);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(90, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Страниц в книге";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 125);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(126, 13);
            this.label3.TabIndex = 15;
            this.label3.Text = "Стр. после содержания";
            // 
            // numericUpDown2
            // 
            this.numericUpDown2.Location = new System.Drawing.Point(151, 123);
            this.numericUpDown2.Margin = new System.Windows.Forms.Padding(3, 3, 10, 10);
            this.numericUpDown2.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.numericUpDown2.Name = "numericUpDown2";
            this.numericUpDown2.Size = new System.Drawing.Size(57, 20);
            this.numericUpDown2.TabIndex = 14;
            this.numericUpDown2.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 158);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 13);
            this.label4.TabIndex = 17;
            this.label4.Text = "Страниц в ПЗ";
            // 
            // CountPagePZ
            // 
            this.CountPagePZ.Location = new System.Drawing.Point(151, 156);
            this.CountPagePZ.Margin = new System.Windows.Forms.Padding(3, 3, 10, 10);
            this.CountPagePZ.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.CountPagePZ.Name = "CountPagePZ";
            this.CountPagePZ.Size = new System.Drawing.Size(57, 20);
            this.CountPagePZ.TabIndex = 16;
            this.CountPagePZ.Value = new decimal(new int[] {
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
            this.TwoSidedPrintCheckBox.Size = new System.Drawing.Size(130, 17);
            this.TwoSidedPrintCheckBox.TabIndex = 19;
            this.TwoSidedPrintCheckBox.Text = "Двустронняя печать";
            this.TwoSidedPrintCheckBox.UseMnemonic = false;
            this.TwoSidedPrintCheckBox.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(642, 323);
            this.Controls.Add(this.TwoSidedPrintCheckBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CountPagePZ);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.numericUpDown2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.numericUpDown1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StartNumberTextBox);
            this.Controls.Add(this.labelCompleted);
            this.Controls.Add(this.infoTextBox);
            this.Controls.Add(this.btnBuild);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.labelNameFolder);
            this.Name = "Form1";
            this.Text = "Smeta++";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.StartNumberTextBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CountPagePZ)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label labelNameFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.Button btnBuild;
        private System.Windows.Forms.TextBox infoTextBox;
        private System.Windows.Forms.Label labelCompleted;
        private System.Windows.Forms.NumericUpDown StartNumberTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.NumericUpDown numericUpDown2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown CountPagePZ;
        private System.Windows.Forms.CheckBox TwoSidedPrintCheckBox;
    }
}
