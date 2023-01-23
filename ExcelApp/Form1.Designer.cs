namespace ExcelApp
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.infoTextBox = new System.Windows.Forms.TextBox();
            this.labelNameFolder = new System.Windows.Forms.Label();
            this.labelCompleted = new System.Windows.Forms.Label();
            this.startNumberTextBox = new System.Windows.Forms.TextBox();
            this.BtnSelectFolder = new System.Windows.Forms.Button();
            this.BtnBuild = new System.Windows.Forms.Button();
            this.Read_Button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // infoTextBox
            // 
            this.infoTextBox.Location = new System.Drawing.Point(273, 12);
            this.infoTextBox.Multiline = true;
            this.infoTextBox.Name = "infoTextBox";
            this.infoTextBox.Size = new System.Drawing.Size(515, 426);
            this.infoTextBox.TabIndex = 0;
            // 
            // labelNameFolder
            // 
            this.labelNameFolder.AutoSize = true;
            this.labelNameFolder.Location = new System.Drawing.Point(23, 36);
            this.labelNameFolder.Name = "labelNameFolder";
            this.labelNameFolder.Size = new System.Drawing.Size(44, 16);
            this.labelNameFolder.TabIndex = 1;
            this.labelNameFolder.Text = "label1";
            // 
            // labelCompleted
            // 
            this.labelCompleted.AutoSize = true;
            this.labelCompleted.Location = new System.Drawing.Point(26, 407);
            this.labelCompleted.Name = "labelCompleted";
            this.labelCompleted.Size = new System.Drawing.Size(44, 16);
            this.labelCompleted.TabIndex = 2;
            this.labelCompleted.Text = "label1";
            // 
            // startNumberTextBox
            // 
            this.startNumberTextBox.Location = new System.Drawing.Point(26, 69);
            this.startNumberTextBox.Name = "startNumberTextBox";
            this.startNumberTextBox.Size = new System.Drawing.Size(100, 22);
            this.startNumberTextBox.TabIndex = 3;
            // 
            // BtnSelectFolder
            // 
            this.BtnSelectFolder.Location = new System.Drawing.Point(26, 107);
            this.BtnSelectFolder.Name = "BtnSelectFolder";
            this.BtnSelectFolder.Size = new System.Drawing.Size(75, 23);
            this.BtnSelectFolder.TabIndex = 4;
            this.BtnSelectFolder.Text = "Выбрать";
            this.BtnSelectFolder.UseVisualStyleBackColor = true;
            this.BtnSelectFolder.Click += new System.EventHandler(this.BtnSelectFolder_Click);
            // 
            // BtnBuild
            // 
            this.BtnBuild.Location = new System.Drawing.Point(146, 107);
            this.BtnBuild.Name = "BtnBuild";
            this.BtnBuild.Size = new System.Drawing.Size(75, 23);
            this.BtnBuild.TabIndex = 5;
            this.BtnBuild.Text = "Собрать";
            this.BtnBuild.UseVisualStyleBackColor = true;
            this.BtnBuild.Click += new System.EventHandler(this.BtnBuild_Click);
            // 
            // Read_Button
            // 
            this.Read_Button.Location = new System.Drawing.Point(29, 150);
            this.Read_Button.Name = "Read_Button";
            this.Read_Button.Size = new System.Drawing.Size(97, 26);
            this.Read_Button.TabIndex = 6;
            this.Read_Button.Text = "Прочитать";
            this.Read_Button.UseVisualStyleBackColor = true;
            this.Read_Button.Click += new System.EventHandler(this.Read_Button_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.Read_Button);
            this.Controls.Add(this.BtnBuild);
            this.Controls.Add(this.BtnSelectFolder);
            this.Controls.Add(this.startNumberTextBox);
            this.Controls.Add(this.labelCompleted);
            this.Controls.Add(this.labelNameFolder);
            this.Controls.Add(this.infoTextBox);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox infoTextBox;
        private System.Windows.Forms.Label labelNameFolder;
        private System.Windows.Forms.Label labelCompleted;
        private System.Windows.Forms.TextBox startNumberTextBox;
        private System.Windows.Forms.Button BtnSelectFolder;
        private System.Windows.Forms.Button BtnBuild;
        private System.Windows.Forms.Button Read_Button;
    }
}

