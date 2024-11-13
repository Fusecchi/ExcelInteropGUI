namespace ExcelInteropGUI
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SendButton = new System.Windows.Forms.Button();
            this.DataSelected = new System.Windows.Forms.Label();
            this.FileType = new System.Windows.Forms.TextBox();
            this.FileName = new System.Windows.Forms.TextBox();
            this.EditButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // SendButton
            // 
            this.SendButton.Location = new System.Drawing.Point(38, 26);
            this.SendButton.Name = "SendButton";
            this.SendButton.Size = new System.Drawing.Size(83, 31);
            this.SendButton.TabIndex = 0;
            this.SendButton.Text = "Select File";
            this.SendButton.UseVisualStyleBackColor = true;
            this.SendButton.Click += new System.EventHandler(this.SendButton_Click);
            // 
            // DataSelected
            // 
            this.DataSelected.AutoSize = true;
            this.DataSelected.Location = new System.Drawing.Point(35, 95);
            this.DataSelected.Name = "DataSelected";
            this.DataSelected.Size = new System.Drawing.Size(93, 16);
            this.DataSelected.TabIndex = 2;
            this.DataSelected.Text = "Data Selected";
            // 
            // FileType
            // 
            this.FileType.Enabled = false;
            this.FileType.Location = new System.Drawing.Point(155, 92);
            this.FileType.Name = "FileType";
            this.FileType.ReadOnly = true;
            this.FileType.Size = new System.Drawing.Size(247, 22);
            this.FileType.TabIndex = 3;
            // 
            // FileName
            // 
            this.FileName.Enabled = false;
            this.FileName.Location = new System.Drawing.Point(155, 30);
            this.FileName.Name = "FileName";
            this.FileName.ReadOnly = true;
            this.FileName.Size = new System.Drawing.Size(247, 22);
            this.FileName.TabIndex = 4;
            // 
            // EditButton
            // 
            this.EditButton.Location = new System.Drawing.Point(364, 232);
            this.EditButton.Name = "EditButton";
            this.EditButton.Size = new System.Drawing.Size(75, 23);
            this.EditButton.TabIndex = 6;
            this.EditButton.Text = "Edit";
            this.EditButton.UseVisualStyleBackColor = true;
            this.EditButton.Click += new System.EventHandler(this.EditButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.EditButton);
            this.Controls.Add(this.FileName);
            this.Controls.Add(this.FileType);
            this.Controls.Add(this.DataSelected);
            this.Controls.Add(this.SendButton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button SendButton;
        private System.Windows.Forms.Label DataSelected;
        private System.Windows.Forms.TextBox FileType;
        private System.Windows.Forms.TextBox FileName;
        private System.Windows.Forms.Button EditButton;
    }
}

