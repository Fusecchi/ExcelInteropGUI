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
            this.SelectData = new System.Windows.Forms.Button();
            this.DataSelected = new System.Windows.Forms.Label();
            this.FileType = new System.Windows.Forms.TextBox();
            this.FileName = new System.Windows.Forms.TextBox();
            this.EditButton = new System.Windows.Forms.Button();
            this.SendButton = new System.Windows.Forms.Button();
            this.SelectTarget = new System.Windows.Forms.Button();
            this.SelectSheet = new System.Windows.Forms.Label();
            this.TargetSheet = new System.Windows.Forms.ComboBox();
            this.TargetName = new System.Windows.Forms.TextBox();
            this.ResetButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // SelectData
            // 
            this.SelectData.Location = new System.Drawing.Point(38, 26);
            this.SelectData.Name = "SelectData";
            this.SelectData.Size = new System.Drawing.Size(83, 31);
            this.SelectData.TabIndex = 0;
            this.SelectData.Text = "Select File";
            this.SelectData.UseVisualStyleBackColor = true;
            this.SelectData.Click += new System.EventHandler(this.SelectData_Click);
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
            this.FileType.Size = new System.Drawing.Size(196, 22);
            this.FileType.TabIndex = 3;
            // 
            // FileName
            // 
            this.FileName.Enabled = false;
            this.FileName.Location = new System.Drawing.Point(155, 30);
            this.FileName.Name = "FileName";
            this.FileName.ReadOnly = true;
            this.FileName.Size = new System.Drawing.Size(196, 22);
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
            // SendButton
            // 
            this.SendButton.Location = new System.Drawing.Point(122, 232);
            this.SendButton.Name = "SendButton";
            this.SendButton.Size = new System.Drawing.Size(69, 23);
            this.SendButton.TabIndex = 7;
            this.SendButton.Text = "Send";
            this.SendButton.UseVisualStyleBackColor = true;
            this.SendButton.Click += new System.EventHandler(this.SendButton_Click);
            // 
            // SelectTarget
            // 
            this.SelectTarget.Location = new System.Drawing.Point(397, 26);
            this.SelectTarget.Name = "SelectTarget";
            this.SelectTarget.Size = new System.Drawing.Size(104, 31);
            this.SelectTarget.TabIndex = 8;
            this.SelectTarget.Text = "SelectTarget";
            this.SelectTarget.UseVisualStyleBackColor = true;
            this.SelectTarget.Click += new System.EventHandler(this.SelectTarget_Click);
            // 
            // SelectSheet
            // 
            this.SelectSheet.AutoSize = true;
            this.SelectSheet.Location = new System.Drawing.Point(408, 95);
            this.SelectSheet.Name = "SelectSheet";
            this.SelectSheet.Size = new System.Drawing.Size(83, 16);
            this.SelectSheet.TabIndex = 9;
            this.SelectSheet.Text = "Select Sheet";
            this.SelectSheet.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // TargetSheet
            // 
            this.TargetSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.TargetSheet.FormattingEnabled = true;
            this.TargetSheet.Location = new System.Drawing.Point(520, 92);
            this.TargetSheet.Name = "TargetSheet";
            this.TargetSheet.Size = new System.Drawing.Size(196, 24);
            this.TargetSheet.TabIndex = 10;
            this.TargetSheet.SelectedIndexChanged += new System.EventHandler(this.TargetSheet_SelectedIndexChanged);
            // 
            // TargetName
            // 
            this.TargetName.Enabled = false;
            this.TargetName.Location = new System.Drawing.Point(520, 30);
            this.TargetName.Name = "TargetName";
            this.TargetName.ReadOnly = true;
            this.TargetName.Size = new System.Drawing.Size(196, 22);
            this.TargetName.TabIndex = 11;
            // 
            // ResetButton
            // 
            this.ResetButton.Location = new System.Drawing.Point(598, 232);
            this.ResetButton.Name = "ResetButton";
            this.ResetButton.Size = new System.Drawing.Size(78, 23);
            this.ResetButton.TabIndex = 12;
            this.ResetButton.Text = "Reset";
            this.ResetButton.UseVisualStyleBackColor = true;
            this.ResetButton.Click += new System.EventHandler(this.ResetButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.ResetButton);
            this.Controls.Add(this.TargetName);
            this.Controls.Add(this.TargetSheet);
            this.Controls.Add(this.SelectSheet);
            this.Controls.Add(this.SelectTarget);
            this.Controls.Add(this.SendButton);
            this.Controls.Add(this.EditButton);
            this.Controls.Add(this.FileName);
            this.Controls.Add(this.FileType);
            this.Controls.Add(this.DataSelected);
            this.Controls.Add(this.SelectData);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button SelectData;
        private System.Windows.Forms.Label DataSelected;
        private System.Windows.Forms.TextBox FileType;
        private System.Windows.Forms.TextBox FileName;
        private System.Windows.Forms.Button EditButton;
        private System.Windows.Forms.Button SendButton;
        private System.Windows.Forms.Button SelectTarget;
        private System.Windows.Forms.Label SelectSheet;
        private System.Windows.Forms.ComboBox TargetSheet;
        private System.Windows.Forms.TextBox TargetName;
        private System.Windows.Forms.Button ResetButton;
    }
}

