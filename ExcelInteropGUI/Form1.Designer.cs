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
            this.PresetLabel = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.FolderBtn = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
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
            this.EditButton.Location = new System.Drawing.Point(364, 339);
            this.EditButton.Name = "EditButton";
            this.EditButton.Size = new System.Drawing.Size(75, 23);
            this.EditButton.TabIndex = 6;
            this.EditButton.Text = "Edit";
            this.EditButton.UseVisualStyleBackColor = true;
            this.EditButton.Click += new System.EventHandler(this.EditButton_Click);
            // 
            // SendButton
            // 
            this.SendButton.Location = new System.Drawing.Point(59, 339);
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
            this.ResetButton.Location = new System.Drawing.Point(638, 339);
            this.ResetButton.Name = "ResetButton";
            this.ResetButton.Size = new System.Drawing.Size(78, 23);
            this.ResetButton.TabIndex = 12;
            this.ResetButton.Text = "Reset";
            this.ResetButton.UseVisualStyleBackColor = true;
            this.ResetButton.Click += new System.EventHandler(this.ResetButton_Click);
            // 
            // PresetLabel
            // 
            this.PresetLabel.AutoSize = true;
            this.PresetLabel.Location = new System.Drawing.Point(48, 156);
            this.PresetLabel.Name = "PresetLabel";
            this.PresetLabel.Size = new System.Drawing.Size(46, 16);
            this.PresetLabel.TabIndex = 13;
            this.PresetLabel.Text = "Preset";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(447, 156);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 16);
            this.label2.TabIndex = 14;
            this.label2.Text = "label2";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(42, 203);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(78, 29);
            this.button1.TabIndex = 16;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // FolderBtn
            // 
            this.FolderBtn.BackColor = System.Drawing.SystemColors.Control;
            this.FolderBtn.Image = global::ExcelInteropGUI.Properties.Resources.Folder_Icon_Template_Design_Vector_Graphics_13725642_1__Custom_1;
            this.FolderBtn.Location = new System.Drawing.Point(736, 26);
            this.FolderBtn.Name = "FolderBtn";
            this.FolderBtn.Size = new System.Drawing.Size(38, 31);
            this.FolderBtn.TabIndex = 15;
            this.FolderBtn.UseVisualStyleBackColor = false;
            this.FolderBtn.Click += new System.EventHandler(this.FolderBtn_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(211, 207);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(90, 24);
            this.button2.TabIndex = 17;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(364, 208);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(96, 24);
            this.button3.TabIndex = 18;
            this.button3.Text = "button3";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(786, 400);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.FolderBtn);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.PresetLabel);
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
        private System.Windows.Forms.Label PresetLabel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button FolderBtn;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
    }
}

