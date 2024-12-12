namespace ExcelInteropGUI
{
    partial class Menu
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Menu));
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
            this.CSVNew_Save = new System.Windows.Forms.TextBox();
            this.SelectPreset = new System.Windows.Forms.ComboBox();
            this.MakePreset = new System.Windows.Forms.Button();
            this.FolderBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.CopyName = new System.Windows.Forms.Label();
            this.New_Sheet = new System.Windows.Forms.TextBox();
            this.Add_NewSheet = new System.Windows.Forms.Button();
            this.EditPreset = new System.Windows.Forms.Button();
            this.DeleteBtn = new System.Windows.Forms.Button();
            this.NewBookSave = new System.Windows.Forms.TextBox();
            this.English_RB = new System.Windows.Forms.RadioButton();
            this.Japanese_RB = new System.Windows.Forms.RadioButton();
            this.Language_Label = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // SelectData
            // 
            this.SelectData.AutoSize = true;
            this.SelectData.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.SelectData.Location = new System.Drawing.Point(39, 58);
            this.SelectData.Name = "SelectData";
            this.SelectData.Size = new System.Drawing.Size(80, 26);
            this.SelectData.TabIndex = 0;
            this.SelectData.Text = "Select File";
            this.SelectData.UseVisualStyleBackColor = true;
            this.SelectData.Click += new System.EventHandler(this.SelectData_Click);
            // 
            // DataSelected
            // 
            this.DataSelected.AutoSize = true;
            this.DataSelected.Location = new System.Drawing.Point(36, 104);
            this.DataSelected.Name = "DataSelected";
            this.DataSelected.Size = new System.Drawing.Size(93, 16);
            this.DataSelected.TabIndex = 2;
            this.DataSelected.Text = "Data Selected";
            // 
            // FileType
            // 
            this.FileType.Enabled = false;
            this.FileType.Location = new System.Drawing.Point(156, 98);
            this.FileType.Name = "FileType";
            this.FileType.ReadOnly = true;
            this.FileType.Size = new System.Drawing.Size(196, 22);
            this.FileType.TabIndex = 3;
            this.FileType.TextChanged += new System.EventHandler(this.FileType_TextChanged);
            // 
            // FileName
            // 
            this.FileName.Enabled = false;
            this.FileName.Location = new System.Drawing.Point(156, 58);
            this.FileName.Name = "FileName";
            this.FileName.ReadOnly = true;
            this.FileName.Size = new System.Drawing.Size(196, 22);
            this.FileName.TabIndex = 4;
            // 
            // EditButton
            // 
            this.EditButton.Location = new System.Drawing.Point(361, 411);
            this.EditButton.Name = "EditButton";
            this.EditButton.Size = new System.Drawing.Size(75, 31);
            this.EditButton.TabIndex = 6;
            this.EditButton.Text = "Edit";
            this.EditButton.UseVisualStyleBackColor = true;
            this.EditButton.Click += new System.EventHandler(this.EditButton_Click);
            // 
            // SendButton
            // 
            this.SendButton.Location = new System.Drawing.Point(56, 411);
            this.SendButton.Name = "SendButton";
            this.SendButton.Size = new System.Drawing.Size(69, 31);
            this.SendButton.TabIndex = 7;
            this.SendButton.Text = "Send";
            this.SendButton.UseVisualStyleBackColor = true;
            this.SendButton.Click += new System.EventHandler(this.SendButton_Click);
            // 
            // SelectTarget
            // 
            this.SelectTarget.AutoSize = true;
            this.SelectTarget.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.SelectTarget.Location = new System.Drawing.Point(398, 62);
            this.SelectTarget.Name = "SelectTarget";
            this.SelectTarget.Size = new System.Drawing.Size(95, 26);
            this.SelectTarget.TabIndex = 8;
            this.SelectTarget.Text = "SelectTarget";
            this.SelectTarget.UseVisualStyleBackColor = true;
            this.SelectTarget.Click += new System.EventHandler(this.SelectTarget_Click);
            // 
            // SelectSheet
            // 
            this.SelectSheet.AutoSize = true;
            this.SelectSheet.Location = new System.Drawing.Point(398, 104);
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
            this.TargetSheet.Location = new System.Drawing.Point(517, 104);
            this.TargetSheet.Name = "TargetSheet";
            this.TargetSheet.Size = new System.Drawing.Size(196, 24);
            this.TargetSheet.TabIndex = 10;
            this.TargetSheet.SelectedIndexChanged += new System.EventHandler(this.TargetSheet_SelectedIndexChanged);
            this.TargetSheet.Click += new System.EventHandler(this.TargetSheet_Click);
            // 
            // TargetName
            // 
            this.TargetName.Enabled = false;
            this.TargetName.Location = new System.Drawing.Point(517, 62);
            this.TargetName.Name = "TargetName";
            this.TargetName.ReadOnly = true;
            this.TargetName.Size = new System.Drawing.Size(196, 22);
            this.TargetName.TabIndex = 11;
            // 
            // ResetButton
            // 
            this.ResetButton.Location = new System.Drawing.Point(635, 411);
            this.ResetButton.Name = "ResetButton";
            this.ResetButton.Size = new System.Drawing.Size(78, 31);
            this.ResetButton.TabIndex = 12;
            this.ResetButton.Text = "Reset";
            this.ResetButton.UseVisualStyleBackColor = true;
            this.ResetButton.Click += new System.EventHandler(this.ResetButton_Click);
            // 
            // PresetLabel
            // 
            this.PresetLabel.AutoSize = true;
            this.PresetLabel.Location = new System.Drawing.Point(36, 175);
            this.PresetLabel.Name = "PresetLabel";
            this.PresetLabel.Size = new System.Drawing.Size(46, 16);
            this.PresetLabel.TabIndex = 13;
            this.PresetLabel.Text = "Preset";
            this.PresetLabel.Click += new System.EventHandler(this.PresetLabel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(36, 139);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 16);
            this.label2.TabIndex = 24;
            this.label2.Text = "Saved As";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // CSVNew_Save
            // 
            this.CSVNew_Save.Location = new System.Drawing.Point(156, 139);
            this.CSVNew_Save.Name = "CSVNew_Save";
            this.CSVNew_Save.Size = new System.Drawing.Size(196, 22);
            this.CSVNew_Save.TabIndex = 25;
            // 
            // SelectPreset
            // 
            this.SelectPreset.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.SelectPreset.FormattingEnabled = true;
            this.SelectPreset.Location = new System.Drawing.Point(156, 175);
            this.SelectPreset.Name = "SelectPreset";
            this.SelectPreset.Size = new System.Drawing.Size(196, 24);
            this.SelectPreset.TabIndex = 19;
            this.SelectPreset.SelectedIndexChanged += new System.EventHandler(this.SelectPreset_SelectedIndexChanged);
            this.SelectPreset.Click += new System.EventHandler(this.SelectPreset_Click);
            this.SelectPreset.Leave += new System.EventHandler(this.SelectPreset_Leave);
            this.SelectPreset.Validating += new System.ComponentModel.CancelEventHandler(this.SelectPreset_Validating);
            // 
            // MakePreset
            // 
            this.MakePreset.AutoSize = true;
            this.MakePreset.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.MakePreset.Location = new System.Drawing.Point(20, 281);
            this.MakePreset.Name = "MakePreset";
            this.MakePreset.Size = new System.Drawing.Size(93, 26);
            this.MakePreset.TabIndex = 16;
            this.MakePreset.Text = "Make Preset";
            this.MakePreset.UseVisualStyleBackColor = true;
            this.MakePreset.Click += new System.EventHandler(this.makePreset_Click);
            // 
            // FolderBtn
            // 
            this.FolderBtn.BackColor = System.Drawing.SystemColors.Control;
            this.FolderBtn.Image = global::ExcelInteropGUI.Properties.Resources.Folder_Icon_Template_Design_Vector_Graphics_13725642_1__Custom_1;
            this.FolderBtn.Location = new System.Drawing.Point(675, 208);
            this.FolderBtn.Name = "FolderBtn";
            this.FolderBtn.Size = new System.Drawing.Size(38, 31);
            this.FolderBtn.TabIndex = 15;
            this.FolderBtn.UseVisualStyleBackColor = false;
            this.FolderBtn.Click += new System.EventHandler(this.FolderBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(398, 212);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 16);
            this.label1.TabIndex = 22;
            this.label1.Text = "Saved As";
            // 
            // CopyName
            // 
            this.CopyName.AutoSize = true;
            this.CopyName.Location = new System.Drawing.Point(398, 151);
            this.CopyName.Name = "CopyName";
            this.CopyName.Size = new System.Drawing.Size(77, 16);
            this.CopyName.TabIndex = 29;
            this.CopyName.Text = "Copy Sheet";
            this.CopyName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // New_Sheet
            // 
            this.New_Sheet.Location = new System.Drawing.Point(517, 151);
            this.New_Sheet.Name = "New_Sheet";
            this.New_Sheet.Size = new System.Drawing.Size(113, 22);
            this.New_Sheet.TabIndex = 21;
            // 
            // Add_NewSheet
            // 
            this.Add_NewSheet.Location = new System.Drawing.Point(649, 145);
            this.Add_NewSheet.Name = "Add_NewSheet";
            this.Add_NewSheet.Size = new System.Drawing.Size(64, 46);
            this.Add_NewSheet.TabIndex = 20;
            this.Add_NewSheet.Text = "Add Sheet";
            this.Add_NewSheet.UseVisualStyleBackColor = true;
            this.Add_NewSheet.Click += new System.EventHandler(this.Add_NewSheet_Click);
            // 
            // EditPreset
            // 
            this.EditPreset.AutoSize = true;
            this.EditPreset.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.EditPreset.Location = new System.Drawing.Point(119, 281);
            this.EditPreset.Name = "EditPreset";
            this.EditPreset.Size = new System.Drawing.Size(82, 26);
            this.EditPreset.TabIndex = 17;
            this.EditPreset.Text = "Edit Preset";
            this.EditPreset.UseVisualStyleBackColor = true;
            this.EditPreset.Click += new System.EventHandler(this.EditPreset_Click);
            // 
            // DeleteBtn
            // 
            this.DeleteBtn.AutoSize = true;
            this.DeleteBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.DeleteBtn.Location = new System.Drawing.Point(201, 281);
            this.DeleteBtn.Name = "DeleteBtn";
            this.DeleteBtn.Size = new System.Drawing.Size(99, 26);
            this.DeleteBtn.TabIndex = 18;
            this.DeleteBtn.Text = "Delete Preset";
            this.DeleteBtn.UseVisualStyleBackColor = true;
            this.DeleteBtn.Click += new System.EventHandler(this.Delete_Click);
            // 
            // NewBookSave
            // 
            this.NewBookSave.Location = new System.Drawing.Point(517, 212);
            this.NewBookSave.Name = "NewBookSave";
            this.NewBookSave.Size = new System.Drawing.Size(152, 22);
            this.NewBookSave.TabIndex = 23;
            // 
            // English_RB
            // 
            this.English_RB.AutoSize = true;
            this.English_RB.Location = new System.Drawing.Point(489, 12);
            this.English_RB.Name = "English_RB";
            this.English_RB.Size = new System.Drawing.Size(72, 20);
            this.English_RB.TabIndex = 26;
            this.English_RB.TabStop = true;
            this.English_RB.Text = "English";
            this.English_RB.UseVisualStyleBackColor = true;
            this.English_RB.CheckedChanged += new System.EventHandler(this.English_RB_CheckedChanged);
            // 
            // Japanese_RB
            // 
            this.Japanese_RB.AutoSize = true;
            this.Japanese_RB.Location = new System.Drawing.Point(585, 12);
            this.Japanese_RB.Name = "Japanese_RB";
            this.Japanese_RB.Size = new System.Drawing.Size(73, 20);
            this.Japanese_RB.TabIndex = 27;
            this.Japanese_RB.TabStop = true;
            this.Japanese_RB.Text = "日本語";
            this.Japanese_RB.UseVisualStyleBackColor = true;
            this.Japanese_RB.CheckedChanged += new System.EventHandler(this.Japanese_RB_CheckedChanged);
            // 
            // Language_Label
            // 
            this.Language_Label.AutoSize = true;
            this.Language_Label.Location = new System.Drawing.Point(395, 14);
            this.Language_Label.Name = "Language_Label";
            this.Language_Label.Size = new System.Drawing.Size(68, 16);
            this.Language_Label.TabIndex = 28;
            this.Language_Label.Text = "Language";
            // 
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(797, 469);
            this.Controls.Add(this.CopyName);
            this.Controls.Add(this.Language_Label);
            this.Controls.Add(this.Japanese_RB);
            this.Controls.Add(this.English_RB);
            this.Controls.Add(this.CSVNew_Save);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.NewBookSave);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.New_Sheet);
            this.Controls.Add(this.Add_NewSheet);
            this.Controls.Add(this.SelectPreset);
            this.Controls.Add(this.DeleteBtn);
            this.Controls.Add(this.EditPreset);
            this.Controls.Add(this.MakePreset);
            this.Controls.Add(this.FolderBtn);
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
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Menu";
            this.Text = "Menu";
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
        private System.Windows.Forms.Button FolderBtn;
        private System.Windows.Forms.Button MakePreset;
        private System.Windows.Forms.Button EditPreset;
        private System.Windows.Forms.Button DeleteBtn;
        private System.Windows.Forms.ComboBox SelectPreset;
        private System.Windows.Forms.Button Add_NewSheet;
        private System.Windows.Forms.TextBox New_Sheet;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox NewBookSave;
        private System.Windows.Forms.TextBox CSVNew_Save;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton English_RB;
        private System.Windows.Forms.RadioButton Japanese_RB;
        private System.Windows.Forms.Label Language_Label;
        private System.Windows.Forms.Label CopyName;
    }
}

