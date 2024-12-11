namespace ExcelInteropGUI
{
    partial class EditWin
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
            this.EditTable = new System.Windows.Forms.DataGridView();
            this.SaveEdit = new System.Windows.Forms.Button();
            this.CLoseButton = new System.Windows.Forms.Button();
            this.LogButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.EditTable)).BeginInit();
            this.SuspendLayout();
            // 
            // EditTable
            // 
            this.EditTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.EditTable.Location = new System.Drawing.Point(12, 12);
            this.EditTable.Name = "EditTable";
            this.EditTable.RowHeadersWidth = 51;
            this.EditTable.RowTemplate.Height = 24;
            this.EditTable.Size = new System.Drawing.Size(477, 230);
            this.EditTable.TabIndex = 6;
            this.EditTable.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.EditTable_CellBeginEdit);
            this.EditTable.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.EditTable_CellContentClick);
            this.EditTable.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.EditTable_CellValueChanged);
            // 
            // SaveEdit
            // 
            this.SaveEdit.Location = new System.Drawing.Point(12, 12);
            this.SaveEdit.Name = "SaveEdit";
            this.SaveEdit.Size = new System.Drawing.Size(150, 50);
            this.SaveEdit.TabIndex = 7;
            this.SaveEdit.Text = "Save";
            this.SaveEdit.UseVisualStyleBackColor = true;
            this.SaveEdit.Click += new System.EventHandler(this.SaveEdit_Click);
            // 
            // CLoseButton
            // 
            this.CLoseButton.Location = new System.Drawing.Point(12, 160);
            this.CLoseButton.Name = "CLoseButton";
            this.CLoseButton.Size = new System.Drawing.Size(150, 50);
            this.CLoseButton.TabIndex = 9;
            this.CLoseButton.Text = "Close";
            this.CLoseButton.UseVisualStyleBackColor = true;
            this.CLoseButton.Click += new System.EventHandler(this.CLoseButton_Click);
            // 
            // LogButton
            // 
            this.LogButton.Location = new System.Drawing.Point(12, 84);
            this.LogButton.Name = "LogButton";
            this.LogButton.Size = new System.Drawing.Size(150, 50);
            this.LogButton.TabIndex = 10;
            this.LogButton.Text = "Log";
            this.LogButton.UseVisualStyleBackColor = true;
            this.LogButton.Click += new System.EventHandler(this.LogButton_Click);
            // 
            // EditWin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.LogButton);
            this.Controls.Add(this.CLoseButton);
            this.Controls.Add(this.SaveEdit);
            this.Controls.Add(this.EditTable);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "EditWin";
            this.Text = "Edit Data";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.EditWin_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.EditWin_FormClosed);
            this.Load += new System.EventHandler(this.EditWin_Load);
            ((System.ComponentModel.ISupportInitialize)(this.EditTable)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView EditTable;
        private System.Windows.Forms.Button SaveEdit;
        private System.Windows.Forms.Button CLoseButton;
        private System.Windows.Forms.Button LogButton;
    }
}