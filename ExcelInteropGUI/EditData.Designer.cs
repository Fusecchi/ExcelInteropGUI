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
            this.EditTable.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.EditTable_CellContentClick);
            // 
            // EditWin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.EditTable);
            this.Name = "EditWin";
            this.Text = "Edit Data";
            this.Load += new System.EventHandler(this.EditWin_Load);
            ((System.ComponentModel.ISupportInitialize)(this.EditTable)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView EditTable;
    }
}