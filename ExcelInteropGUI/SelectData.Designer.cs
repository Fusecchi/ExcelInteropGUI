﻿namespace ExcelInteropGUI
{
    partial class SelectDataForm
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
            this.CellToChose = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.CellToChose)).BeginInit();
            this.SuspendLayout();
            // 
            // CellToChose
            // 
            this.CellToChose.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.CellToChose.Location = new System.Drawing.Point(5, 14);
            this.CellToChose.Name = "CellToChose";
            this.CellToChose.RowHeadersWidth = 51;
            this.CellToChose.RowTemplate.Height = 24;
            this.CellToChose.Size = new System.Drawing.Size(115, 89);
            this.CellToChose.TabIndex = 0;
            this.CellToChose.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.CellToChose_CellClick);
            this.CellToChose.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.CellToChose_CellDoubleClick);
            this.CellToChose.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.CellToChose_CellMouseDoubleClick);
            // 
            // SelectDataForm
            // 
            this.ClientSize = new System.Drawing.Size(513, 371);
            this.Controls.Add(this.CellToChose);
            this.Name = "SelectDataForm";
            this.Text = "Select Data";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SelectDataForm_FormClosed);
            this.Load += new System.EventHandler(this.SelectDataForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.CellToChose)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView CellToChose;
    }
}