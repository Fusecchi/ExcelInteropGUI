﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelInteropGUI
{
    public partial class SelectDataForm : Form
    {
        public System.Data.DataTable DatatoClick { get; set; } = new System.Data.DataTable();
        public Action<(int rowInd, int ColInd, object CellVal)> selectedData;
        public List<(int row, int col)> Highlight { get; set; } = new List<(int row, int col)>();
        public SelectDataForm()
        {
            InitializeComponent();
        }

        private void SelectDataForm_Load(object sender, EventArgs e)
        {
            int width = 0;
            CellToChose.DataSource = DatatoClick;
            CellToChose.Refresh();
            CellToChose.Invalidate();
            if(CellToChose.Columns.Count<0)
                CellToChose.AutoResizeColumn((int)DataGridViewAutoSizeColumnMode.AllCells);
            CellToChose.Dock = DockStyle.Fill;
            CellToChose.PerformLayout();
            if(Highlight.Count > 0 && Highlight.Count<CellToChose.RowCount) 
                CellToChose.Rows[Highlight[0].row-1].Cells[Highlight[0].col-1].Style.BackColor = Color.Yellow;
            foreach(DataGridViewColumn column in CellToChose.Columns)
            {
                width += column.Width;
            }
            if (width > 800) width = 800;
            this.ClientSize = new Size(width, CellToChose.Height);
        }

        private void CellToChose_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void SelectDataForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            CellToChose.DataSource = null;
            CellToChose.Rows.Clear();
            CellToChose.Columns.Clear();
        }

        private void CellToChose_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void CellToChose_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            selectedData?.Invoke((e.RowIndex, e.ColumnIndex, DatatoClick.Rows[e.RowIndex][e.ColumnIndex]));
            this.Close();
        }
    }
}
