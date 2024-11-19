﻿using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelInteropGUI
{
    public partial class EditWin : Form
    {
        public DataTable EditData { get; set; }
        public event Action<DataTable> DataSaved;
        public List<(int ChangedRow, int ChangedCol, object ChangedVal)> Log { get; set; } = new List<(int ChangedRow, int ChangedCol, object ChangedVal)> ();
        private bool changed;
        public EditWin()
        {
            InitializeComponent();
        }

        private void EditTable_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void EditWin_Load(object sender, EventArgs e)
        {
            EditTable.DataSource = EditData;
            EditTable.AutoResizeColumn((int)DataGridViewAutoSizeColumnMode.AllCells);
            EditTable.AutoResizeRow((int)DataGridViewAutoSizeRowMode.AllCells);
            EditTable.Dock = DockStyle.Fill;
            EditTable.PerformLayout();
            int newWidth = EditTable.Width+700;
            int newHeight = EditTable.Height;
            SaveEdit.Location = new Point(newWidth+20, SaveEdit.Location.Y);
            LogButton.Location = new Point(newWidth + 20, LogButton.Location.Y);
            CLoseButton.Location = new Point(newWidth + 20, CLoseButton.Location.Y);
            //Debug.WriteLine($"current pos : ${SaveEdit.Location}");
            //Debug.WriteLine($"current pos : ${CLoseButton.Location}");
            //Debug.WriteLine($"The Height: { newHeight} The Width:{newWidth}");

            this.ClientSize = new Size(SaveEdit.Location.X+130, EditTable.Size.Height);
            //Debug.WriteLine($"The Height: { this.ClientSize.Height} The Width:{this.ClientSize.Width}");
        }

        private void EditWin_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void SaveEdit_Click(object sender, EventArgs e)
        {
            changed = false;
            DataSaved?.Invoke(EditData);
            MessageBox.Show("Succesfully Saved");
        }

        private void EditTable_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            changed = true;
            if (e.ColumnIndex > 0 && e.RowIndex > 0) 
            {
                
            }
        }

        private void EditWin_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (changed) 
            {
                DialogResult ConfirmExit =  MessageBox.Show("Any Change won't be Saved",
                    "Unsaved Change",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);
                switch (ConfirmExit) { 
                    case DialogResult.Yes:
                        foreach(var RollBack in Log)
                        {
                            EditData.Rows[RollBack.ChangedRow][RollBack.ChangedCol] = RollBack.ChangedVal;
                        }
                        break;
                    case DialogResult.No:
                        SaveEdit_Click(sender, e);
                        break;
                    case DialogResult.Cancel:
                        e.Cancel = true;
                        break;
                    
                }
            }
        }

        private void EditTable_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Log.Add((e.RowIndex, e.ColumnIndex, EditData.Rows[e.RowIndex][e.ColumnIndex]));
        }

        private void CLoseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void LogButton_Click(object sender, EventArgs e)
        {

        }
    }
}
