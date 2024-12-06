using DocumentFormat.OpenXml.Bibliography;
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
        public event Action<DataTable, List<(int ChangedRow, int ChangedCol, object ChangedVal)>> DataSaved;
        public static List<(int ChangedRow, int ChangedCol, object ChangedVal)> LocalLog = new List<(int ChangedRow, int ChangedCol, object ChangedVal)>();
        public List<EventLogEntry> EventLog = new List<EventLogEntry>();
        private bool changed;
        public EditWin()
        {
            InitializeComponent();
            this.MaximizeBox = false;
        }

        private void EditTable_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void EditWin_Load(object sender, EventArgs e)
        {
            EditTable.DataSource = EditData;
            EditTable.AutoResizeColumn((int)DataGridViewAutoSizeColumnMode.AllCells);
            EditTable.Dock = DockStyle.Fill;
            EditTable.PerformLayout();
            EditTable.Refresh();
            EditTable.Invalidate();
            int newWidth = EditTable.Width+700;
            int newHeight = EditTable.Height;
            SaveEdit.Location = new Point(newWidth+20, SaveEdit.Location.Y);
            LogButton.Location = new Point(newWidth + 20, LogButton.Location.Y);
            CLoseButton.Location = new Point(newWidth + 20, CLoseButton.Location.Y);
            this.ClientSize = new Size(SaveEdit.Location.X+130, EditTable.Size.Height);

        }

        private void EditWin_FormClosed(object sender, FormClosedEventArgs e)
        {
            EditTable.Dispose();
        }

        private void SaveEdit_Click(object sender, EventArgs e)
        {
            changed = false;
            DataSaved?.Invoke(EditData,LocalLog);
            MessageBox.Show("Succesfully Saved");
        }

        private void EditTable_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            changed = true;
            SharedData.NewValues.Add((EditData.Rows[e.RowIndex][e.ColumnIndex], DateTime.Now));
        }

        private void EditWin_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (changed) 
            {
                DialogResult ConfirmExit =  MessageBox.Show("Are you sure you want to close?, any change won't be saved",
                    "Unsaved Change",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);
                switch (ConfirmExit) { 
                    case DialogResult.Yes:
                        foreach(var RollBack in LocalLog)
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
            LocalLog.Add((e.RowIndex, e.ColumnIndex, EditData.Rows[e.RowIndex][e.ColumnIndex]));
            SharedData.Log.Add((e.RowIndex, e.ColumnIndex, EditData.Rows[e.RowIndex][e.ColumnIndex]));
        }

        private void CLoseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void LogButton_Click(object sender, EventArgs e)
        {
            EditLog editLog = new EditLog();
            editLog.Show();
            editLog.FormClosed += (s, args) => this.Show();
            editLog.SelectedAction += _SelectedAction;
            this.Hide();
        }
        private void _SelectedAction(int RollBack)
        {
            foreach(var RB in SharedData.Log.AsEnumerable().Reverse())
            {
                EditData.Rows[RB.ChangedRow][RB.ChangedCol] = RB.ChangedVal;
            }
            SharedData.Log.RemoveRange(RollBack, SharedData.Log.Count);

        }
    }
}
