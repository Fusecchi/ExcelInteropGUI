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
            Debug.WriteLine($"The Height: { newHeight} The Width:{newWidth}");

            this.ClientSize = new Size(newWidth, newHeight);

        }
    }
}
