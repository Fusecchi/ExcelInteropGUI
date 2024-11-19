using System;
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
    public partial class EditLog : Form
    {
        public List<(int ChangedRow, int ChangedCol, object ChangedVal)> Log { get; set; } = new List<(int ChangedRow, int ChangedCol, object ChangedVal)>();
        public List<object> NewValues { get; set; } = new List<object>();

        public EditLog()
        {
            InitializeComponent();
        }

        private void EditLog_Load(object sender, EventArgs e)
        {
            ShowLogText.Multiline = true;
            for (int i = 0; i<Log.Count; i++ )
            {
                ShowLogText.AppendText($"Changed The Value in Row {Log[i].ChangedRow} Collumn {Log[i].ChangedCol} From: {Log[i].ChangedVal} To: {NewValues[i]} \n");
            }
        }
    }
}
