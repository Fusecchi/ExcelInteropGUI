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

        public EditLog()
        {
            InitializeComponent();
        }

        private void EditLog_Load(object sender, EventArgs e)
        {
            ShowLogText.Multiline = true;
            for (int i = 0; i<SharedData.Log.Count; i++ )
            {
                if (SharedData.NewValues[i].ChangedVal.ToString() == "" || SharedData.NewValues[i].ChangedVal.ToString() == "0")
                {
                    ShowLogText.SelectionColor = Color.Red;
                    ShowLogText.AppendText($"{SharedData.NewValues[i].EdittedTime} Deleted The Value in Row {SharedData.Log[i].ChangedRow} Collumn {SharedData.Log[i].ChangedCol} \n");
                }
                else
                {
                    ShowLogText.SelectionColor = Color.Black;
                    ShowLogText.AppendText($"{SharedData.NewValues[i].EdittedTime} Changed The Value in Row: ");
                    ShowLogText.SelectionColor = Color.Blue;
                    ShowLogText.AppendText($"{SharedData.Log[i].ChangedRow} ");
                    ShowLogText.SelectionColor = Color.Black;
                    ShowLogText.AppendText($"Collumn: ");
                    ShowLogText.SelectionColor = Color.Blue;
                    ShowLogText.AppendText($"{SharedData.Log[i].ChangedCol} ");
                    ShowLogText.SelectionColor = Color.Black;
                    ShowLogText.AppendText($"From: ");
                    ShowLogText.SelectionColor = Color.Blue;
                    ShowLogText.AppendText($"{SharedData.Log[i].ChangedVal} ");
                    ShowLogText.SelectionColor = Color.Black;
                    ShowLogText.AppendText($"To: ");
                    ShowLogText.SelectionColor = Color.Blue;
                    ShowLogText.AppendText($"{SharedData.NewValues[i].ChangedVal} \n");
                }

            }
        }

        private void CloseEditLog_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
