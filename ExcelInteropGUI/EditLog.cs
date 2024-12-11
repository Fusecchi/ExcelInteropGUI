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
    public partial class EditLog : Form
    {
        private int SelectedRB;
        string deleted_val;
        string changed_val;
        public bool Japanese;
        public Action<int> SelectedAction;
        public EditLog()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.ResizeRedraw = false;
        }

        private void EditLog_Load(object sender, EventArgs e)
        {
            if (Japanese)
            {
                ChangeLanguage.change("ja-JP", this);

            }else
                ChangeLanguage.change("en-EN", this);

            for (int i = 0; i<SharedData.Log.Count; i++ )
            {
                if (Japanese)
                {
                     deleted_val = $"{SharedData.NewValues[i].EdittedTime} に列 {SharedData.Log[i].ChangedRow} 行 {SharedData.Log[i].ChangedCol} が解除されました \n";
                     changed_val = $"{SharedData.NewValues[i].EdittedTime} 行: " +
                            $"{SharedData.Log[i].ChangedRow} 列:{SharedData.Log[i].ChangedCol}  {SharedData.Log[i].ChangedVal}" +
                            $"から: {SharedData.NewValues[i].ChangedVal} になりました";
                }
                else
                {
                     deleted_val = $"{SharedData.NewValues[i].EdittedTime} Deleted The Value in Row {SharedData.Log[i].ChangedRow} Collumn {SharedData.Log[i].ChangedCol} \n";
                    changed_val = $"{SharedData.NewValues[i].EdittedTime} Changed The Value in Row: " +
       $"{SharedData.Log[i].ChangedRow} Collumn:{SharedData.Log[i].ChangedCol} From: {SharedData.Log[i].ChangedVal}" +
       $"To: {SharedData.NewValues[i].ChangedVal}";
                }
                if (SharedData.NewValues[i].ChangedVal.ToString() == "" || SharedData.NewValues[i].ChangedVal.ToString() == "0")
                {
                    listBox1.Items.Add(deleted_val);
                }
                else
                {
                    listBox1.Items.Add(changed_val);
                }

            }
        }

        private void CloseEditLog_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left) {
                SelectedRB = listBox1.SelectedIndex;
            }
        }

        private void Rollback_Click(object sender, EventArgs e)
        {
            SelectedAction?.Invoke(SelectedRB);
            this.Close();
        }
    }
}
