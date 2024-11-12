using Microsoft.Office.Interop.Excel;
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
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelInteropGUI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void SendButton_Click(object sender, EventArgs e)
        {
            try
            {
                Application Excelapp = new Application();
                Excelapp.Visible = true;
                Workbook WB = Excelapp.Workbooks.Add();
                Worksheet WS = (Worksheet)WB.Sheets[1];

                WS.Cells[1,1].Value = "Clicked";
                Debug.WriteLine("Clicked");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
