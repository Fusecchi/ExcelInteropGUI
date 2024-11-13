using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.Command;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace ExcelInteropGUI
{
    public partial class Form1 : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        public Form1()
        {
            InitializeComponent();
        }

        private void SendButton_Click(object sender, EventArgs e)
        {
            try
            {
                ofd.Title = "Select Excel File";
                ofd.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv";
                if(ofd.ShowDialog() == DialogResult.OK)
                {
                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void LoadData()
        {
            string fp = ofd.FileName;
            string fn = Path.GetFileName(fp);
            //Debug.WriteLine(fp);
            if (!string.IsNullOrEmpty(fp))
            {
                FileName.Text = fn;
                using (var workbook = new XLWorkbook(fp))
                {
                    var sheet = workbook.Worksheet(1);
                    var lastRow = sheet.LastRowUsed().RowNumber();
                    var lastCol = sheet.LastColumnUsed().ColumnNumber();
                    var DataRange = sheet.Range(2, 1, lastRow, lastCol);
                    for (int i = 0; i < lastRow-1; i++) {
                        if (!sheet.Cell(2, 1).Value.Equals(sheet.Cell(2+i,1).Value))
                        {
                            MessageBox.Show($"Cell in: {sheet.Cell(2 + i, 1)} has value of: {sheet.Cell(2+i,1).Value}");
                            fp = null;
                            FileName.Text = null;
                            return;
                        }
                    }
                    FileType.Text = sheet.Cell(2,1).Value.ToString();
                    var CellAddr = new List<int>();
                    foreach (var cell in DataRange.Cells())
                    {
                        if (cell.Value.Equals(0) || cell.IsEmpty())
                        {
                            //Debug.WriteLine(cell.Address.RowNumber);
                            if (!CellAddr.Contains(cell.Address.RowNumber)){
                                CellAddr.Add(cell.Address.RowNumber);
                            }
                        }
                    }
                    Debug.WriteLine(CellAddr.Count);
                    System.Data.DataTable EditData = new System.Data.DataTable();
                    for(int MakeCol =1; MakeCol<=lastCol;MakeCol++)
                    {
                        EditData.Columns.Add(sheet.Cell(1, MakeCol).Value.ToString());
                    }
                    foreach (var addr in CellAddr)
                    {
                        DataRow row = EditData.NewRow();
                        for (int i = 0; i < lastCol; i++)
                        {
                            row[i] = sheet.Cell(addr, i+1).Value;
                        }
                        EditData.Rows.Add(row);
                    }
                    EditTable.DataSource = EditData;
                }
            }
        }
        private void reset()
        {

        }
        private void MachineName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
