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
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelInteropGUI
{
    public partial class Form1 : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        System.Data.DataTable EditData = new System.Data.DataTable();
        int selectedSheet =1;
        IXLWorksheet From, To,sheet, ToSheet;
        XLWorkbook workbook, PasteBook;

        public Form1()
        {
            InitializeComponent();
        }

        private void SelectData_Click(object sender, EventArgs e)
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
                workbook = new XLWorkbook(fp);
                workbook.CalculateMode = XLCalculateMode.Manual;
                sheet = workbook.Worksheet(1);
                From = sheet;
                var lastRow = sheet.LastRowUsed().RowNumber();
                var lastCol = sheet.LastColumnUsed().ColumnNumber();
                var DataRange = sheet.Range(2, 1, lastRow, lastCol);
                for (int i = 0; i < lastRow - 1; i++)
                {
                    if (!sheet.Cell(2, 1).Value.Equals(sheet.Cell(2 + i, 1).Value))
                    {
                        MessageBox.Show($"Cell in: {sheet.Cell(2 + i, 1)} has value of: {sheet.Cell(2 + i, 1).Value}");
                        fp = null;
                        FileName.Text = null;
                        return;
                    }
                }
                FileType.Text = sheet.Cell(2, 1).Value.ToString();
                var CellAddr = new List<int>();
                foreach (var cell in DataRange.Cells())
                {
                    if (cell.Value.Equals(0) || cell.IsEmpty())
                    {
                        //Debug.WriteLine(cell.Address.RowNumber);
                        if (!CellAddr.Contains(cell.Address.RowNumber))
                        {
                            CellAddr.Add(cell.Address.RowNumber);
                        }
                    }
                }
                //Debug.WriteLine(CellAddr.Count);
                for (int MakeCol = 1; MakeCol <= lastCol; MakeCol++)
                {
                    EditData.Columns.Add(sheet.Cell(1, MakeCol).Value.ToString());
                }
                foreach (var addr in CellAddr)
                {
                    DataRow row = EditData.NewRow();
                    for (int i = 0; i < lastCol; i++)
                    {
                        row[i] = sheet.Cell(addr, i + 1).Value;
                    }
                    EditData.Rows.Add(row);
                }



            }
        }
        private void reset()
        {

        }
        private void MachineName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void EditButton_Click(object sender, EventArgs e)
        {
            EditWin editwin = new EditWin();
            editwin.EditData = EditData;
            editwin.Show();
        }

        private void SelectSheet_Click(object sender, EventArgs e)
        {

        }

        private void SelectTarget_Click(object sender, EventArgs e)
        {
            ofd.Title = "Select Excel File";
            ofd.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                TargetFile();
            }
        }

        private void TargetFile()
        {
            string Tp = ofd.FileName;
            string Tn = Path.GetFileName(Tp);
            if (!string.IsNullOrEmpty(Tp)) 
            {
                PasteBook = new XLWorkbook(Tp);
                PasteBook.CalculateMode = XLCalculateMode.Manual;
                TargetName.Text = Tn;
                foreach(var Sheet in PasteBook.Worksheets)
                {
                    TargetSheet.Items.Add(Sheet);
                }
                TargetSheet.SelectedIndex = 0;

                
            }
        }

        private void TargetSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedSheet = TargetSheet.SelectedIndex;
            ToSheet = PasteBook.Worksheet(selectedSheet+1);
            To = ToSheet;
            //MessageBox.Show($"Selected {selectedSheet}");
        }

        private void SendButton_Click(object sender, EventArgs e)
        {
            //MessageBox.Show($"Sending To {To} Chosed: {ToSheet}");
            var LastRowTo = To.LastRowUsed().RowNumber();
            for(var row = 0; row< LastRowTo; row++)
            {
                To.Cell(17 + row, 3).Value = From.Cell(2+row, 4).Value; //板厚
                To.Cell(17 + row, 7).Value = From.Cell(2 + row, 5).Value; //寸法W
                To.Cell(17 + row, 8).Value = From.Cell(2 + row, 6).Value; //寸法H
                To.Cell(17 + row, 14).Value = From.Cell(2 + row, 7).Value; //歩留
                string CleanTimeData = From.Cell(2 + row, 9).Value.ToString();
                var CharToRemove = "()（）分";
                foreach(char c in CharToRemove)
                {
                    CleanTimeData = CleanTimeData.Replace(c.ToString(), "");
                }
                To.Cell(17 + row, 12).Value = CleanTimeData; //三菱加工時間
            }
            PasteBook.Save();
            MessageBox.Show($"Sending {To} to {From}Complete");
        }
    }
}
