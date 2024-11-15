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
using System.Xml;
using DocumentFormat.OpenXml.Office2016.Drawing.Command;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using static ClosedXML.Excel.XLPredefinedFormat;

namespace ExcelInteropGUI
{
    public partial class Form1 : Form
    {
        OpenFileDialog ofd = new OpenFileDialog();
        System.Data.DataTable EditData { get; set; } = new System.Data.DataTable();
        System.Data.DataTable TempTableToConvert = new System.Data.DataTable();
        int selectedSheet =1;
        IXLWorksheet From, To,sheet, ToSheet;
        XLWorkbook workbook, PasteBook;
        string ConvertFromCSV;
        bool Converted;
        List<int> CellAddr = new List<int>();
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
        private void CSVConvert()
        {
            using (StreamReader stream = new StreamReader(ofd.FileName, Encoding.GetEncoding("shift-jis"))) 
            { //please change ofd.filename
                string HeaderLine = stream.ReadLine();
                if (HeaderLine != null) 
                {
                    string[] headers = HeaderLine.Split(',');
                    foreach (string header in headers) 
                    {
                        TempTableToConvert.Columns.Add(header.Trim());
                    }
                    while (!stream.EndOfStream)
                    {
                        string[] values = stream.ReadLine().Split(',');
                        TempTableToConvert.Rows.Add(values);
                    }
                    
                }
                
            }
            using (var workbook = new XLWorkbook()) 
            {
                var worksheet = workbook.Worksheets.Add("sheet1");
                worksheet.Cell(1, 1).InsertTable(TempTableToConvert);
                string FilePath = Path.ChangeExtension(ofd.FileName, ".xlsx");
                workbook.SaveAs(FilePath);
                ConvertFromCSV = FilePath;
            }
            Converted = true;

        }

        private void LoadData()
        {
            string fp = ofd.FileName;
            string fn = Path.GetFileName(fp);
            //Debug.WriteLine(fp);
            if (!string.IsNullOrEmpty(fp))
            {
                if (fp.EndsWith(".csv"))
                {
                    CSVConvert();
                    fp = ConvertFromCSV;
                }
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
                DetectBlank(DataRange,lastCol);
            }
        }
        private void DetectBlank(IXLRange DataRange, int lastCol)
        {
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

        private void reset()
        {
            XmlDocument xmlDocument = new XmlDocument();

        }
        private void MachineName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void EditButton_Click(object sender, EventArgs e)
        {

                EditWin editwin = new EditWin();
                editwin.EditData = EditData;
                editwin.DataSaved += editwin_Datasaved;
                editwin.FormClosed += (s, args) => this.Show();
                editwin.Show();
                this.Hide();

        }
        private void editwin_Datasaved(System.Data.DataTable updatedTable)
        {
            EditData = updatedTable;
            List<int> EditedTable = new List<int>();
            foreach (DataRow row in EditData.Rows) 
            {
                foreach (DataColumn col in EditData.Columns) {
                    if (!EditedTable.Contains(EditData.Rows.IndexOf(row)))
                    {
                        //Debug.WriteLine($"Currently on Row: {EditData.Rows.IndexOf(row).ToString()} Col: {EditData.Columns.IndexOf(col)}, Add to List");
                        EditedTable.Add(EditData.Rows.IndexOf(row));
                        //Debug.WriteLine("Current List contents: " + string.Join(", ", EditedTable));
                    }
                    if (row[col] == DBNull.Value || string.IsNullOrWhiteSpace(row[col].ToString()))
                    {
                        //Debug.WriteLine($"Currently on Row: {EditData.Rows.IndexOf(row).ToString()} Col: {EditData.Columns.IndexOf(col)}, Skipped and Removed");
                        EditedTable.Remove(EditData.Rows.IndexOf(row));
                        //Debug.WriteLine("Current List contents: " + string.Join(", ", EditedTable));
                        break ;
                    }
                }
            }
            //Debug.WriteLine("List contents: " + string.Join(", ", EditedTable));
            //Debug.WriteLine("Address " + string.Join(", ", CellAddr));
            foreach (var index in EditedTable)
            {
                foreach(DataColumn colins in EditData.Columns) 
                {
                    //Debug.WriteLine(EditData.Rows[index][colins]?.ToString());
                    var x = CellAddr[index];
                    var y = EditData.Columns.IndexOf(colins)+1;
                    //Debug.WriteLine($"Value x = {x}, Value y = {y}");
                    From.Cell(x, y).Value = EditData.Rows[index][colins]?.ToString();
                }
            }

            EditedTable.Clear();
        }
        private void Editwin_FormClosed(object sender, FormClosedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void SelectSheet_Click(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Converted)
            {
                if (File.Exists(ConvertFromCSV))
                {
                    File.Delete(ConvertFromCSV);
                }
            }
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
            var LastRowTo = From.LastRowUsed().RowNumber();
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
