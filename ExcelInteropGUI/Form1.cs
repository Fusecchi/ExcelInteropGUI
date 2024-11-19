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
        //Declare FIle Dialog
        OpenFileDialog ofd = new OpenFileDialog();
        //Data Table to Edit the File Directly
        System.Data.DataTable EditData { get; set; } = new System.Data.DataTable();
        //Temporary Table to Transfer the data from CSV
        System.Data.DataTable TempTableToConvert = new System.Data.DataTable();
        int selectedSheet =1;
        IXLWorksheet From, To,sheet, ToSheet;
        XLWorkbook workbook, PasteBook;
        //Path of the Converted CSV
        string ConvertFromCSV;
        //CSV convert flag
        bool Converted;
        //Cell address of the Blank FIle
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
                return;
            }

}

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadData();
            
        }
        private void CSVConvert()
        {
            //Open the CSV File
            using (StreamReader stream = new StreamReader(ofd.FileName, Encoding.GetEncoding("shift-jis"))) 
            { 
                //Read the all the line
                string HeaderLine = stream.ReadLine();
                if (HeaderLine != null) 
                {
                    //Split all the line that are separated bt comma
                    string[] headers = HeaderLine.Split(',');
                    foreach (string header in headers) 
                    {
                        //Clean the white space and add it into the header
                        TempTableToConvert.Columns.Add(header.Trim());
                    }
                    while (!stream.EndOfStream)
                    {
                        //Read the next line and loop it unti it Find Blank
                        string[] values = stream.ReadLine().Split(',');
                        TempTableToConvert.Rows.Add(values);
                    }
                    
                }
                
            }
            using (var workbook = new XLWorkbook()) 
            {
                //Make the new xlsx Book
                var worksheet = workbook.Worksheets.Add("sheet1");
                //Paste the value in A1
                worksheet.Cell(1, 1).InsertTable(TempTableToConvert);
                //Add xlsx Extension
                string FilePath = Path.ChangeExtension(ofd.FileName, ".xlsx");
                //Save the workbook
                workbook.SaveAs(FilePath);
                ConvertFromCSV = FilePath;
            }
            Converted = true;

        }

        private void LoadData()
        {
            ResetData();
            //The File Path of the selected file
            string fp = ofd.FileName;
            //The File Name of the file
            string fn = Path.GetFileName(fp);
            //Debug.WriteLine(fp);
            if (!string.IsNullOrEmpty(fp))
            {
                //Check if the file is CSV extension
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
                //Check the validity off all the data for contamination
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
                if (CellAddr.Count == 0)
                {
                    EditButton.Enabled = false;
                }
                else
                {
                    EditButton.Enabled = true;
                }
            }
        }
        //Detect the blank in the file 
        private void DetectBlank(IXLRange DataRange, int lastCol)
        {
            foreach (var cell in DataRange.Cells())
            {
                if (cell.Value.Equals(0) || cell.IsEmpty())
                {
                    //Save the Row of the EMpty Cell
                    //Debug.WriteLine(cell.Address.RowNumber);
                    if (!CellAddr.Contains(cell.Address.RowNumber))
                    {
                        CellAddr.Add(cell.Address.RowNumber);
                    }
                }
            }
            //Debug.WriteLine(CellAddr.Count);
            //Make a table for edit 
            for (int MakeCol = 1; MakeCol <= lastCol; MakeCol++)
            {
                //Make the collumn according to the file selected
                EditData.Columns.Add(sheet.Cell(1, MakeCol).Value.ToString());
            }
            foreach (var addr in CellAddr)
            {
                //Clone the row of the blank cell
                DataRow row = EditData.NewRow();
                for (int i = 0; i < lastCol; i++)
                {
                    //Save the row in a Array and add it in bulk
                    row[i] = sheet.Cell(addr, i + 1).Value;
                }
                EditData.Rows.Add(row);
            }
        }


        private void MachineName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //Executed if the user click Edit Button
        private void EditButton_Click(object sender, EventArgs e)
        {
            try
            {
                //Hide this form and open the Edit Form
                EditWin editwin = new EditWin();
                editwin.EditData = EditData;
                //Subscribe to the event when the Editwin form Button Save is Clicked
                editwin.DataSaved += editwin_Datasaved;
                //Show this form when EditWin Form is closed
                editwin.FormClosed += (s, args) => this.Show();
                editwin.Show();
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }
        private void editwin_Datasaved(System.Data.DataTable updatedTable)
        {
            //This Function will run when the event is raised
            //The table from Editwin will Replace the EditData
            EditData = updatedTable;
            //Save the changed value in the list below
            HashSet<int> EditedTable = new HashSet<int>();
            foreach (DataRow row in EditData.Rows) 
            {
                //Save the Row in EditedTable list if it doesnt exist
                if (!EditedTable.Contains(EditData.Rows.IndexOf(row)))
                {
                    //Debug.WriteLine($"Currently on Row: {EditData.Rows.IndexOf(row).ToString()}, Add to List");
                    EditedTable.Add(EditData.Rows.IndexOf(row));
                    //Debug.WriteLine("Current List contents: " + string.Join(", ", EditedTable));
                }
                foreach (DataColumn col in EditData.Columns) {
                    //If the said Row Has Blank remove the Row from list
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
            //Replace the Row of the original data with the row of the Edited Value from user
            //It replace the row by replacing detecting the index of the non blank row and get the row address from celladdr
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

        private void ResetButton_Click(object sender, EventArgs e)
        {
            ResetData();
            ResetTarget();
        }
        private void ResetData() 
        {
            From = null;
            sheet = null;
            workbook = null;
            Converted = false;
            ConvertFromCSV = null;
            FileType.Clear();
            FileName.Clear();
            CellAddr.Clear();
            EditButton.Enabled = false;
            Converted = false;

        }
        private void ResetTarget() 
        {
            To = null;
            TargetSheet.Items.Clear();
            ToSheet = null;
            PasteBook = null;
            TargetName.Clear();

        }
        private void SelectTarget_Click(object sender, EventArgs e)
        {
            try
            {
                ResetTarget();
                ofd.Title = "Select Excel File";
                ofd.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    TargetFile();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
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
            try {
                //MessageBox.Show($"Sending To {To} Chosed: {ToSheet}");
                var LastRowTo = From.LastRowUsed().RowNumber();
                for (var row = 0; row < LastRowTo; row++)
                {
                    To.Cell(17 + row, 3).Value = From.Cell(2 + row, 4).Value; //板厚
                    To.Cell(17 + row, 7).Value = From.Cell(2 + row, 5).Value; //寸法W
                    To.Cell(17 + row, 8).Value = From.Cell(2 + row, 6).Value; //寸法H
                    To.Cell(17 + row, 14).Value = From.Cell(2 + row, 7).Value; //歩留
                    string CleanTimeData = From.Cell(2 + row, 9).Value.ToString();
                    var CharToRemove = "()（）分";
                    foreach (char c in CharToRemove)
                    {
                        CleanTimeData = CleanTimeData.Replace(c.ToString(), "");
                    }
                    To.Cell(17 + row, 12).Value = CleanTimeData; //三菱加工時間
                }
                PasteBook.Save();
                if (Converted)
                {
                    StringBuilder CSVConv = new StringBuilder();
                    for (int row = 1; row < From.LastRowUsed().RowNumber(); row++)
                    {
                        for (int col = 1; col < From.LastColumnUsed().ColumnNumber(); col++)
                        {
                            var cellvalue = From.Cell(row, col).Value.ToString();
                            if (cellvalue.Contains(",")) {
                                cellvalue.Replace(",", ".");
                            }
                            CSVConv.Append(cellvalue);
                        }
                        CSVConv.AppendLine();
                    }
                    string BatchDate = ToSheet.Name;
                    int dotPos = ConvertFromCSV.LastIndexOf(".");
                    ConvertFromCSV = ConvertFromCSV.Substring(0, dotPos) + " " + BatchDate + ".CSV";
                    File.WriteAllText(ConvertFromCSV, CSVConv.ToString());

                }
                MessageBox.Show($"Operation Finished");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message );
            }
            }
    }
}
