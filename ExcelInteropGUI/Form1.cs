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
using System.Threading;

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
        int DataChecker, TargetType;
        public Action<string> OnFunctionStart;


        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            SendButton.Enabled = false;
            EditButton.Enabled = false;
            
        }
        private void SelectData_Click(object sender, EventArgs e)
        {
            try
            {
                ofd.Title = "Select Excel File";
                ofd.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    using (LoadingBar load = new LoadingBar(LoadData, this))
                    {
                        load.ShowDialog(this);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }


        private void SelectTarget_Click(object sender, EventArgs e)
        {
            try
            {
                ofd.Title = "Select Excel File";
                ofd.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    using(LoadingBar load = new LoadingBar(TargetFile, this))
                    {
                        load.ShowDialog(this);
                    } 
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
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
        private void TargetSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedSheet = TargetSheet.SelectedIndex;
            ToSheet = PasteBook.Worksheet(selectedSheet + 1);
            To = ToSheet;
        }

        private void SendButton_Click(object sender, EventArgs e)
        {
            try
            {
                if(DataChecker != TargetType)
                {
                    MessageBox.Show("Selected File Isn't Compatible");
                    return;
                }
                using(LoadingBar load = new LoadingBar(TransferData, this))
                {
                    load.ShowDialog();
                }
                if (Converted)
                {
                    using (LoadingBar load = new LoadingBar(SaveBacktoCSV, this))
                    {
                        load.ShowDialog();
                    }
                }
                MessageBox.Show($"Operation Finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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
            OnFunctionStart?.Invoke("Clean the Data");
            ResetData();
            //The File Path of the selected file
            string fp = ofd.FileName;
            //The File Name of the file
            string fn = Path.GetFileName(fp);
            if (!string.IsNullOrEmpty(fp))
            {
                //Check if the file is CSV extension
                if (fp.EndsWith(".csv"))
                {
                    OnFunctionStart?.Invoke("Processing CSV");
                    CSVConvert();
                    fp = ConvertFromCSV;
                }
                //Setup for the Workbook Data
                SafeInvoke(FileName, () =>
                {
                    FileName.Text = fn;
                });

                workbook = new XLWorkbook(fp);
                //workbook.CalculateMode = XLCalculateMode.Manual;
                sheet = workbook.Worksheet(1);
                From = sheet;
                var lastRow = sheet.LastRowUsed().RowNumber();
                var lastCol = sheet.LastColumnUsed().ColumnNumber();
                var DataRange = sheet.Range(2, 1, lastRow, lastCol);
                //Check the validity off all the data for contamination
                OnFunctionStart?.Invoke("Checking For Contamination");
                for (int i = 0; i < lastRow - 1; i++)
                {
                    if (!sheet.Cell(2, 1).Value.Equals(sheet.Cell(2 + i, 1).Value))
                    {
                        MessageBox.Show($"Cell in: {sheet.Cell(2 + i, 1)} has value of: {sheet.Cell(2 + i, 1).Value}");
                        fp = null;
                        SafeInvoke(FileName, () =>
                        {
                            FileName.Text = null;
                        });
                        return;
                    }
                }
                OnFunctionStart?.Invoke("Checking For Blank");
                DetectBlank(DataRange, lastCol);
                SafeInvoke(FileType, () =>
                {
                    FileType.Text = sheet.Cell(2, 1).Value.ToString();
                    EditButton.Enabled = CellAddr.Count> 0;
                });
                OnFunctionStart?.Invoke("Determining the File");
                //Check if the the data and the targetfile is the samefile
                switch (fn.ToUpper())
                {
                   case string s when s.Contains("MITSUBISHI"):
                        DataChecker = 1;
                        break;
                    case string s when s.Contains("KOMATSU"):
                        DataChecker = 2;
                        break;
                    case string s when s.Contains("ASTES"):
                        DataChecker= 3;
                        break;
                    
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
                    if (!CellAddr.Contains(cell.Address.RowNumber))
                    {
                        CellAddr.Add(cell.Address.RowNumber);
                    }
                }
            }
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

        private void SaveBacktoCSV()
        {
            StringBuilder CSVConv = new StringBuilder();
            for (int row = 1; row < From.LastRowUsed().RowNumber(); row++)
            {
                for (int col = 1; col < From.LastColumnUsed().ColumnNumber(); col++)
                {
                    var cellvalue = From.Cell(row, col).Value.ToString();
                    if (cellvalue.Contains(","))
                    {
                        cellvalue.Replace(",", ".");
                    }

                    CSVConv.Append(cellvalue + ",");
                }
                CSVConv.AppendLine();
            }
            string BatchDate = ToSheet.Name;
            int dotPos = ConvertFromCSV.LastIndexOf("\\");
            ConvertFromCSV = ConvertFromCSV.Substring(0, dotPos) + "\\" + FileType.Text + " " + BatchDate + ".CSV";
            File.WriteAllText(ConvertFromCSV, CSVConv.ToString());
            if (File.Exists(ConvertFromCSV))
            {
                File.Delete(ConvertFromCSV);
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
                    EditedTable.Add(EditData.Rows.IndexOf(row));
                }
                foreach (DataColumn col in EditData.Columns) {
                    //If the said Row Has Blank remove the Row from list

                    if (row[col] == DBNull.Value || string.IsNullOrWhiteSpace(row[col].ToString()))
                    {
                        EditedTable.Remove(EditData.Rows.IndexOf(row));
                        break ;
                    }
                }
            }

            //Replace the Row of the original data with the row of the Edited Value from user
            //It replace the row by replacing detecting the index of the non blank row and get the row address from celladdr
            foreach (var index in EditedTable)
            {
                foreach(DataColumn colins in EditData.Columns) 
                {
                    var x = CellAddr[index];
                    var y = EditData.Columns.IndexOf(colins)+1;
                    From.Cell(x, y).Value = EditData.Rows[index][colins]?.ToString();
                }
            }

            EditedTable.Clear();
        }
        private void TransferData()
        {
            var LastRowTo = From.LastRowUsed().RowNumber();
            for (var row = 0; row < LastRowTo-1; row++)
            {
                float code = float.Parse(From.Cell(2 + row, 4).Value.ToString());
                OnFunctionStart?.Invoke($"Fill in 板厚 on {To.Name} in Cell: {To.Cell(17 + row, 3).Address}");
                To.Cell(17 + row, 3).Value = code; //板厚
                OnFunctionStart?.Invoke($"Fill in 寸法W on {To.Name} in Cell: {To.Cell(17 + row, 7).Address}");
                To.Cell(17 + row, 7).Value = From.Cell(2 + row, 5).Value; //寸法W
                OnFunctionStart?.Invoke($"Fill in 寸法H on {To.Name} in Cell: {To.Cell(17 + row, 8).Address}");
                To.Cell(17 + row, 8).Value = From.Cell(2 + row, 6).Value; //寸法H
                OnFunctionStart?.Invoke($"Fill in 歩留 on {To.Name} in Cell: {To.Cell(17 + row, 14).Address}");
                To.Cell(17 + row, 14).Value = From.Cell(2 + row, 7).Value; //歩留
                OnFunctionStart?.Invoke($"Fill in 三菱加工時間 on {To.Name} in Cell: {To.Cell(17 + row, 12).Address}");
                string CleanTimeData = From.Cell(2 + row, 9).Value.ToString();
                var CharToRemove = "()（）分";
                foreach (char c in CharToRemove)
                {
                    CleanTimeData = CleanTimeData.Replace(c.ToString(), "");
                }
                To.Cell(17 + row, 12).Value = CleanTimeData; //三菱加工時間
            }
            OnFunctionStart?.Invoke("Saving Book");
            PasteBook.Save();
        }
        private void TargetFile()
        {
            OnFunctionStart?.Invoke("Resetting the data");
            ResetTarget();
            string Tp = ofd.FileName;
            string Tn = Path.GetFileName(Tp);
            OnFunctionStart?.Invoke("Reading the File");
            if (!string.IsNullOrEmpty(Tp))
            {
                PasteBook = new XLWorkbook(Tp);
                PasteBook.CalculateMode = XLCalculateMode.Auto;
                OnFunctionStart?.Invoke("Reading the Sheets");
                foreach (var Sheet in PasteBook.Worksheets)
                {
                    SafeInvoke(TargetSheet, () =>
                    {
                        TargetSheet.Items.Add(Sheet);
                    });
                }
                SafeInvoke(this, () =>
                {
                    TargetName.Text = Tn;
                    SendButton.Enabled = true;
                    TargetSheet.SelectedIndex = 0;
                });

            }
            switch (Tn.ToUpper())
            {
                case string s when s.Contains("MITSUBISHI"):
                    TargetType = 1;
                    break;
                case string s when s.Contains("KOMATSU"):
                    TargetType = 2;
                    break;
                case string s when s.Contains("ASTES"):
                    TargetType = 3;
                    break;

            }
        }
        private void ResetData() 
        {
            From = null;
            sheet = null;
            workbook = null;
            Converted = false;
            ConvertFromCSV = null;
            SafeInvoke(this, () =>
            {
                FileType.Clear();
                FileName.Clear();
                EditButton.Enabled = false;
            });
            CellAddr.Clear();
            Converted = false;
            EditData.Reset();
            TempTableToConvert.Reset();
        }
        private void ResetTarget() 
        {
            To = null;
            SafeInvoke(this, () =>
            {
                TargetSheet.Items.Clear();
                TargetName.Clear();
            });
            ToSheet = null;
            PasteBook = null;

        }

        private void SafeInvoke(System.Windows.Forms.Control control, Action uiUpdate)
        {
            if (control.InvokeRequired)
            {
                control.Invoke(uiUpdate);
            }
            else { 
            uiUpdate();
            }
        }

    }
}
