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
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DateTime = System.DateTime;
using System.Resources;
using ExcelInteropGUI.Properties;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Action = System.Action;



namespace ExcelInteropGUI
{
    public partial class Menu : Form
    {
        //Declare FIle Dialog
        OpenFileDialog ofd = new OpenFileDialog();

        //Data Table to Edit the File Directly
        System.Data.DataTable EditData { get; set; } = new System.Data.DataTable();
        System.Data.DataTable TempTableToConvert = new System.Data.DataTable();
        public System.Data.DataTable DataTable { get; set; } = new System.Data.DataTable();
        public System.Data.DataTable TargetTable { get; set; } = new System.Data.DataTable();

        //INT Variable
        int selectedSheet =1;
        int DataChecker, TargetType;

        //Closed XML Variable
        IXLWorksheet From, To,sheet, ToSheet;
        XLWorkbook workbook, PasteBook;

        //String
        string ConvertFromCSV;
        string Tp;
        string[] PresetAddr;
        public string Selected_json;

        //Bool
        bool Converted;
        bool SourceFileClicked, TargetFileClicked;
        public bool Japanese;

        //Action
        public Action<string> OnFunctionStart;
        public event Action editPresetClicked;

        //Culture Info
        CultureInfo culture;

        //List
        List<int> CellAddr = new List<int>();
        List<(string index, string setting, int PresetRow, int PresetCol)> preset = new List<(string index, string setting, int PresetRow, int PresetCol)>();
        List<(int index, string type, string rmvchr)> dataHandle = new List<(int index, string type, string rmvchr)>();
        List<(string index, string setting, int PresetRow, int PresetCol)> GetAddrData = new List<(string index, string setting, int PresetRow, int PresetCol)>();
        List<(string index, string setting, int PresetRow, int PresetCol)> GetAddrTarget = new List<(string index, string setting, int PresetRow, int PresetCol)>();
        List<string> PresetPath = new List<string>();

        //Menu Form Function
        public Menu()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            Japanese_RB.Checked = true;
            Japanese = true;
            this.EditPreset.Location = new System.Drawing.Point(MakePreset.Bounds.Right + 20, MakePreset.Bounds.Y);
            this.DeleteBtn.Location = new System.Drawing.Point(EditPreset.Bounds.Right + 20, EditPreset.Bounds.Y);

            this.DataSelected.Location = new System.Drawing.Point(SelectData.Location.X, SelectData.Location.Y+40);
            this.FileName.Location = new System.Drawing.Point(SelectData.Bounds.Right + 20, SelectData.Location.Y);
            this.FileType.Location = new System.Drawing.Point(FileName.Location.X, DataSelected.Location.Y);

            this.label2.Location = new System.Drawing.Point(DataSelected.Location.X, FileType.Location.Y + 40);
            this.CSVNew_Save.Location = new System.Drawing.Point(FileType.Location.X, label2.Location.Y );

            this.PresetLabel.Location = new System.Drawing.Point(label2.Location.X, label2.Location.Y+40);
            this.SelectPreset.Location = new System.Drawing.Point(CSVNew_Save.Location.X, PresetLabel.Location.Y);


            this.SelectTarget.Location = new System.Drawing.Point(FileName.Bounds.Right + 30, SelectData.Location.Y);
            this.TargetName.Location = new System.Drawing.Point(SelectTarget.Bounds.Right+20, SelectTarget.Location.Y);
            this.SelectSheet.Location = new System.Drawing.Point(SelectTarget.Location.X, SelectTarget.Location.Y + 40);
            this.TargetSheet.Location = new System.Drawing.Point(TargetName.Location.X, SelectSheet.Location.Y);
            this.CopyName.Location = new System.Drawing.Point(SelectSheet.Location.X, SelectSheet.Location.Y + 40);
            this.New_Sheet.Location = new System.Drawing.Point(TargetSheet.Location.X,CopyName.Location.Y);
            this.Add_NewSheet.Location = new System.Drawing.Point(TargetSheet.Bounds.Right-this.Add_NewSheet.Width,this.Add_NewSheet.Location.Y);
            this.label1.Location = new System.Drawing.Point(CopyName.Location.X, CopyName.Location.Y+40);
            this.NewBookSave.Location = new System.Drawing.Point(New_Sheet.Location.X,label1.Location.Y);
            this.FolderBtn.Location = new System.Drawing.Point(Add_NewSheet.Bounds.Right - this.FolderBtn.Width, NewBookSave.Location.Y + ((NewBookSave.Height/2) - (FolderBtn.Height/2)));

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            SendButton.Enabled = false;
            EditButton.Enabled = false;
            PresetAddr = Directory.GetFiles( Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Preset"), "*.json");
            foreach (string p in PresetAddr) {
                SelectPreset.Items.Add( Path.GetFileName(p));
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

        //When user click the data button
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
                CSVNew_Save.Text = FileName.Text.Substring(0 , FileName.Text.LastIndexOf(".")) + " - "+DateTime.Now.ToString("yyyy-MM-dd");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

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
                for (int i = 1; i <= lastCol; i++)
                {
                    DataTable.Columns.Add();
                }
                for (int i = 0; i < lastRow; i++)
                {
                    DataRow row = DataTable.NewRow();
                    for (int j = 0; j < lastCol; j++)
                    {
                        row[j] = From.Cell(i + 1, j + 1).Value;
                    }
                    DataTable.Rows.Add(row);
                }
                //Check the validity off all the data for contamination
                OnFunctionStart?.Invoke("Checking For Contamination");
                for (int i = 0; i < lastRow - 1; i++)
                {
                    if (!sheet.Cell(2, 1).Value.Equals(sheet.Cell(2 + i, 1).Value))
                    {
                        MessageBox.Show(Languages.ContaminationWarning + "\n" + sheet.Cell(2 + i, 1) +":"+ sheet.Cell(2 + i, 1).Value);
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
                    EditButton.Enabled = CellAddr.Count > 0;
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
                        DataChecker = 3;
                        break;

                }
                SourceFileClicked = true;
                SafeInvoke(this, () =>
                {
                    CheckSendBtn();
                });
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

        //When user click target button
        private void SelectTarget_Click(object sender, EventArgs e)
        {
            try
            {
                ofd.Title = "Select Excel File";
                ofd.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    //TargetFile();
                    using (LoadingBar load = new LoadingBar(TargetFile, this))
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
        private void TargetSheet_Click(object sender, EventArgs e)
        {
            if (PasteBook != null)
            {
                TargetSheet.Items.Clear();
                foreach (var item in PasteBook.Worksheets)
                    TargetSheet.Items.Add(item);

            }

        }
        private void TargetSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            OnFunctionStart?.Invoke("Resetting Table");
            selectedSheet = TargetSheet.SelectedIndex;
            ToSheet = PasteBook.Worksheet(selectedSheet + 1);
            To = ToSheet;
            int dotpos;
            if (TargetName.Text.Contains('-'))
            {
                 dotpos= TargetName.Text.IndexOf('-');
            }
            else
            {
                dotpos = TargetName.Text.IndexOf('.');
            }
            NewBookSave.Text = TargetName.Text.Substring(0, dotpos) + " - " + TargetSheet.Text;
        }
        private void TargetFile()
        {
            OnFunctionStart?.Invoke("Resetting the data");
            ResetTarget();
            Tp = ofd.FileName;
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
                    TargetSheet.SelectedIndex = 0;
                });
                //Process the file if the file has weird format

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
            TargetFileClicked = true;
            SafeInvoke(this, () =>
            {
                CheckSendBtn();
            });
            if (To.RangeUsed() != null)
            {
                OnFunctionStart?.Invoke("Mapping Table");
                var lastrow = To.LastRowUsed().RowNumber();
                var lastcol = To.LastColumnUsed().ColumnNumber();


                // Add columns to the DataTable
                for (int i = 0; i < lastcol; i++)
                {
                    TargetTable.Columns.Add();
                }

                for (int i = 1; i <= 100; i++) // Use 1-based index for rows
                {
                    DataRow datarow = TargetTable.NewRow();
                    for (int j = 1; j <= lastcol; j++) // Use 1-based index for columns
                    {
                        if (!To.Cell(i, j).HasFormula)
                        {
                            var cellValue = To.Cell(i, j).Value; // Store cell value in a variable
                            datarow[j - 1] = cellValue; // Adjust for 0-based index in DataRow
                        }
                    }
                    TargetTable.Rows.Add(datarow);
                }

            }
        }
        private void Add_NewSheet_Click(object sender, EventArgs e)
        {
            if (PasteBook != null)
            {
                var copysheet = PasteBook.AddWorksheet(New_Sheet.Text);
                var firstrow = To.FirstRowUsed().RowNumber();
                var firscol = To.FirstColumnUsed().ColumnNumber();
                List<int> cells = new List<int>();
                foreach (var col in To.Columns())
                {
                    if (col.IsHidden)
                    {
                        cells.Add(col.ColumnNumber());
                    }
                }
                To.RangeUsed().CopyTo(copysheet.Cell(firstrow,firscol));
                foreach(var cell in cells)
                {
                    copysheet.Column(cell).Hide();
                }
            }
            else
                MessageBox.Show(Languages.TargetFilenotSelected);
        }

        //Edit the Data Button
        private void EditButton_Click(object sender, EventArgs e)
        {
            try
            {
                //Hide this form and open the Edit Form
                EditWin editwin = new EditWin();
                editwin.EditData = EditData;
                editwin.Japanese = Japanese;
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
        private void editwin_Datasaved(System.Data.DataTable updatedTable, List<(int ChangedRow, int ChangedCol, object ChangedVal)> values)
        {
            //This Function will run when the event is raised
            //The table from Editwin will Replace the EditData
            EditData = updatedTable;
            foreach(var item in values)
            {
                Debug.WriteLine("Before: "+From.Cell(CellAddr[item.ChangedRow], item.ChangedCol+1).Value);
                From.Cell(CellAddr[item.ChangedRow], item.ChangedCol+1).Value = EditData.Rows[item.ChangedRow][item.ChangedCol]?.ToString();
                Debug.WriteLine("After: " + From.Cell(CellAddr[item.ChangedRow], item.ChangedCol+1).Value);
            }

        }
        
        //Send data
        private void SendButton_Click(object sender, EventArgs e)
        {
            var confirmation = MessageBox.Show(Languages.FillinConfirmation, Languages.FillinConfirmationLabel, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirmation == DialogResult.Yes)
                try
                {
                    if (SelectPreset.SelectedItem == null || string.IsNullOrWhiteSpace(SelectPreset.Text))
                    {
                        MessageBox.Show(Languages.PresetNotSelected);
                        return;
                    }
                    if (TargetSheet.SelectedItem == null || string.IsNullOrWhiteSpace(TargetSheet.Text))
                    {
                        MessageBox.Show(Languages.SheetNotSelected);
                        return;
                    }
                    if (DataChecker != TargetType)
                    {
                        MessageBox.Show(Languages.FileIsntCompatible);
                        return;
                    }

                    using (LoadingBar load = new LoadingBar(TransferData, this))
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
                    else
                    {
                        using (LoadingBar load = new LoadingBar(() =>
                        {
                            int dotpos = TargetName.Text.LastIndexOf(".");
                            string FullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "レーザー機払出管理");
                            workbook.SaveAs(Path.Combine(FullPath, NewBookSave.Text, $"{CSVNew_Save.Text}.xlsx"));
                        }, this))
                        {
                            load.ShowDialog();
                        }
                    }
                    MessageBox.Show(Languages.OperationFinished);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            else
                return;
        }
        private void TransferData()
        {
            var LastRowTo = From.LastRowUsed().RowNumber();
            for (var row = 0; row < LastRowTo - 1; row++)
            {
                for (int i = 0; i < GetAddrTarget.Count(); i++)
                {
                    if (
                        GetAddrTarget[i].PresetRow.Equals(0) ||
                        GetAddrTarget[i].PresetCol.Equals(0) ||
                        GetAddrData[i].PresetRow.Equals(0) ||
                        GetAddrData[i].PresetCol.Equals(0))
                    {
                        continue;
                    }
                    var valuetoParse = From.Cell(GetAddrData[i].PresetRow + row, GetAddrData[i].PresetCol).Value;
                    if (!string.IsNullOrEmpty(dataHandle[i].rmvchr))
                    {
                        foreach (char c in dataHandle[i].rmvchr)
                        {
                            valuetoParse = valuetoParse.ToString().Replace(c.ToString(), "");
                        }
                    }
                    switch (dataHandle[i].type)
                    {
                        case "Number":
                            if (int.TryParse(valuetoParse.ToString(), out var result))
                            {
                                valuetoParse = result;
                                break;
                            }
                            else if (float.TryParse(valuetoParse.ToString(), out var resultFloat))
                            {
                                valuetoParse = resultFloat;
                            }
                            else { break; }

                            break;

                        case "Text":
                            valuetoParse = valuetoParse.ToString();
                            break;

                    }
                    To.Cell(GetAddrTarget[i].PresetRow + row, GetAddrTarget[i].PresetCol).Value = valuetoParse;
                }
            }

            OnFunctionStart?.Invoke("Saving Book");
            string FullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "レーザー機械管理");
            if (Directory.Exists(FullPath))
            {
                PasteBook.SaveAs(Path.Combine(FullPath, NewBookSave.Text, $"{NewBookSave.Text}.xlsx"));
            }
            else
            {
                Directory.CreateDirectory(FullPath);
                PasteBook.SaveAs(Path.Combine(FullPath, NewBookSave.Text, $"{NewBookSave.Text}.xlsx"));
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
            string doc = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "レーザー機払出管理");
            if(!Directory.Exists(doc))
                Directory.CreateDirectory(doc);
            string fullpath = Path.Combine(doc, NewBookSave.Text, $"{BatchDate}.CSV");
            File.WriteAllText(fullpath, CSVConv.ToString());
            if (File.Exists(ConvertFromCSV))
            {
                File.Delete(ConvertFromCSV);
            }

        }
        private void FolderBtn_Click(object sender, EventArgs e)
        {
            Process.Start(Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "レーザー機払出管理"), NewBookSave.Text, $"{NewBookSave.Text}.xlsx"));
        }

        // edit/make Preset
        private void makePreset_Click(object sender, EventArgs e)
        {
            
            OnFunctionStart?.Invoke("Making Table");
            Setting setting = new Setting(this);
            setting.Japanese = Japanese;
            setting.DataTable = DataTable;
            setting.to = TargetName.Text;
            setting.from = FileName.Text;
            setting.TargetTable = TargetTable;
            setting.Preset = preset;
            setting.datahandletoRtn = dataHandle;
            setting.FormClosed += (s, args) => { 
                this.Show();
                SelectPreset_SelectedIndexChanged(null, EventArgs.Empty);
            };
            if (!string.IsNullOrEmpty(TargetName.Text) && preset != null && !string.IsNullOrEmpty(FileType.Text))
                editPresetClicked?.Invoke();
            setting.FormClosed += (s, args) => this.Show();
            setting.Show();
            this.Hide();

        }
        public void EditPreset_Click(object sender, EventArgs e)
        {
            if (SelectPreset.SelectedItem != null || !string.IsNullOrWhiteSpace(SelectPreset.Text))
            {
                makePreset_Click(sender, e);
            }
            else
                MessageBox.Show(Languages.PresetNotSelected);
        }
        private void SelectPreset_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SelectPreset.SelectedItem != null || !string.IsNullOrWhiteSpace(SelectPreset.Text))
            {
                ReadJson();
            }
           
        }
        private void SelectPreset_Click(object sender, EventArgs e)
        {
            SelectPreset.Items.Clear();
            PresetAddr = Directory.GetFiles(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Preset"), "*.json");
            foreach (string p in PresetAddr)
            {
                SelectPreset.Items.Add(Path.GetFileName(p));
            }
        }
        private void Delete_Click(object sender, EventArgs e)
        {
            if(SelectPreset.SelectedItem == null)
            {
                MessageBox.Show(Languages.NoFiletoDelete);
            }
            else
            {
                preset.Clear();
                DialogResult deleteconfirmation = MessageBox.Show(Languages.DeleteConfirmation + "\n" + SelectPreset.SelectedItem,
                    "Unsaved Change",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                switch (deleteconfirmation)
                {
                    case DialogResult.Yes:
                        File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Preset", SelectPreset.SelectedItem.ToString()));
                        SelectPreset.Items.Clear();
                        PresetAddr = Directory.GetFiles(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Preset"), "*.json");
                        foreach (string p in PresetAddr)
                        {
                            SelectPreset.Items.Add(Path.GetFileName(p));
                        }
                        break;
                    case DialogResult.No:
                        return;
                }
            }

        }
        private void ReadJson()
        {
            string file = SelectPreset.SelectedItem.ToString();
            Selected_json = file.Substring(0, file.LastIndexOf("."));
            string json = File.ReadAllText($"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Preset", SelectPreset.SelectedItem.ToString())}");
            dynamic data = JsonConvert.DeserializeObject<dynamic>(json);
            GetAddrData.Clear();
            GetAddrTarget.Clear();
            dataHandle.Clear();
            preset.Clear();
            if (data.Preset != null)
            {
                foreach (var item in data.Preset)
                {
                    string index = item.Item1;
                    string setting = item.Item2;
                    int presetRow = (int)item.Item3;
                    int presetCol = (int)item.Item4;
                    preset.Add((index, setting, presetRow, presetCol));
                }
            }

            int Highest = preset.Select(t => int.Parse(t.index.Last().ToString())).Max();
            for (int i = 0; i < Highest; i++)
            {
                GetAddrData.Add(("", "", 0, 0));
                GetAddrTarget.Add(("", "", 0, 0));
                dataHandle.Add((0, "", ""));
            }
            foreach (var item in preset)
            {
                if (item.index.Contains("Data"))
                {
                    GetAddrData[int.Parse(item.index.Last().ToString()) - 1] = item;
                }
                else
                {
                    GetAddrTarget[int.Parse(item.index.Last().ToString()) - 1] = item;
                }
            }
            if (data.datahandle != null)
            {
                foreach (var item in data.datahandle)
                {
                    string item1 = item.Item1;
                    string item2 = item.Item2;
                    string item3 = item.Item3;
                    if (int.Parse(item1.Last().ToString()) > 0)
                    {
                        dataHandle[int.Parse(item1.Last().ToString()) - 1] = ((int.Parse(item1.Last().ToString()), item2, item3));
                    }
                }

            }
        }

        //Reset Button
        private void ResetButton_Click(object sender, EventArgs e)
        {
            ResetData();
            ResetTarget();
        }
        private void ResetData() 
        {
            From = null;
            sheet = null;
            Tp = null;
            workbook = null;
            Converted = false;
            ConvertFromCSV = null;
            SafeInvoke(this, () =>
            {
                FileType.Clear();
                FileName.Clear();
                EditButton.Enabled = false;
                CSVNew_Save.Clear();
            });
            SharedData.Log.Clear();
            SharedData.NewValues.Clear();
            DataTable.Reset();
            CellAddr.Clear();
            Converted = false;
            EditData.Reset();
            TempTableToConvert.Reset();
            SourceFileClicked = false;
        }
        private void ResetTarget() 
        {
            To = null;
            SafeInvoke(this, () =>
            {
                TargetSheet.Items.Clear();
                TargetName.Clear();
                New_Sheet.Clear();
                NewBookSave.Clear();
            });
            ToSheet = null;
            PasteBook = null;
            TargetFileClicked = false;
            TargetTable.Reset();
        }

        private void PresetLabel_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void FileType_TextChanged(object sender, EventArgs e)
        {

        }

        //Update the UI within multithread
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

        //Unused or unrelated 
        private void CheckSendBtn()
        {
            if (TargetFileClicked && SourceFileClicked) SendButton.Enabled = true;
        }
        private void SelectPreset_Validating(object sender, CancelEventArgs e)
        {
        }
        private void SelectPreset_Leave(object sender, EventArgs e)
        {
        }

        //Change Language
        private void English_RB_CheckedChanged(object sender, EventArgs e)
        {
            if (English_RB.Checked)
            {
                Japanese = false;
                ChangeLanguage.change("en-EN", this);
                culture = new CultureInfo("en-US");
                Thread.CurrentThread.CurrentCulture = culture;
                Thread.CurrentThread.CurrentUICulture = culture;
            }
        }

        private void Japanese_RB_CheckedChanged(object sender, EventArgs e)
        {
            if (Japanese_RB.Checked)
            {
                Japanese = true;
                ChangeLanguage.change("ja-JP",this);
                culture = new CultureInfo("ja-JP");
                Thread.CurrentThread.CurrentCulture = culture;
                Thread.CurrentThread.CurrentUICulture = culture;
            }
        }
    }
}
