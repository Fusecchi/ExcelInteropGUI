using DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Path = System.IO.Path;


namespace ExcelInteropGUI
{
    public partial class Setting : Form

    {
        public System.Data.DataTable DataTable { get; set; } = new System.Data.DataTable();
        public System.Data.DataTable TargetTable { get; set; } = new System.Data.DataTable();
        public List<(string DataIndex, string setting, int PresetRow, int PresetCol)> Preset { get; set; } = new List<(string DataIndex, string setting, int PresetRow, int PresetCol)>();
        public List<(string DataIndex, string handle, string remove)> datahandle { get; set; } = new List<(string DataIndex, string handle, string remove)>();
        public List<(int DataIndex, string handle, string remove)> datahandletoRtn { get; set; } = new List<(int DataIndex, string handle, string remove)>();

        private List<GroupBox> groupBoxes = new List<GroupBox>();
        private int BtnIterration = 0;
        public string from, to;

        public Setting(Menu menu)
        {
            InitializeComponent();
            menu.editPresetClicked += EditPreset;
<<<<<<< HEAD
            textBox1.Text = menu.Selected_json;
            this.FormClosed += (s, args) => menu.editPresetClicked -= EditPreset;
            this.MaximizeBox = false;
=======
            this.FormClosed += (s, args) => menu.editPresetClicked -= EditPreset;
>>>>>>> 0d5efcebdb9e9205644ac9bead6c701bfb0d828f
        }


        private void button1_Click(object sender, EventArgs e)
        {
            var data = new
            {
                Preset = Preset,
                datahandle = datahandle
        };
            string json =   JsonConvert.SerializeObject(data, Formatting.Indented);
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("File Name Can't be Empty");
                return;
            }
            File.WriteAllText($"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Preset",textBox1.Text)}.json", json);
            MessageBox.Show($"Saved as {textBox1.Text}.json");
<<<<<<< HEAD

=======
>>>>>>> 0d5efcebdb9e9205644ac9bead6c701bfb0d828f
        }
        private void button3_Click(object sender, EventArgs e)
        {
            BtnIterration++;
            //Button Property
            int buttonHeight = 30;
            int buttonWidth = 100;
            //Lbutton Property
            int LbuttonHeight = 30;
            int LbuttonWidth = 100;
            int GapX = 15;
            int ypos = (100) * (BtnIterration - 1) + scrollPanel.AutoScrollPosition.Y;
            scrollPanel.AutoScrollMinSize = new Size(Width, ypos);
            //Initiate Button
            GroupBox groupBox = new GroupBox
            {
                Location = new System.Drawing.Point(10, 10),
            };
            Button button = new Button {
                Text = Text = $"Data {BtnIterration}",
                Size = new Size(buttonWidth, buttonHeight),
                Location = new System.Drawing.Point(5, ypos)
            };
            //Initiate Label
            Label label = new Label
            {
                Location = new System.Drawing.Point(button.Bounds.Right, ypos),
                Size = new Size(50, 30),
                Tag = $"Data {BtnIterration}"
            };

            //Initiate Radio Button
            RadioButton RbInt = new RadioButton
            {
                Location = new System.Drawing.Point(label.Location.X,label.Bounds.Bottom),
                Font = new Font("Microsoft Sans Serif",8),
                Text = "Number",
                AutoSize = true,
            };
            RadioButton RbStr = new RadioButton
            {
                Location = new System.Drawing.Point(RbInt.Bounds.Right, RbInt.Location.Y),
                Text = "None",
                Font = new Font("Microsoft Sans Serif", 8),
                AutoSize = true,
            };
            RadioButton RbFloat = new RadioButton
            {
                Location = new System.Drawing.Point(RbStr.Bounds.Right, RbInt.Location.Y),
                Text = "Text",
                Font = new Font("Microsoft Sans Serif", 8),
                AutoSize = true,
            };
            TextMeasure(RbInt, button, label.Bounds.Bottom);
            TextMeasure(RbFloat, RbInt, label.Bounds.Bottom);
            TextMeasure(RbStr, RbFloat, label.Bounds.Bottom);
            TextBox chartoRemove = new TextBox {
                Size = new Size(RbStr.Bounds.Right-RbInt.Bounds.Left, label.Size.Height),
                Location = new System.Drawing.Point(label.Bounds.Right, label.Location.Y),
                Visible = true
            };

            Button Lbutton = new Button
            {
                Text = $"Target File {BtnIterration}",
                Size = new Size(LbuttonWidth, LbuttonHeight),
                Location = new System.Drawing.Point(chartoRemove.Bounds.Right + GapX, ypos)
            };
            Label Llabel = new Label
            {
<<<<<<< HEAD
                Location = new System.Drawing.Point(Lbutton.Width + Lbutton.Location.X, ypos),
                Size = new Size(50, 50),
=======
                Location = new System.Drawing.Point(Lbutton.Width + Lbutton.Location.X + GapX, ypos),
                Size = new Size(50, 30),
>>>>>>> 0d5efcebdb9e9205644ac9bead6c701bfb0d828f
                Tag = $"Target File {BtnIterration}"
            };
            Button RmBtn = new Button
            {
                Size = new Size(buttonWidth, buttonHeight),
                Location = new System.Drawing.Point(((RbStr.Bounds.Right - RbInt.Bounds.Left)/2)+RbInt.Bounds.Left  , RbInt.Bounds.Bottom+GapX),
                Text = "Delete"
            };
  
            groupBox.Controls.Add(button);
            groupBox.Controls.Add(label);
            groupBox.Controls.Add(Lbutton);
            groupBox.Controls.Add(Llabel);
            groupBox.Controls.Add(RbInt);
            groupBox.Controls.Add(RbStr);
            groupBox.Controls.Add(RbFloat);
            groupBox.Controls.Add(chartoRemove);
            groupBox.Controls.Add(RmBtn);
            groupBox.Size = new Size(Llabel.Bounds.Right, RmBtn.Bounds.Bottom);
            scrollPanel.Controls.Add(groupBox);
            groupBoxes.Add(groupBox);
            RmBtn.Click += (s, args) =>
            {
                BtnIterration--;
                scrollPanel.Controls.Remove(groupBox);
                var groupboxind = groupBoxes.IndexOf(groupBox);
                var index = Preset.FindIndex(k => k.DataIndex == button.Text);
                var indexT = Preset.FindIndex(j =>  j.DataIndex == Lbutton.Text);
                if (index != -1 && indexT != -1)
                {
                    if (index< indexT)
                    {
                        Preset.RemoveAt(indexT);
                        Preset.RemoveAt(index);
                    }
                    if(index> indexT)
                    {
                        Preset.RemoveAt(index);
                        Preset.RemoveAt(indexT);
                    }
                }
                if( index != -1)
                {
                    Preset.RemoveAt(index);
                }
                if (indexT != -1)
                {
                    Preset.RemoveAt(indexT);
                }
                var DataHandleIndex = datahandle.FindIndex(l => l.DataIndex == button.Text);
                if (DataHandleIndex != -1)
                {
                    datahandle.RemoveAt(DataHandleIndex);
                }
                groupBoxes.Remove(groupBox);
                for (int i = groupboxind; i<groupBoxes.Count; i++)
                {
                    GroupBox gb = groupBoxes[i];
                    gb.Location = new System.Drawing.Point(gb.Location.X, gb.Location.Y-100);
                    foreach (Control control in gb.Controls)
                    {
                        if(control is Button button1 && button1.Text == $"Data {i+2}"){
                            button1.Text = $"Data {i+1}";
                        }
                        if (control is Button button2 && button2.Text == $"Target File {i + 2}")
                        {
                            button2.Text = $"Target File {i+1}";
                        }
                    }
                }
                for(int j = 0;j<Preset.Count;j++)
                {
                    var original_list = Preset[j];
                    var control = groupboxind + 2;
                    int lastval = int.Parse(Preset[j].DataIndex.Last().ToString());
                    if (lastval >= control)
                    {
                        string newDataIndex = original_list.DataIndex.Substring(0, original_list.DataIndex.Length - 1) + (lastval - 1).ToString();
                        Preset[j] = (newDataIndex, original_list.setting, original_list.PresetRow, original_list.PresetCol);
                    }
                }
            };

            button.Click += (s, args) =>
            {
                SelectDataForm select = new SelectDataForm();
                select.DatatoClick = DataTable;
                List<(int row, int col)> Highlight = new List<(int row, int col)>();
                select.FormClosed += (c, arg) => this.Show();
                for(int j = 0;j<Preset.Count;j++)
                {
                    if (Preset[j].DataIndex == button.Text)
                    {
                        Highlight.Add((Preset[j].PresetRow, Preset[j].PresetCol));
                        select.Highlight = Highlight;
                    }
                }
                select.selectedData += tuple =>
                {
                    int index = Preset.FindIndex(k => k.DataIndex == button.Text);
                    if (index != -1)
                    {
                        Preset[index] = (button.Text, tuple.CellVal.ToString(), tuple.rowInd + 1, tuple.ColInd + 1);
                    }
                    else
                    {
                        Preset.Add((button.Text, tuple.CellVal.ToString(), tuple.rowInd + 1, tuple.ColInd + 1));
                    }

                    label.Text = tuple.CellVal.ToString();
                };
                select.Show();
                this.Hide();
            };
            Lbutton.Click += (s, args) => {
                SelectDataForm TSelect = new SelectDataForm();
                TSelect.DatatoClick = TargetTable;
                List<(int row, int col)> THighlight = new List<(int row, int col)>();
                TSelect.FormClosed += (D, arg) => this.Show();
                for (int j = 0; j < Preset.Count; j++)
                {
                    if (Preset[j].DataIndex == Lbutton.Text)
                    {
                        THighlight.Add((Preset[j].PresetRow, Preset[j].PresetCol));
                        TSelect.Highlight = THighlight;
                    }
                }
                TSelect.selectedData += tuple =>
                {
                    int index = Preset.FindIndex(k => k.DataIndex == Lbutton.Text);
                    if (index != -1)
                    {
                        Preset[index] = (Lbutton.Text, tuple.CellVal.ToString(), tuple.rowInd + 1, tuple.ColInd + 1);
                    }
                    else
                    {
                        Preset.Add((Lbutton.Text, tuple.CellVal.ToString(), tuple.rowInd + 1, tuple.ColInd + 1));
                    }

                    Llabel.Text = tuple.CellVal.ToString();
                };
                TSelect.Show();
                this.Hide();
            };

            chartoRemove.TextChanged += (s, arg) => {
                int index = datahandle.FindIndex(k => k.DataIndex == button.Text);
                string selectedvalue = RbInt.Checked ? RbInt.Text : RbFloat.Checked ? RbFloat.Text : RbStr.Checked ? RbStr.Text : string.Empty;
                if (index != -1)
                {
                    datahandle[index] = (button.Text,selectedvalue, chartoRemove.Text);
                }
                else
                {
                    datahandle.Add((button.Text, selectedvalue, chartoRemove.Text));
                }
            };
            RbInt.CheckedChanged += (s, arg) => {
                int index = datahandle.FindIndex(k => k.DataIndex == button.Text);
                if (index != -1)
                {
                    datahandle[index] = (button.Text, RbInt.Text, chartoRemove.Text);
                }
                else
                {
                    datahandle.Add((button.Text, RbInt.Text, chartoRemove.Text));
                }

            };
            RbFloat.CheckedChanged += (s, arg) =>
            {
                int index = datahandle.FindIndex(k => k.DataIndex == button.Text);
                if (index != -1)
                {
                    datahandle[index] = (button.Text, RbFloat.Text, chartoRemove.Text);
                }
                else
                {
                    datahandle.Add((button.Text, RbFloat.Text, chartoRemove.Text));
                }
            };
            RbStr.CheckedChanged += (s, arg) =>
            {
                int index = datahandle.FindIndex(k => k.DataIndex == button.Text);
                if (index != -1)
                {
                    datahandle[index] = (button.Text, RbStr.Text, chartoRemove.Text);
                }
                else
                {
                    datahandle.Add((button.Text, RbStr.Text, chartoRemove.Text));
                }

            };
        }
        private void CloseBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void Setting_Load(object sender, EventArgs e)
        {
            FromName.Text = from;
            ToName.Text = to;
        }
        private void Setting_FormClosed(object sender, FormClosedEventArgs e)
        {

        }
        private void TextMeasure(Control btn, Control anchorX, int anchorY)
        {
            Size Textsize = new Size(TextRenderer.MeasureText(btn.Text, btn.Font).Width, TextRenderer.MeasureText(btn.Text, btn.Font).Height);
            btn.Size = Textsize;
            btn.Location = new System.Drawing.Point(anchorX.Bounds.Right+18, anchorY);
        }
        private void EditPreset()
        {
<<<<<<< HEAD

=======
>>>>>>> 0d5efcebdb9e9205644ac9bead6c701bfb0d828f
            int Highest = Preset.Select(t => int.Parse(t.DataIndex.Last().ToString())).Max();
            for (int i = 0; i < Highest; i++) {
                button3_Click(this,EventArgs.Empty);
            }
            foreach(var item in Preset)
            {
                for (int i = 0; i < groupBoxes.Count; i++) {
                    GroupBox gb = groupBoxes[i];
                    foreach (Control control in gb.Controls) { 
                        if(int.Parse(item.DataIndex.Last().ToString()) == i + 1){
                            if (control is Label Dlbl && Dlbl.Tag.ToString() == item.DataIndex)
                            {
                                Dlbl.Text = "Row : " + item.PresetRow + " \n " +
                                    "Col : " + item.PresetCol;
<<<<<<< HEAD
                                Debug.WriteLine("-------");
                                Debug.WriteLine(Dlbl.Tag);
                                Debug.WriteLine($"Row : { item.PresetRow}");
                                Debug.WriteLine($"Col : {item.PresetCol}");
=======
                                Debug.WriteLine("------- This Part is For Data");
                                Debug.WriteLine(Dlbl.Tag.ToString());
                                Debug.WriteLine("Current Group Box : " + (i + 1));
                                Debug.WriteLine(item);
                                Debug.WriteLine("-------");
>>>>>>> 0d5efcebdb9e9205644ac9bead6c701bfb0d828f
                            }
                        }

                        
                    }
                }
            }
            foreach(var data in datahandletoRtn)
            {
                for (int i = 0; i < groupBoxes.Count; i++)
                {
                    GroupBox gb = groupBoxes[i];
<<<<<<< HEAD
                    if (datahandletoRtn[i].DataIndex != i)
=======
                    if (datahandletoRtn[0].DataIndex-1 != i)
>>>>>>> 0d5efcebdb9e9205644ac9bead6c701bfb0d828f
                    {
                        foreach (Control control in gb.Controls)
                        {
                            if (control is RadioButton RB && RB.Text == $"{datahandletoRtn[i].handle}")
                            {
                                RB.Checked = true;
                            }
                            if (control is TextBox TB )
                            {
                                TB.Text = $"{datahandletoRtn[i].remove}"; 
                            }
                        }
                    }
                }
            }
        }
    }

}
