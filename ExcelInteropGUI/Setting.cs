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
        private int BtnIterration = 0;
        public string from, to;
        
        public Setting()
        {
            InitializeComponent();
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
        }

        private void button3_Click(object sender, EventArgs e)
        {
            BtnIterration++;
            //Button Property
            int buttonHeight = 30;
            int buttonWidth = 100;
            int buttonPosX = 15;
            //Lbutton Property
            int LbuttonHeight = 30;
            int LbuttonWidth = 100;

            int GapX = 15;
            int ypos = (buttonHeight + 50) * (BtnIterration - 1) + scrollPanel.AutoScrollPosition.Y;
            scrollPanel.AutoScrollMinSize = new Size(Width, ypos);
            //Initiate Button
            Button button = new Button {
                Text = Text = $"Data {BtnIterration}",
                Size = new Size(buttonWidth, buttonHeight),
                Location = new System.Drawing.Point(buttonPosX, ypos)
            };
            //Initiate Label
            Label label = new Label
            {
                Location = new System.Drawing.Point(buttonPosX + button.Width + GapX, ypos),
                Size = new Size(50, 30)
            };
            GroupBox groupBox = new GroupBox
            {
                Location = new System.Drawing.Point(label.Location.X, label.Location.Y + label.Height),
                AutoSize =true
            };

            //Initiate Radio Button
            RadioButton RbInt = new RadioButton
            {
                Location = new System.Drawing.Point(0,0),
                Font = new Font("Microsoft Sans Serif",8),
                Text = "Number",
                AutoSize = true,
            };
            Size Textsize = new Size(TextRenderer.MeasureText(RbInt.Text, RbInt.Font).Width, TextRenderer.MeasureText(RbInt.Text, RbInt.Font).Height + 5);
            RbInt.Size = Textsize;
            RadioButton RbStr = new RadioButton
            {
                Location = new System.Drawing.Point(RbInt.Location.X + RbInt.Width, RbInt.Location.Y),
                Text = "Text",
                Font = new Font("Microsoft Sans Serif", 8),
                AutoSize = true,
            };
            RadioButton RbFloat = new RadioButton
            {
                Location = new System.Drawing.Point(RbStr.Location.X + RbStr.Width, RbInt.Location.Y),
                Text = "Decimal",
                Font = new Font("Microsoft Sans Serif", 8),
                AutoSize = true,
            };
            groupBox.Size = new Size(RbFloat.Bounds.Right - label.Bounds.Left, RbInt.Size.Height);
            TextMeasure(RbFloat, RbInt);
            TextMeasure(RbStr, RbFloat);
            TextBox chartoRemove = new TextBox {
                Size = new Size(groupBox.Size.Width, label.Size.Height),
                Location = new System.Drawing.Point(label.Location.X + label.Width, label.Location.Y),
                Visible = true
            };

            Button Lbutton = new Button
            {
                Text = $"Target File {BtnIterration}",
                Size = new Size(LbuttonWidth, LbuttonHeight),
                Location = new System.Drawing.Point(chartoRemove.Location.X+chartoRemove.Width + GapX, ypos)
            };
            Label Llabel = new Label
            {
                Location = new System.Drawing.Point(Lbutton.Width + Lbutton.Location.X + GapX, ypos),
                Size = new Size(50, 30)
            };
            
            scrollPanel.Controls.Add(button);
            scrollPanel.Controls.Add(label);
            scrollPanel.Controls.Add(Lbutton);
            scrollPanel.Controls.Add(Llabel);
            groupBox.Controls.Add(RbInt);
            groupBox.Controls.Add(RbStr);
            groupBox.Controls.Add(RbFloat);
            scrollPanel.Controls.Add(groupBox);
            scrollPanel.Controls.Add(chartoRemove);

            button.Click += (s, args) =>
            {
                SelectDataForm select = new SelectDataForm();
                select.DatatoClick = DataTable;
                select.FormClosed += (c, arg) => this.Show();
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
                TSelect.FormClosed += (D, arg) => this.Show();
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
        private void TextMeasure(Control btn, Control anchorbtn)
        {
            Size Textsize = new Size(TextRenderer.MeasureText(btn.Text, btn.Font).Width+5, TextRenderer.MeasureText(btn.Text, btn.Font).Height+5);
            btn.Size = Textsize;
            btn.Location = new System.Drawing.Point(anchorbtn.Location.X + anchorbtn.Width +(btn.Width/2), anchorbtn.Location.Y);
        }
    }
}
