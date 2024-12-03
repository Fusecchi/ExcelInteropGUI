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
        private int BtnIterration = 0;
        private Panel scrollPanel = new Panel();
        public string from, to;
        //private int yoffset = 20;
        public Setting()
        {
            InitializeComponent();
        }
        

        private void button1_Click(object sender, EventArgs e)
        {
            string json =   JsonConvert.SerializeObject(Preset, Formatting.Indented);
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
            Button button = new Button { 
                Text = Text = $"Data {BtnIterration}",
                Size = new Size(100, 30),
                Location = new System.Drawing.Point(100, 30 * BtnIterration)
            };
            Button Lbutton = new Button{
                Text = $"Target File {BtnIterration}",
                Size = new Size(100, 30),
                Location = new System.Drawing.Point(300, 30 * BtnIterration)
            };
            Label label = new Label
            {
                Location = new System.Drawing.Point(200, 30 * BtnIterration)
            };
            Label Llabel = new Label
            {
                Location = new System.Drawing.Point(400, 30 * BtnIterration)
            };
            scrollPanel.Controls.Add(button);
            scrollPanel.Controls.Add(label);
            scrollPanel.Controls.Add(Lbutton);
            scrollPanel.Controls.Add(Llabel);
            scrollPanel.AutoScrollMinSize = new Size(Width, (30 + 50) * BtnIterration);
            Debug.WriteLine($"Location of the Button {Lbutton.Location.X}, {Lbutton.Location.Y}");
            button.Click += (s, args) =>
            {
                SelectDataForm select = new SelectDataForm();
                select.DatatoClick = DataTable;
                select.FormClosed += (c, arg) => this.Show();
                select.selectedData += tuple =>
                {
                    Preset.Add((button.Text,tuple.CellVal.ToString(), tuple.rowInd+1,tuple.ColInd+1));
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
                    Preset.Add((Lbutton.Text,tuple.CellVal.ToString(), tuple.rowInd+1, tuple.ColInd + 1));
                    Llabel.Text = tuple.CellVal.ToString();
                };
                TSelect.Show();
                this.Hide();
            };
            //Debug.WriteLine(yoffset);

            //yoffset += 100;
        }

        private void CloseBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Setting_Load(object sender, EventArgs e)
        {
            scrollPanel.AutoScroll = true;  // Enable scrolling
            scrollPanel.Dock = DockStyle.Fill;  // Dock the panel to fill the form
            this.Controls.Add(scrollPanel);
            FromName.Text = from;
            ToName.Text = to;
        }

        private void Setting_FormClosed(object sender, FormClosedEventArgs e)
        {

        }
    }
}
