using DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelInteropGUI
{
    public partial class Setting : Form

    {
        public System.Data.DataTable DataTable { get; set; } = new System.Data.DataTable();
        public List<(string setting, int PresetRow, int PresetCol)> Preset { get; set; } = new List<(string setting, int PresetRow, int PresetCol)>();
        private int BtnIterration = 0;
        private Panel scrollPanel = new Panel();
        //private int yoffset = 20;
        public Setting()
        {
            InitializeComponent();
            scrollPanel.AutoScroll = true;  // Enable scrolling
            scrollPanel.Dock = DockStyle.Fill;  // Dock the panel to fill the form
            this.Controls.Add(scrollPanel);
        }
        

        private void button1_Click(object sender, EventArgs e)
        {
            string json =   JsonConvert.SerializeObject(Preset, Formatting.Indented);
            File.WriteAllText("Preset.json", json);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            BtnIterration++;
            Button button = new Button();
            Label label = new Label();
            label.Location = new System.Drawing.Point(200, 100 * BtnIterration);
            button.Text = $"Data{BtnIterration}";
            button.Size = new Size(100, 100);
            button.Location = new System.Drawing.Point(100, 100 * BtnIterration);
            button.Click += (s, args) =>
            {
                SelectDataForm select = new SelectDataForm();
                select.DatatoClick = DataTable;
                select.FormClosed += (c, arg) => this.Show();
                select.selectedData += tuple =>
                {
                    Preset.Add((tuple.CellVal.ToString(), tuple.rowInd,tuple.ColInd));
                    label.Text = tuple.CellVal.ToString();
                };
                select.Show();
                this.Hide();
            };
            //Debug.WriteLine(yoffset);
            scrollPanel.Controls.Add(button);
            scrollPanel.Controls.Add(label);
            //yoffset += 100;
        }

        private void CloseBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Setting_FormClosed(object sender, FormClosedEventArgs e)
        {

        }
    }
}
