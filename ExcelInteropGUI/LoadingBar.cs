﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelInteropGUI
{
    public partial class LoadingBar : Form
    {
        public Action Worker { get; set; }
        private Menu _form1;

        public LoadingBar(Action worker, Menu form1)
        {
            InitializeComponent();
            if (worker == null)
                throw new ArgumentNullException();
            Worker = worker;
            _form1 = form1;
            this.MaximizeBox = false;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Task.Factory.StartNew(Worker).ContinueWith(task => { this.Close(); }, TaskScheduler.FromCurrentSynchronizationContext());
            if (_form1 != null)
            {
                _form1.OnFunctionStart += Form1OnFunctionCompleted;
            }
        }

        private void Form1OnFunctionCompleted(string Update)
        {
            if (ProcessingLabel.InvokeRequired)
            {
                ProcessingLabel.Invoke(new Action(() =>
                {
                    ProcessingLabel.Text = Update;
                }));
            }
            else
            {
                ProcessingLabel.Text = Update;
            }
        }
    }


}
