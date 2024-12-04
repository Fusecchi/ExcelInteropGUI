namespace ExcelInteropGUI
{
    partial class Setting
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Save_Preset = new System.Windows.Forms.Button();
            this.CloseBtn = new System.Windows.Forms.Button();
            this.AddData = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.FromLabel = new System.Windows.Forms.Label();
            this.ToLabel = new System.Windows.Forms.Label();
            this.FromName = new System.Windows.Forms.Label();
            this.ToName = new System.Windows.Forms.Label();
            this.scrollPanel = new System.Windows.Forms.Panel();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // Save_Preset
            // 
            this.Save_Preset.Location = new System.Drawing.Point(539, 396);
            this.Save_Preset.Name = "Save_Preset";
            this.Save_Preset.Size = new System.Drawing.Size(108, 42);
            this.Save_Preset.TabIndex = 1;
            this.Save_Preset.Text = "Save Preset";
            this.Save_Preset.UseVisualStyleBackColor = true;
            this.Save_Preset.Click += new System.EventHandler(this.button1_Click);
            // 
            // CloseBtn
            // 
            this.CloseBtn.Location = new System.Drawing.Point(680, 396);
            this.CloseBtn.Name = "CloseBtn";
            this.CloseBtn.Size = new System.Drawing.Size(108, 42);
            this.CloseBtn.TabIndex = 2;
            this.CloseBtn.Text = "Close ";
            this.CloseBtn.UseVisualStyleBackColor = true;
            this.CloseBtn.Click += new System.EventHandler(this.CloseBtn_Click);
            // 
            // AddData
            // 
            this.AddData.Location = new System.Drawing.Point(12, 12);
            this.AddData.Name = "AddData";
            this.AddData.Size = new System.Drawing.Size(108, 42);
            this.AddData.TabIndex = 3;
            this.AddData.Text = "Add Data";
            this.AddData.UseVisualStyleBackColor = true;
            this.AddData.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(314, 406);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(194, 22);
            this.textBox1.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(235, 408);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 16);
            this.label1.TabIndex = 5;
            this.label1.Text = "Save As";
            // 
            // FromLabel
            // 
            this.FromLabel.AutoSize = true;
            this.FromLabel.Location = new System.Drawing.Point(158, 25);
            this.FromLabel.Name = "FromLabel";
            this.FromLabel.Size = new System.Drawing.Size(44, 16);
            this.FromLabel.TabIndex = 6;
            this.FromLabel.Text = "From :";
            // 
            // ToLabel
            // 
            this.ToLabel.AutoSize = true;
            this.ToLabel.Location = new System.Drawing.Point(556, 25);
            this.ToLabel.Name = "ToLabel";
            this.ToLabel.Size = new System.Drawing.Size(24, 16);
            this.ToLabel.TabIndex = 7;
            this.ToLabel.Text = "To";
            // 
            // FromName
            // 
            this.FromName.AutoSize = true;
            this.FromName.Location = new System.Drawing.Point(208, 25);
            this.FromName.Name = "FromName";
            this.FromName.Size = new System.Drawing.Size(75, 16);
            this.FromName.TabIndex = 8;
            this.FromName.Text = "FromName";
            // 
            // ToName
            // 
            this.ToName.AutoSize = true;
            this.ToName.Location = new System.Drawing.Point(586, 25);
            this.ToName.Name = "ToName";
            this.ToName.Size = new System.Drawing.Size(61, 16);
            this.ToName.TabIndex = 9;
            this.ToName.Text = "ToName";
            // 
            // scrollPanel
            // 
            this.scrollPanel.AutoScroll = true;
            this.scrollPanel.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.scrollPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.scrollPanel.Location = new System.Drawing.Point(95, 72);
            this.scrollPanel.Name = "scrollPanel";
            this.scrollPanel.Size = new System.Drawing.Size(693, 308);
            this.scrollPanel.TabIndex = 10;
            // 
            // Setting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(878, 504);
            this.Controls.Add(this.scrollPanel);
            this.Controls.Add(this.ToName);
            this.Controls.Add(this.FromName);
            this.Controls.Add(this.ToLabel);
            this.Controls.Add(this.FromLabel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.AddData);
            this.Controls.Add(this.CloseBtn);
            this.Controls.Add(this.Save_Preset);
            this.Name = "Setting";
            this.Text = "Setting";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Setting_FormClosed);
            this.Load += new System.EventHandler(this.Setting_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void s(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        #endregion
        private System.Windows.Forms.Button Save_Preset;
        private System.Windows.Forms.Button CloseBtn;
        private System.Windows.Forms.Button AddData;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label FromLabel;
        private System.Windows.Forms.Label ToLabel;
        private System.Windows.Forms.Label FromName;
        private System.Windows.Forms.Label ToName;
        private System.Windows.Forms.Panel scrollPanel;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }
}