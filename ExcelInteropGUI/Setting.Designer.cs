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
            // Setting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(849, 473);
            this.Controls.Add(this.AddData);
            this.Controls.Add(this.CloseBtn);
            this.Controls.Add(this.Save_Preset);
            this.Name = "Setting";
            this.Text = "Setting";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Setting_FormClosed);
            this.ResumeLayout(false);

        }

        private void s(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        #endregion
        private System.Windows.Forms.Button Save_Preset;
        private System.Windows.Forms.Button CloseBtn;
        private System.Windows.Forms.Button AddData;
    }
}