namespace ExcelInteropGUI
{
    partial class EditLog
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
            this.ShowLogText = new System.Windows.Forms.RichTextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.Rollback = new System.Windows.Forms.Button();
            this.SaveLog = new System.Windows.Forms.Button();
            this.CloseEditLog = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ShowLogText
            // 
            this.ShowLogText.Location = new System.Drawing.Point(11, 8);
            this.ShowLogText.Name = "ShowLogText";
            this.ShowLogText.Size = new System.Drawing.Size(551, 357);
            this.ShowLogText.TabIndex = 0;
            this.ShowLogText.Text = "";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(14, 385);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(355, 24);
            this.comboBox1.TabIndex = 1;
            // 
            // Rollback
            // 
            this.Rollback.Location = new System.Drawing.Point(403, 385);
            this.Rollback.Name = "Rollback";
            this.Rollback.Size = new System.Drawing.Size(119, 24);
            this.Rollback.TabIndex = 2;
            this.Rollback.Text = "Rollback";
            this.Rollback.UseVisualStyleBackColor = true;
            // 
            // SaveLog
            // 
            this.SaveLog.Location = new System.Drawing.Point(601, 8);
            this.SaveLog.Name = "SaveLog";
            this.SaveLog.Size = new System.Drawing.Size(167, 66);
            this.SaveLog.TabIndex = 3;
            this.SaveLog.Text = "Save Log";
            this.SaveLog.UseVisualStyleBackColor = true;
            // 
            // CloseEditLog
            // 
            this.CloseEditLog.Location = new System.Drawing.Point(601, 105);
            this.CloseEditLog.Name = "CloseEditLog";
            this.CloseEditLog.Size = new System.Drawing.Size(167, 66);
            this.CloseEditLog.TabIndex = 4;
            this.CloseEditLog.Text = "Close";
            this.CloseEditLog.UseVisualStyleBackColor = true;
            // 
            // EditLog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.CloseEditLog);
            this.Controls.Add(this.SaveLog);
            this.Controls.Add(this.Rollback);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.ShowLogText);
            this.Name = "EditLog";
            this.Text = "EditLog";
            this.Load += new System.EventHandler(this.EditLog_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox ShowLogText;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button Rollback;
        private System.Windows.Forms.Button SaveLog;
        private System.Windows.Forms.Button CloseEditLog;
    }
}