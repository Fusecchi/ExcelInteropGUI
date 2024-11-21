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
            this.Rollback = new System.Windows.Forms.Button();
            this.SaveLog = new System.Windows.Forms.Button();
            this.CloseEditLog = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // Rollback
            // 
            this.Rollback.Location = new System.Drawing.Point(670, 30);
            this.Rollback.Name = "Rollback";
            this.Rollback.Size = new System.Drawing.Size(160, 70);
            this.Rollback.TabIndex = 2;
            this.Rollback.Text = "Rollback";
            this.Rollback.UseVisualStyleBackColor = true;
            this.Rollback.Click += new System.EventHandler(this.Rollback_Click);
            // 
            // SaveLog
            // 
            this.SaveLog.Location = new System.Drawing.Point(670, 130);
            this.SaveLog.Name = "SaveLog";
            this.SaveLog.Size = new System.Drawing.Size(160, 70);
            this.SaveLog.TabIndex = 3;
            this.SaveLog.Text = "Save Log";
            this.SaveLog.UseVisualStyleBackColor = true;
            // 
            // CloseEditLog
            // 
            this.CloseEditLog.Location = new System.Drawing.Point(670, 230);
            this.CloseEditLog.Name = "CloseEditLog";
            this.CloseEditLog.Size = new System.Drawing.Size(160, 70);
            this.CloseEditLog.TabIndex = 4;
            this.CloseEditLog.Text = "Close";
            this.CloseEditLog.UseVisualStyleBackColor = true;
            this.CloseEditLog.Click += new System.EventHandler(this.CloseEditLog_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 16;
            this.listBox1.Location = new System.Drawing.Point(12, 8);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(622, 452);
            this.listBox1.TabIndex = 5;
            this.listBox1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.listBox1_MouseClick);
            // 
            // EditLog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(870, 515);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.CloseEditLog);
            this.Controls.Add(this.SaveLog);
            this.Controls.Add(this.Rollback);
            this.Name = "EditLog";
            this.Text = "EditLog";
            this.Load += new System.EventHandler(this.EditLog_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button Rollback;
        private System.Windows.Forms.Button SaveLog;
        private System.Windows.Forms.Button CloseEditLog;
        private System.Windows.Forms.ListBox listBox1;
    }
}