using System.Windows.Forms;

namespace ExcelInteropGUI
{
    partial class LoadingBar
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
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.ProcessingLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.BackColor = System.Drawing.SystemColors.Control;
            this.progressBar1.Location = new System.Drawing.Point(12, 30);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(417, 45);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar1.TabIndex = 0;
            // 
            // ProcessingLabel
            // 
            this.ProcessingLabel.AutoSize = true;
            this.ProcessingLabel.Location = new System.Drawing.Point(18, 7);
            this.ProcessingLabel.Name = "ProcessingLabel";
            this.ProcessingLabel.Size = new System.Drawing.Size(94, 20);
            this.ProcessingLabel.TabIndex = 1;
            this.ProcessingLabel.Text = "Processing";
            // 
            // LoadingBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(458, 93);
            this.Controls.Add(this.ProcessingLabel);
            this.Controls.Add(this.progressBar1);
            this.Name = "LoadingBar";
            this.Text = "LoadingBar";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private Label ProcessingLabel;
    }
}