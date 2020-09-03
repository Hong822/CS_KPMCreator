using Microsoft.Office.Interop.Excel;

namespace CS_KPMCreator
{
    partial class Form1
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
            this.bExcelSelect = new System.Windows.Forms.Button();
            this.tExcelPath = new System.Windows.Forms.TextBox();
            this.ExcelOpenDialog = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // bExcelSelect
            // 
            this.bExcelSelect.Location = new System.Drawing.Point(552, 34);
            this.bExcelSelect.Name = "bExcelSelect";
            this.bExcelSelect.Size = new System.Drawing.Size(184, 80);
            this.bExcelSelect.TabIndex = 0;
            this.bExcelSelect.Text = "Select Excel File";
            this.bExcelSelect.UseVisualStyleBackColor = true;
            this.bExcelSelect.Click += new System.EventHandler(this.bExcelSelect_Click);
            // 
            // tExcelPath
            // 
            this.tExcelPath.Location = new System.Drawing.Point(41, 34);
            this.tExcelPath.Multiline = true;
            this.tExcelPath.Name = "tExcelPath";
            this.tExcelPath.Size = new System.Drawing.Size(481, 80);
            this.tExcelPath.TabIndex = 1;
            // 
            // ExcelOpenDialog
            // 
            this.ExcelOpenDialog.FileName = "ExcelOpenDialog";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.tExcelPath);
            this.Controls.Add(this.bExcelSelect);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button bExcelSelect;
        private System.Windows.Forms.TextBox tExcelPath;
        private System.Windows.Forms.OpenFileDialog ExcelOpenDialog;        
    }
}

