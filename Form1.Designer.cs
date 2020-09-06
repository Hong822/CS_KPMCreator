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
            this.rbB2B = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.tB2BPW = new System.Windows.Forms.TextBox();
            this.tB2BID = new System.Windows.Forms.TextBox();
            this.rbB2C = new System.Windows.Forms.RadioButton();
            this.bStartCreation = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rbChrome = new System.Windows.Forms.RadioButton();
            this.rbFirefox = new System.Windows.Forms.RadioButton();
            this.rbIE = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // bExcelSelect
            // 
            this.bExcelSelect.Location = new System.Drawing.Point(410, 34);
            this.bExcelSelect.Name = "bExcelSelect";
            this.bExcelSelect.Size = new System.Drawing.Size(110, 80);
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
            this.tExcelPath.Size = new System.Drawing.Size(350, 80);
            this.tExcelPath.TabIndex = 1;
            // 
            // ExcelOpenDialog
            // 
            this.ExcelOpenDialog.FileName = "ExcelOpenDialog";
            // 
            // rbB2B
            // 
            this.rbB2B.AutoSize = true;
            this.rbB2B.Location = new System.Drawing.Point(17, 30);
            this.rbB2B.Name = "rbB2B";
            this.rbB2B.Size = new System.Drawing.Size(45, 16);
            this.rbB2B.TabIndex = 2;
            this.rbB2B.Text = "B2B";
            this.rbB2B.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.textBox3);
            this.groupBox1.Controls.Add(this.tB2BPW);
            this.groupBox1.Controls.Add(this.tB2BID);
            this.groupBox1.Controls.Add(this.rbB2C);
            this.groupBox1.Controls.Add(this.rbB2B);
            this.groupBox1.Location = new System.Drawing.Point(41, 130);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(480, 91);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select KPM Type";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(86, 29);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(31, 21);
            this.textBox2.TabIndex = 8;
            this.textBox2.Text = "ID:";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(292, 29);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(31, 21);
            this.textBox3.TabIndex = 7;
            this.textBox3.Text = "PW:";
            // 
            // tB2BPW
            // 
            this.tB2BPW.Location = new System.Drawing.Point(329, 29);
            this.tB2BPW.Name = "tB2BPW";
            this.tB2BPW.PasswordChar = '*';
            this.tB2BPW.Size = new System.Drawing.Size(128, 21);
            this.tB2BPW.TabIndex = 5;
            this.tB2BPW.Text = "ta790909-1234";
            // 
            // tB2BID
            // 
            this.tB2BID.Location = new System.Drawing.Point(123, 29);
            this.tB2BID.Name = "tB2BID";
            this.tB2BID.Size = new System.Drawing.Size(154, 21);
            this.tB2BID.TabIndex = 4;
            this.tB2BID.Text = "dvkomiy";
            // 
            // rbB2C
            // 
            this.rbB2C.AutoSize = true;
            this.rbB2C.Checked = true;
            this.rbB2C.Location = new System.Drawing.Point(17, 66);
            this.rbB2C.Name = "rbB2C";
            this.rbB2C.Size = new System.Drawing.Size(46, 16);
            this.rbB2C.TabIndex = 3;
            this.rbB2C.TabStop = true;
            this.rbB2C.Text = "B2C";
            this.rbB2C.UseVisualStyleBackColor = true;
            // 
            // bStartCreation
            // 
            this.bStartCreation.Location = new System.Drawing.Point(40, 382);
            this.bStartCreation.Name = "bStartCreation";
            this.bStartCreation.Size = new System.Drawing.Size(479, 56);
            this.bStartCreation.TabIndex = 4;
            this.bStartCreation.Text = "Start Creation";
            this.bStartCreation.UseVisualStyleBackColor = true;
            this.bStartCreation.Click += new System.EventHandler(this.bStartCreation_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rbChrome);
            this.groupBox2.Controls.Add(this.rbFirefox);
            this.groupBox2.Controls.Add(this.rbIE);
            this.groupBox2.Location = new System.Drawing.Point(40, 232);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(480, 76);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Web Browser Type";
            // 
            // rbChrome
            // 
            this.rbChrome.AutoSize = true;
            this.rbChrome.Location = new System.Drawing.Point(293, 33);
            this.rbChrome.Name = "rbChrome";
            this.rbChrome.Size = new System.Drawing.Size(161, 16);
            this.rbChrome.TabIndex = 2;
            this.rbChrome.Text = "Chrome (Not supported)";
            this.rbChrome.UseVisualStyleBackColor = true;
            // 
            // rbFirefox
            // 
            this.rbFirefox.AutoSize = true;
            this.rbFirefox.Location = new System.Drawing.Point(176, 33);
            this.rbFirefox.Name = "rbFirefox";
            this.rbFirefox.Size = new System.Drawing.Size(61, 16);
            this.rbFirefox.TabIndex = 1;
            this.rbFirefox.Text = "Firefox";
            this.rbFirefox.UseVisualStyleBackColor = true;
            // 
            // rbIE
            // 
            this.rbIE.AutoSize = true;
            this.rbIE.Checked = true;
            this.rbIE.Location = new System.Drawing.Point(18, 33);
            this.rbIE.Name = "rbIE";
            this.rbIE.Size = new System.Drawing.Size(115, 16);
            this.rbIE.TabIndex = 0;
            this.rbIE.TabStop = true;
            this.rbIE.Text = "Internet Explorer";
            this.rbIE.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(555, 450);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.bStartCreation);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.tExcelPath);
            this.Controls.Add(this.bExcelSelect);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button bExcelSelect;
        private System.Windows.Forms.TextBox tExcelPath;
        private System.Windows.Forms.OpenFileDialog ExcelOpenDialog;
        private System.Windows.Forms.RadioButton rbB2B;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox tB2BPW;
        private System.Windows.Forms.TextBox tB2BID;
        private System.Windows.Forms.RadioButton rbB2C;
        private System.Windows.Forms.Button bStartCreation;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rbChrome;
        private System.Windows.Forms.RadioButton rbFirefox;
        private System.Windows.Forms.RadioButton rbIE;
    }
}

