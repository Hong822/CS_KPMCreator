using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;


namespace CS_KPMCreator
{
    public partial class Form1 : Form
    {
        private ExcelControl g_ExcelTool = new ExcelControl();
        private WebControl g_WebControl = new WebControl();

        public Form1()
        {
            InitializeComponent();
            g_ExcelTool.SetStatusBox(ref richTB_Status);
            g_WebControl.SetStatusBox(ref richTB_Status);
        }

        private void bExcelSelect_Click(object sender, EventArgs e)
        {
            tExcelPath.Clear();
            ExcelOpenDialog.RestoreDirectory = false;
            ExcelOpenDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            //if (ExcelOpenDialog.ShowDialog() == DialogResult.OK)
            {
                var Dir = Directory.GetCurrentDirectory();
                Dir = Dir.Substring(0, Dir.LastIndexOf("\\"));
                Dir = Dir.Substring(0, Dir.LastIndexOf("\\"));

                //tExcelPath.Text = ExcelOpenDialog.FileName;
                tExcelPath.Text = Dir + "\\KPM_Ticket_Creator_V1.xlsm";
            }
        }

        private void bStartCreation_Click(object sender, EventArgs e)
        {
            List<Dictionary<string, string>> LTicketItemList = new List<Dictionary<string, string>>();
            List<Dictionary<string, string>> LActionList = new List<Dictionary<string, string>>();

            g_ExcelTool.ReadExcelValue(tExcelPath, rbB2B, rbB2C, rbAudi, rbPorsche, ref LTicketItemList, ref LActionList);   // Date read from Excel Files

            richTB_Status.Text = "I'm setting browser type...";
            g_WebControl.SetBrowser(rbFirefox, rbChrome);   // Get Browser

            richTB_Status.Text = "I'm accessing KPM...";
            g_WebControl.GoToTheSite(rbB2B, rbB2C, tB2BID, tB2BPW);  // Go to KPM site

            richTB_Status.Text = "I'm creating KPM Ticket...";
            g_WebControl.CreateTickets(LTicketItemList);   // Start Ticket Creation
        }
    }
}