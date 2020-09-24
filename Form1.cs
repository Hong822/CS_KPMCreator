using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace CS_KPMCreator
{
    public partial class Form1 : Form
    {
        private ExcelControl g_ExcelTool = new ExcelControl();
        private WebControl_SHDoc g_WebControl = null;

        public Form1()
        {
            InitializeComponent();
            this.FormClosing += Form1_FormClosing;
            g_ExcelTool.SetStatusBox(ref richTB_Status);
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

            var nStartTick = DateTime.Now;

            g_ExcelTool.ReadExcelValue(tExcelPath, rbB2B, rbB2C, rbAudi, rbPorsche, ref LTicketItemList, ref LActionList);   // Date read from Excel Files

            if (rbIE.Checked == true)
            {
                g_WebControl = new WebControl_SHDoc();
            }
            else
            {
                g_WebControl = new WebControl_SHDoc();
            }
            g_WebControl.SetStatusBox(ref richTB_Status);

            g_WebControl.OpenWebSite(rbB2B, rbB2C, tB2BID, tB2BPW);  // Go to KPM site
            g_WebControl.GoToMainPage(LActionList[0]);
            bool bCreateResult = false;
            int tryCnt = 0;
            while (bCreateResult == false && tryCnt < 3)
            {
                bCreateResult = g_WebControl.CreateTickets(ref LTicketItemList, ref LActionList);   // Start Ticket Creation
                tryCnt++;
            }

            g_ExcelTool.UpdateKPMDocument(LTicketItemList);

            var nEndTick = DateTime.Now;
            long nGap = nEndTick.Ticks - nStartTick.Ticks;
            var nDiffSpan = new TimeSpan(nGap);

            string ResultReport = "";
            if (bCreateResult == false)
            {
                ResultReport = "Something happen. Please try it later";
            }
            else
            {
                ResultReport = "Finish.  " + LTicketItemList.Count + "Tickets (" + nDiffSpan.Hours + "hr:" + nDiffSpan.Minutes + "min:" + nDiffSpan.Seconds + "sec)";
            }

            richTB_Status.Text = ResultReport;
            System.Diagnostics.Debug.WriteLine(ResultReport);
        }

        private void Form1_FormClosing(Object sender, FormClosingEventArgs e)
        {
            g_ExcelTool.CloseExcelControl();

            //if (e.CloseReason == CloseReason.UserClosing)
            //{
            //    int ete = 2;
            //}
        }
    }
}