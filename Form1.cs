﻿using System;
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
            rbKPMRead.Checked = true;

            List<Dictionary<string, string>> LTicketItemList = new List<Dictionary<string, string>>();
            List<Dictionary<string, string>> LActionList = new List<Dictionary<string, string>>();

            var nStartTick = DateTime.Now;

            bool bExcelReadResult = true;
            bool bCreateResult = false;
            int tryCnt = 0;
            if (g_ExcelTool.ReadExcelValue(tExcelPath, rbB2B, rbB2C, rbAudi, rbPorsche, rbKPMRead, ref LTicketItemList, ref LActionList) == true)   // Data read from Excel Files
            {
                if (rbIE.Checked == true)
                {
                    g_WebControl = new WebControl_SHDoc();
                }
                else
                {
                    g_WebControl = new WebControl_SHDoc();
                }
                g_WebControl.SetStatusBox(ref richTB_Status);

                g_WebControl.OpenWebSite(rbB2B, rbB2C, rbKPMRead, tB2BID, tB2BPW);  // Go to KPM site
                g_WebControl.GoToMainPage(LActionList[0]);

                if (rbKPMRead.Checked == true)                {
                    

                    while (bCreateResult == false && tryCnt < 3)
                    {
                        bCreateResult = g_WebControl.KPMRead(ref LActionList, ref g_ExcelTool);   // Start Ticket Creation
                        tryCnt++;
                    }

                }
                else
                {
                    while (bCreateResult == false && tryCnt < 3)
                    {
                        bCreateResult = g_WebControl.CreateTickets(ref LTicketItemList, ref LActionList);   // Start Ticket Creation
                        tryCnt++;
                    }
                }

                g_ExcelTool.UpdateKPMDocument(LTicketItemList);
            }
            else
            {
                bExcelReadResult = false;
            }
            var nEndTick = DateTime.Now;
            long nGap = nEndTick.Ticks - nStartTick.Ticks;
            var nDiffSpan = new TimeSpan(nGap);

            string ResultReport = "";
            if (bExcelReadResult == false)
            {
                ResultReport = "[Abnormal Termination!] Excel Path is Strange or ReadOnly. Please check your Excel File.";
            }
            else if (bCreateResult == false)
            {
                ResultReport = "[Abnormal Termination!] Something happen during creation. Please try it later. Try Count= "+ tryCnt;
            }
            else
            {
                ResultReport = "Creation Success. " + LTicketItemList.Count + " Tickets (" + nDiffSpan.Hours + "hr:" + nDiffSpan.Minutes + "min:" + nDiffSpan.Seconds + "sec). Try Count= " + tryCnt;
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