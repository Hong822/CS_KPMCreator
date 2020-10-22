using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;

namespace CS_KPMCreator
{
    public partial class Form1 : Form
    {
        private Util g_Util = null;
        private ExcelControl g_ExcelTool = null;
        private WebControl_SHDoc g_WebControl_SHDOC = null;
        private WebControl_Selenium g_WebControl_Selenium = null;
        private List<Process> processes = new List<Process>();

        public Form1()
        {
            InitializeComponent();
            this.FormClosing += Form1_FormClosing;
            this.Load += Form1_FormLoading;
            g_Util = new Util(ref richTB_Status);
        }

        private void bExcelSelect_Click(object sender, EventArgs e)
        {
            tExcelPath.Clear();
            ExcelOpenDialog.RestoreDirectory = false;
            ExcelOpenDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (ExcelOpenDialog.ShowDialog() == DialogResult.OK)
            {
                tExcelPath.Text = ExcelOpenDialog.FileName;

                //var Dir = Directory.GetCurrentDirectory();
                //Dir = Dir.Substring(0, Dir.LastIndexOf("\\"));
                //Dir = Dir.Substring(0, Dir.LastIndexOf("\\"));
                //tExcelPath.Text = Dir + "\\KPM_Ticket_Creator_V1.xlsm";
            }
        }

        private void bStartCreation_Click(object sender, EventArgs e)
        {
            try
            {
                List<Dictionary<string, string>> LTicketItemList = new List<Dictionary<string, string>>();
                List<Dictionary<string, string>> LActionList = new List<Dictionary<string, string>>();

                var nStartTick = DateTime.Now;

                bool bExcelReadResult = true;
                bool bCreateResult = false;
                int tryCnt = 0;

                if (rbKPMRead.Checked == true || rbTKCancel.Checked == true)
                {
                    rbB2B.Checked = true;
                    rbIE.Checked = true;
                }

                if (rbB2C.Checked == true)
                {
                    // Because of the error below, IE can't be used for B2C. (Use FF)
                    // The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)'
                    rbFirefox.Checked = true;
                }
                else
                {
                    rbIE.Checked = true;
                }

                g_ExcelTool = new ExcelControl(ref g_Util);
                if (g_ExcelTool.ReadExcelValue(tExcelPath, rbB2B, rbB2C, rbAudi, rbPorsche, rbKPMRead, rbTKCancel, ref LTicketItemList, ref LActionList, ref processes) == true)   // Data read from Excel Files
                {
                    if (rbIE.Checked == true)
                    {
                        g_WebControl_SHDOC = new WebControl_SHDoc(ref g_Util);

                        g_WebControl_SHDOC.OpenWebSite(rbB2B, rbB2C, rbKPMRead, tB2BID, tB2BPW, ref processes);  // Go to KPM site
                        g_WebControl_SHDOC.GoToMainPage(LActionList[0]);

                        if (rbKPMRead.Checked == true)
                        {
                            while (bCreateResult == false && tryCnt < 3)
                            {
                                bCreateResult = g_WebControl_SHDOC.KPMRead(ref LActionList, ref g_ExcelTool);   // KPM Read
                                tryCnt++;
                            }
                        }
                        else if (rbTKCancel.Checked == true)
                        {
                            bCreateResult = g_WebControl_SHDOC.Delete(ref LTicketItemList, ref LActionList, ref g_ExcelTool);   // Cancel Ticket
                        }
                        else
                        {
                            while (bCreateResult == false && tryCnt < 3)
                            {
                                bCreateResult = g_WebControl_SHDOC.CreateTickets(ref LTicketItemList, ref LActionList);   // Start Ticket Creation
                                tryCnt++;
                            }
                        }
                    }
                    else
                    {
                        g_WebControl_Selenium = new WebControl_Selenium(ref g_Util);
                        g_WebControl_Selenium.OpenWebSite(rbB2B, rbB2C, rbKPMRead, tB2BID, tB2BPW, ref processes);  // Go to KPM site
                        g_WebControl_Selenium.GoToMainPage(ref LTicketItemList, ref LActionList);

                        while (bCreateResult == false && tryCnt < 3)
                        {
                            bCreateResult = g_WebControl_Selenium.CreateTickets(ref LTicketItemList, ref LActionList);   // Start Ticket Creation
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
                    ResultReport = "[Abnormal Termination!] Something happen during creation. Please try it later. Try Count= " + tryCnt;
                }
                else
                {
                    ResultReport = "Creation Success. " + LTicketItemList.Count + " Tickets (" + nDiffSpan.Hours + "hr:" + nDiffSpan.Minutes + "min:" + nDiffSpan.Seconds + "sec). Try Count= " + tryCnt;
                }

                g_Util.DebugPrint(ResultReport);
            }
            catch (Exception error)
            {
                g_Util.DebugPrint(error.ToString());
                CloseAll();
            }
        }

        private void Form1_FormLoading(Object sender, EventArgs e)
        {
            //Properties.Settings.Default.Reload();
        }

        private void Form1_Closing(Object sender, CancelEventArgs e)
        {
            CloseAll();
        }

        private void Form1_Closed(Object sender, EventArgs e)
        {
            CloseAll();
        }

        private void Form1_FormClosed(Object sender, FormClosedEventArgs e)
        {
            CloseAll();
        }

        private void Form1_FormClosing(Object sender, FormClosingEventArgs e)
        {
            CloseAll();
        }

        private void CloseAll()
        {
            if (g_ExcelTool != null)
            {
                g_ExcelTool.CloseExcelControl(ref processes);
            }

            if (processes != null)
            {
                foreach (Process Iter in processes)
                {
                    System.Diagnostics.Debug.WriteLine("Kill " + Iter.ProcessName);
                    try
                    {
                        Process.GetProcessById(Iter.Id);
                        Iter.Kill();
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
        }
    }
}