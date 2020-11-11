using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace CS_KPMCreator
{
    internal class WebControl_Selenium
    {
        private Util g_Util = null;
        private IWebDriver FF_driver = null;
        private WebAction WA = null;

        public WebControl_Selenium(ref Util util)
        {
            g_Util = util;
            WA = new WebAction(ref g_Util);
        }

        public void FindNewProcess(Process[] processesBefore, Process[] processesAfter, ref List<Process> processes)
        {
            foreach (Process AfterIter in processesAfter)
            {
                bool bFind = false;
                foreach (Process BeforeIter in processesBefore)
                {
                    if (AfterIter.Id == BeforeIter.Id)
                    {
                        bFind = true;
                        break;
                    }
                }
                if (bFind == false)
                {
                    processes.Add(AfterIter);
                }
            }
        }

        public void OpenWebSite(RadioButton rbB2B, RadioButton rbB2C, RadioButton rbKPMRead, TextBox tB2BID, TextBox tB2BPW, ref List<Process> processes)
        {
            g_Util.DebugPrint("I'm accessing KPM...");

            Process[] processesBefore = Process.GetProcessesByName("geckodriver");

            //var Dir = System.IO.Directory.GetCurrentDirectory();
            //FirefoxDriverService FF_driverSerivce = FirefoxDriverService.CreateDefaultService("C:\\KPM_Creator\\WebDriver");
            FirefoxDriverService FF_driverSerivce = FirefoxDriverService.CreateDefaultService();
            FF_driverSerivce.HideCommandPromptWindow = true;
            FirefoxOptions FF_options = new FirefoxOptions();
            FF_driver = new FirefoxDriver(FF_driverSerivce, FF_options);
            //WA.wait = new WebDriverWait(FF_driver, TimeSpan.FromMinutes(1));
            WA.wait = new WebDriverWait(FF_driver, TimeSpan.FromMilliseconds(100));

            Process[] processesAfter = Process.GetProcessesByName("geckodriver");
            FindNewProcess(processesBefore, processesAfter, ref processes);

            string URL = "";

            //rbB2B.Checked = true;
            //rbB2C.Checked = false;

            if (rbKPMRead.Checked == true || rbB2B.Checked == true)
            {
                if (tB2BID.Text != null && tB2BPW.Text != null)
                {
                    URL = "https://" + tB2BID.Text + ":" + tB2BPW.Text + "@sso.volkswagen.de/kpmweb/Index.action";
                }
                else
                {
                    URL = "https://sso.volkswagen.de/kpmweb/Index.action";
                }
            }
            else
            {
                URL = "https://quasi.vw.vwg/kpm/kpmweb/Index.action";
            }

            GoToURL(URL, rbB2C);
        }

        public void GoToURL(string URL, RadioButton rbB2C = null)
        {
            FF_driver.Navigate().GoToUrl(URL);
            WA.TotalWait(FF_driver);
        }

        public bool GoToMainPage(ref List<Dictionary<string, string>> TicketItemList, ref List<Dictionary<string, string>> LActionList, RadioButton rbB2C)
        {
            g_Util.DebugPrint("Please Login KPM... ");
            bool bResult = true;

            string ID = LActionList[0]["ID"];
            string SearchType = LActionList[0]["SearchType"];
            WA.TotalWait(FF_driver, ID, SearchType, 600000);    // Wait for 10 min to Login.
            WA.ManualWait(5000);

            Dictionary<string, string> TicketItem = TicketItemList[0];
            List<Dictionary<KPMReadInfo, List<string>>> DummyList = null;

            for (int nIdx = 0; nIdx < LActionList.Count; nIdx++)
            {
                Dictionary<string, string> ActionItem = LActionList[nIdx];
                if (ActionItem["Step"] == "StartEvent" && ActionItem["Execute"] != "X")
                {
                    if (CallAction(ref TicketItem, ref ActionItem, ref DummyList, rbB2C) == false)
                    {
                        bResult = false;
                        break;
                    }
                }
            }
            return bResult;
        }

        public bool CreateTickets(ref List<Dictionary<string, string>> TicketItemList, ref List<Dictionary<string, string>> LActionList, RadioButton rbB2C, ref int nTicketCNT)
        {
            bool bResult = true;

            g_Util.DebugPrint("I'm creating KPM Ticket...");

            //Access each ticket items
            for (int nIdx = 0; nIdx < TicketItemList.Count; nIdx++)
            //for (int nIdx = 0; nIdx < 1; nIdx++)
            {
                Dictionary<string, string> dItem = TicketItemList[nIdx];
                
                bool bCreateTicket = (dItem["Number"] == null) ? true : false;
                bool bUplaodData = (dItem["Re-upload Attachment"] != "X") ? true : false;
                bool bExecute = false;

                if (bCreateTicket == true)
                {
                    if (CreateOneTicket(ref dItem, ref LActionList, rbB2C) == false) // create ticket
                    {
                        bResult = false;
                        break;
                    }

                    bExecute = true;
                }

                bool bNum = true;
                try
                {
                    int temp = Int32.Parse(dItem["Number"]);
                    bNum = true;
                }
                catch
                {
                    bNum = false;
                }
                //bool bNum = int.Parse(dItem["Number"]).GetType().Equals(typeof(int));
                if (bUplaodData == true && dItem["Documents"] != null && bNum == true)
                {
                    bExecute = true;

                    if (GoToAttachmentPage(ref dItem, ref LActionList, rbB2C) == false)
                    {
                        bResult = false;
                        break;
                    }

                    if (UploadAttachment(ref dItem, ref LActionList, rbB2C) == false)
                    {
                        bResult = false;
                        break;
                    }
                }

                if (bExecute == true)
                {
                    nTicketCNT++;
                }

            }
            return bResult;
        }

        private bool CreateOneTicket(ref Dictionary<string, string> TicketItem, ref List<Dictionary<string, string>> LActionList, RadioButton rbB2C)
        {
            bool bResult = true;
            List<Dictionary<KPMReadInfo, List<string>>> DummyList = null;

            for (int nIdx = 0; nIdx < LActionList.Count; nIdx++)
            {
                Dictionary<string, string> ActionItem = LActionList[nIdx];
                if (ActionItem["Step"] == "CreateTicket" && ActionItem["Execute"] != "X")
                {
                    if (CallAction(ref TicketItem, ref ActionItem, ref DummyList, rbB2C) == false)
                    {
                        bResult = false;
                        break;
                    }
                }
            }
            return bResult;
        }

        private bool GoToAttachmentPage(ref Dictionary<string, string> TicketItem, ref List<Dictionary<string, string>> LActionList, RadioButton rbB2C)
        {
            bool bResult = true;
            List<Dictionary<KPMReadInfo, List<string>>> DummyList = null;

            for (int nIdx = 0; nIdx < LActionList.Count; nIdx++)
            {
                Dictionary<string, string> ActionItem = LActionList[nIdx];
                if (ActionItem["Step"] == "GoToAttach" && ActionItem["Execute"] != "X")
                {
                    if (CallAction(ref TicketItem, ref ActionItem, ref DummyList, rbB2C) == false)
                    {
                        bResult = false;
                        break;
                    }
                }
            }
            return bResult;
        }

        private bool UploadAttachment(ref Dictionary<string, string> TicketItem, ref List<Dictionary<string, string>> LActionList, RadioButton rbB2C)
        {
            bool bResult = true;

            string[] attachments = TicketItem["Documents"].Split('\n');
            string[] Comments = TicketItem["Doc Comment"].Split('\n');

            int nAttSize = attachments.Length;
            int nCommentSize = Comments.Length;

            List<Dictionary<KPMReadInfo, List<string>>> DummyList = null;

            if (nAttSize != nCommentSize)
            {
                TicketItem["Re-upload Attachment"] = "Attachment/Comment Unmatching!";
                bResult = false;
                return bResult;
            }

            for (int idx = 0; idx < nAttSize; idx++)
            {
                for (int nIdx = 0; nIdx < LActionList.Count; nIdx++)
                {
                    Dictionary<string, string> ActionItem = LActionList[nIdx];
                    if (ActionItem["Step"] == "UploadFiles" && ActionItem["Execute"] != "X")
                    {
                        string nText = null;
                        if (ActionItem["InputString"] == "Doc Comment")
                        {
                            nText = Comments[idx];
                        }
                        else if (ActionItem["InputString"] == "Documents")
                        {
                            nText = attachments[idx];
                        }

                        if (CallAction(ref TicketItem, ref ActionItem, ref DummyList, rbB2C, nText) == false)
                        {
                            bResult = false;
                            return bResult;
                        }
                    }
                }
            }

            TicketItem["Re-upload Attachment"] = "X";
            return bResult;
        }

        public bool KPMRead(ref List<Dictionary<string, string>> LActionList, ref ExcelControl g_ExcelTool)
        {
            bool bResult = true;

            g_Util.DebugPrint("I'm reading KPM Page...");

            List<Dictionary<KPMReadInfo, List<string>>> ReadList = new List<Dictionary<KPMReadInfo, List<string>>>();
            Dictionary<string, string> dummyitem = null;

            for (int nIdx = 0; nIdx < LActionList.Count; nIdx++)
            {
                Dictionary<string, string> ActionItem = LActionList[nIdx];
                if (ActionItem["Step"] == "ReadKPM" && ActionItem["Execute"] != "X")
                {
                    if (CallAction(ref dummyitem, ref ActionItem, ref ReadList, null, ActionItem["InputString"]) == false)
                    {
                        bResult = false;
                        break;
                    }
                    else
                    {
                        if (ActionItem["ActionType"] == "READ_DROPBOX")
                        {
                            g_ExcelTool.UpdateKPMReadSheet(ReadList);
                            ReadList.Clear();
                        }
                    }
                }
            }
            return bResult;
        }

        private bool CoordinatorCheck(ref Dictionary<string, string> TicketItem, ref Dictionary<string, string> ActionItem)
        {
            bool bResult = true;
            string TicketKey = "";
            if (ActionItem["Comment"] == "Coordinator")
            {
                TicketKey = "Coordinator user";
            }
            else if (ActionItem["Comment"] == "S.Coordinator")
            {
                TicketKey = "Spclst. coord User";
            }
            else if (ActionItem["Comment"] == "Solver")
            {
                TicketKey = "Problem Solver User";
            }

            if (TicketKey != "" && TicketItem[TicketKey] == null)
            {
                bResult = false;
            }

            return bResult;
        }

        private bool CallAction(ref Dictionary<string, string> TicketItem, ref Dictionary<string, string> ActionItem, ref List<Dictionary<KPMReadInfo, List<string>>> ReadList, RadioButton rbB2C, string TxtForUpload = null, string Read_Input_String = "")
        {
            bool bResult = true;

            if (CoordinatorCheck(ref TicketItem, ref ActionItem) == false)
            {
                return true;
            }
            
            g_Util.DebugPrint("\nI am doing... " + ActionItem["Comment"]);
            bool bSkipable = (ActionItem["Skipable"] == "O")? true: false;

            if (ActionItem["ActionType"] == "CLICK")
            {
                bResult = WA.ClickButton(FF_driver, ActionItem["ID"], ActionItem["SearchType"], bSkipable);
            }
            else if (ActionItem["ActionType"] == "JAVASCRIPT")
            {
                bResult = WA.ExecuteJS(FF_driver, ActionItem["ID"]);
            }
            else if (ActionItem["ActionType"] == "DROPBOX")
            {
                string TicketKey = ActionItem["InputString"];
                bResult = WA.SetComboItem(FF_driver, ActionItem["ID"], TicketItem[TicketKey]);
            }
            else if (ActionItem["ActionType"] == "DROPBOX_BYSELECT")
            {
                string TicketKey = ActionItem["InputString"];
                //bResult = WA.SetComboItemBySelect(FF_driver, ActionItem["ID"], TicketItem[TicketKey]);
            }
            else if (ActionItem["ActionType"] == "INPUT_TEXT")
            {
                string TicketKey = ActionItem["InputString"];
                bResult = WA.SetTextBox(FF_driver, ActionItem["ID"], TicketItem[TicketKey]);
            }
            else if (ActionItem["ActionType"] == "COPY_NUM")
            {
                TicketItem["Number"] = WA.GetText(FF_driver, ActionItem["ID"]);
            }
            else if (ActionItem["ActionType"] == "GOTOURL")
            {
                string URL = ActionItem["ID"] + TicketItem["Number"];
                if (rbB2C.Checked == true)
                {
                    URL += "&oid=";
                }  

                GoToURL(URL);
            }
            else if (ActionItem["ActionType"] == "INPUT_UPLOAD")
            {
                bResult = WA.SetTextBox(FF_driver, ActionItem["ID"], TxtForUpload);
            }
            else if (ActionItem["ActionType"] == "CALLEVENT")
            {
                bResult = WA.CallEvent(FF_driver, ActionItem["ID"], ActionItem["SubID"]);
            }
            else if (ActionItem["ActionType"] == "READ_DROPBOX")
            {
                //int nDepth = Convert.ToInt32(ActionItem["Depth"]);
                //bResult = WA.ReadComboItem(IE, doc,
                //                            "", "", "", "",
                //                            0, nDepth,
                //                            ActionItem["ID"], ActionItem["ID2"], ActionItem["ID3"], ActionItem["ID4"], ActionItem["Comment"],
                //                            ref ReadList);
            }
            else if (ActionItem["ActionType"] == "READ_INPUT_TEXT")
            {
                //bResult = WA.SetTextBox(IE, doc, ActionItem["ID"], Read_Input_String);
            }

            WA.TotalWait(FF_driver);

            if (ActionItem["ManualWait"] == "O")
            {
                WA.ManualWait(Int32.Parse(ActionItem["WaitTime"]));
            }


            return bResult;
        }
    }
}