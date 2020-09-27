using mshtml;
using SHDocVw;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CS_KPMCreator
{
    internal class WebControl_SHDoc
    {
        private RichTextBox g_richTB_Status = null;
        private InternetExplorer IE = null;
        private IWebBrowserApp webBrowser = null;
        private HTMLDocument doc = null;

        private WebAction WA = new WebAction();

        public void SetStatusBox(ref RichTextBox richTB_Status)
        {
            g_richTB_Status = richTB_Status;
        }

        public void DebugPrint(string sDebugString)
        {
            g_richTB_Status.Text = sDebugString;
            System.Diagnostics.Debug.WriteLine(sDebugString);
        }

        public void OpenWebSite(RadioButton rbB2B, RadioButton rbB2C, TextBox tB2BID, TextBox tB2BPW)
        {
            DebugPrint("I'm accessing KPM...");
            WA.SetStatusBox(ref g_richTB_Status);

            IE = new InternetExplorer();
            webBrowser = (IWebBrowserApp)IE;
            object URL = null;

            rbB2B.Checked = true;
            rbB2C.Checked = false;
            if (rbB2C.Checked == true)
            {
                URL = "https://quasi.vw.vwg/kpm/kpmweb";
            }
            else
            {
                if (rbB2B.Checked == true)
                {
                    //if (sID != null && sPW != null)
                    //{
                    //    URL = "https://" + sID + ":" + sPW + "@sso.volkswagen.de/kpmweb/Index.action";
                    //}
                    //else
                    {
                        URL = "https://sso.volkswagen.de/kpmweb/Index.action";
                    }
                }
                else
                {
                    MessageBox.Show("This case should happen.");
                }
            }

            if (rbB2C.Checked == true)
            {
                MessageBox.Show("Wait!!! Login and Go to Main page. And then press OK", "Please Login KPM");
            }

            GoToURL(URL);
        }

        public void GoToURL(object URL)
        {
            IE.Visible = true;
            IE.Navigate2(ref URL);
            WA.WaitPageLoading(IE);
            doc = (HTMLDocument)IE.Document;
        }

        public bool GoToMainPage(Dictionary<string, string> Item)
        {
            //IHTMLElement SelectedElement = doc.getElementById(ID);
            //if ()
            string ID = Item["ID"];
            return WA.ClickButton(IE, doc, ID);
        }

        public bool CreateTickets(ref List<Dictionary<string, string>> TicketItemList, ref List<Dictionary<string, string>> LActionList)
        {
            bool bResult = true;

            DebugPrint("I'm creating KPM Ticket...");

            //Access each ticket items
            for (int nIdx = 0; nIdx < TicketItemList.Count; nIdx++)
            //for (int nIdx = 0; nIdx < 1; nIdx++)
            {
                Dictionary<string, string> dItem = TicketItemList[nIdx];
                if (dItem["Number"] == null)
                {
                    if (CreateOneTicket(ref dItem, ref LActionList) == false) // create ticket
                    {
                        bResult = false;
                        break;
                    }
                }

                if ( (dItem["Re-upload Attachment"] == "O" || dItem["Re-upload Attachment"] == null ) && dItem["Documents"] != null)
                {
                    if (GoToAttachmentPage(ref dItem, ref LActionList) == false)
                    {
                        bResult = false;
                        break;
                    }

                    if (UploadAttachment(ref dItem, ref LActionList) == false)
                    {
                        bResult = false;
                        break;
                    }
                }
            }
            return bResult;
        }

        private bool CreateOneTicket(ref Dictionary<string, string> TicketItem, ref List<Dictionary<string, string>> LActionList)
        {
            bool bResult = true;
            for (int nIdx = 0; nIdx < LActionList.Count; nIdx++)
            {
                Dictionary<string, string> ActionItem = LActionList[nIdx];
                if (ActionItem["Step"] == "CreateTicket" && ActionItem["Execute"] != "X")
                {
                    if (CallAction(ref TicketItem, ref ActionItem) == false)
                    {
                        bResult = false;
                        break;
                    }
                }
            }
            return bResult;
        }

        private bool GoToAttachmentPage(ref Dictionary<string, string> TicketItem, ref List<Dictionary<string, string>> LActionList)
        {
            bool bResult = true;
            for (int nIdx = 0; nIdx < LActionList.Count; nIdx++)
            {
                Dictionary<string, string> ActionItem = LActionList[nIdx];
                if (ActionItem["Step"] == "GoToAttach" && ActionItem["Execute"] != "X")
                {
                    if (CallAction(ref TicketItem, ref ActionItem) == false)
                    {
                        bResult = false;
                        break;
                    }
                }
            }
            return bResult;
        }

        private bool UploadAttachment(ref Dictionary<string, string> TicketItem, ref List<Dictionary<string, string>> LActionList)
        {
            bool bResult = true;

            string[] attachments = TicketItem["Documents"].Split('\n');
            string[] Comments = TicketItem["Doc Comment"].Split('\n');

            int nAttSize = attachments.Length;
            int nCommentSize = Comments.Length;

            if(nAttSize != nCommentSize)
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

                        if (CallAction(ref TicketItem, ref ActionItem, nText) == false)
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

        private bool CallAction(ref Dictionary<string, string> TicketItem, ref Dictionary<string, string> ActionItem, string TxtForUpload = null)
        {
            bool bResult = true;

            DebugPrint("I am doing... " + ActionItem["Comment"]);

            if (ActionItem["ActionType"] == "CLICK")
            {
                bResult = WA.ClickButton(IE, doc, ActionItem["ID"], ActionItem["SearchType"]);
            }
            else if (ActionItem["ActionType"] == "JAVASCRIPT")
            {
                bResult = WA.ExecuteJS(IE, doc, ActionItem["ID"]);
            }
            else if (ActionItem["ActionType"] == "DROPBOX")
            {
                string TicketKey = ActionItem["InputString"];
                bResult = WA.SetComboItem(IE, doc, ActionItem["ID"], ActionItem["ListID"], TicketItem[TicketKey]);
            }
            else if (ActionItem["ActionType"] == "INPUT_TEXT")
            {
                string TicketKey = ActionItem["InputString"];
                bResult = WA.SetTextBox(IE, doc, ActionItem["ID"], TicketItem[TicketKey]);
            }
            else if (ActionItem["ActionType"] == "COPY_NUM")
            {
                TicketItem["Number"] = WA.GetText(IE, doc, ActionItem["ID"]);
            }
            else if (ActionItem["ActionType"] == "GOTOURL")
            {
                object URL = ActionItem["ID"] + TicketItem["Number"];
                GoToURL(URL);
            }
            else if (ActionItem["ActionType"] == "INPUT_UPLOAD")
            {
                bResult = WA.SetTextBox(IE, doc, ActionItem["ID"], TxtForUpload);
            }
            else if (ActionItem["ActionType"] == "CALLEVENT")
            {
                bResult = WA.CallEvent(IE, doc, ActionItem["ID"], ActionItem["ListID"]);
            }

            WA.LoadingWait(doc);

            if (ActionItem["ManualWait"] == "O")
            {
                WA.ManualWait(500);
            }

            return bResult;
        }
    }
}