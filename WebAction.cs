using mshtml;
using SHDocVw;
using System.Windows.Forms;

namespace CS_KPMCreator
{
    internal class WebAction
    {
        private RichTextBox g_richTB_Status = null;

        public void SetStatusBox(ref RichTextBox richTB_Status)
        {
            g_richTB_Status = richTB_Status;
        }

        private void DebugPrint(string sDebugString)
        {
            g_richTB_Status.Text = sDebugString;
            System.Diagnostics.Debug.WriteLine(sDebugString);
        }

        public void ManualWait(int nTimer = 5000)
        {
            int nCurTime = 0;
            int nTickGap = 1000;
            while (nCurTime < nTimer)
            {
#if (DEBUG)
                DebugPrint("Manuel Wait. (" + nCurTime / 1000 + "/" + nTimer / 1000 + "sec)");
#endif
                System.Threading.Thread.Sleep(nTickGap);
                nCurTime += nTickGap;
            }
        }

        public void LoadingWait(HTMLDocument doc, int nTimer = 1000)
        {
            int nCurTime = 0;
            while (true)
            {
                if (doc.getElementById("progressDialog_modal") != null)
                {
#if (DEBUG)
                    DebugPrint("ProgressDialog Wait (" + nCurTime / 1000 + "sec)");
#endif
                    nCurTime += nTimer;
                    System.Threading.Thread.Sleep(nTimer);
                }
                else
                {
                    break;
                }
            }
        }

        public void WaitPageLoading(InternetExplorer IE, int nTimer = 1000)
        {
            int nCurTime = 0;
#if (DEBUG)
            DebugPrint("Wait for Page Loading.");
#endif
            while (IE.Busy == true || IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
#if (DEBUG)
                DebugPrint("Wait for Page Loading. (" + nCurTime / 1000 + "sec)");
#endif
                nCurTime += nTimer;
                System.Threading.Thread.Sleep(nTimer);
            }
        }

        private void WaitUntilID(HTMLDocument doc, string ID, string ElemType = "ID", string WaitType = "Exist")
        {
            int nTimer = 60;
            int nTry = 0;
            //if (WaitType == "Exist")
            if (ElemType == "ID")
            {
                while (nTry < nTimer)
                {
                    if (doc.getElementById(ID) == null)
                    {
                        nTry++;
#if (DEBUG)
                        DebugPrint("Wait ID = " + ID);
#endif
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                    {
#if (DEBUG)
                        DebugPrint("Find ID = " + ID);
#endif
                        break;
                    }
                }
            }
            else if (ElemType == "NAME")
            {
                while (nTry < nTimer)
                {
                    if (doc.getElementsByName(ID) == null)
                    {
                        nTry++;
#if (DEBUG)
                        DebugPrint("Wait ID = " + ID);
#endif
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                    {
#if (DEBUG)
                        DebugPrint("Find ID = " + ID);
#endif
                        break;
                    }
                }
            }
        }

        public bool ClickButton(InternetExplorer IE, HTMLDocument doc, string ID, string ElemType = "ID")
        {
            bool bResult = true;
#if (DEBUG)
            DebugPrint("[ClickButton]" + ID);
#endif
            WaitPageLoading(IE);

            if (ElemType == "ID" || ElemType == null)
            {
                WaitUntilID(doc, ID);
                IHTMLElement SelectedElement = doc.getElementById(ID);
                try
                {
                    SelectedElement.click();
                }
                catch (System.Exception e)
                {
                    DebugPrint("Problem happen" + e);
                    bResult = false;
                }
            }
            else if (ElemType == "NAME")
            {
                WaitUntilID(doc, ID, ElemType);
                IHTMLElementCollection SelectedElement = doc.getElementsByName(ID);

                foreach (IHTMLElement elem in SelectedElement)
                {
                    try
                    {
                        elem.click();
                    }
                    catch (System.Exception e)
                    {
                        DebugPrint("[Exception] " + e);
                        bResult = false;
                    }
                }
            }

            return bResult;
        }

        public bool ExecuteJS(InternetExplorer IE, HTMLDocument doc, string ID)
        {
            bool bResult = true;
#if (DEBUG)
            DebugPrint("[ExecuteJS] " + ID);
#endif
            WaitPageLoading(IE);
            var win = (IHTMLWindow2)doc.parentWindow;
            // here you call the Javascript

            try
            {
                win.execScript(ID);
            }
            catch (System.Exception e)
            {
                DebugPrint("[Exception] " + e);
                bResult = false;
            }

            return bResult;
        }

        public bool SetComboItem(InternetExplorer IE, HTMLDocument doc, string ParentID, string ListID, string InputValue)
        {
            bool bResult = true;

#if (DEBUG)
            DebugPrint("[SetComboItem]" + ParentID + ", " + InputValue);
#endif
            WaitPageLoading(IE);
            WaitUntilID(doc, ParentID);
            IHTMLElement SelectedElement = doc.getElementById(ParentID);
            try
            {
                SelectedElement.click();
            }
            catch (System.Exception e)
            {
                DebugPrint("[Exception] " + e);
                bResult = false;
                return bResult;
            }

            WaitUntilID(doc, ListID);
            IHTMLElementCollection elemcoll = doc.getElementById(ListID).children as IHTMLElementCollection;

            bool bFind = false;
            foreach (IHTMLElement elem in elemcoll)
            {
                //System.Diagnostics.Debug.WriteLine(" elem.GetType() ===> " + elem.GetType().ToString());
                //System.Diagnostics.Debug.WriteLine(" tagName == " + elem.tagName);
                //System.Diagnostics.Debug.WriteLine(" innerText == " + elem.innerText);
                //System.Diagnostics.Debug.WriteLine(" outerHTML == " + elem.outerHTML);
                //if (elem.getAttribute("id") != null)
                //{
                //    System.Diagnostics.Debug.WriteLine(" id == " + elem.getAttribute("id"));
                //}

                if (elem.innerText == InputValue)
                {
                    try
                    {
                        elem.click();
                    }
                    catch (System.Exception e)
                    {
                        DebugPrint("[Exception] " + e);
                        bResult = false;
                        return bResult;
                    }

                    elem.setAttribute("innerText", InputValue);
                    WaitPageLoading(IE);
                    bFind = true;
                }
                //System.Diagnostics.Debug.WriteLine(" ==================================================");
            }

            if (bFind == false)
            {
                DebugPrint("No Item in List. Input Value = " + InputValue);
            }

            return bResult;
        }

        public bool SetTextBox(InternetExplorer IE, HTMLDocument doc, string ID, string InputText)
        {
            bool bResult = true;
#if (DEBUG)
            DebugPrint("[SetTextBox]" + ID + ", " + InputText);
#endif

            WaitPageLoading(IE);
            WaitUntilID(doc, ID);

            IHTMLElement SelectedElement = doc.getElementById(ID);

            try
            {
                SelectedElement.setAttribute("innerText", InputText);
            }
            catch (System.Exception e)
            {
                DebugPrint("[Exception] " + e);
                bResult = false;
            }

            return bResult;
        }

        public bool CallEvent(InternetExplorer IE, HTMLDocument doc, string ID, string EventType)
        {
            bool bResult = true;

#if (DEBUG)
            DebugPrint("[CallEvent]" + ID + ", " + EventType);
#endif

            WaitPageLoading(IE);
            WaitUntilID(doc, ID);
            IHTMLElementCollection SelectedElement = doc.getElementsByName(ID);

            foreach (IHTMLElement3 elem in SelectedElement)
            {
                try
                {
                    elem.FireEvent("onkeyup", null);
                }
                catch (System.Exception e)
                {
                    DebugPrint("[Exception] " + e);
                    bResult = false;
                }
            }

            return bResult;
        }

        public string GetText(InternetExplorer IE, HTMLDocument doc, string ID)
        {
#if (DEBUG)
            DebugPrint("[GetText]" + ID);
#endif
            WaitPageLoading(IE);
            WaitUntilID(doc, ID);

            IHTMLElement SelectedElement = doc.getElementById(ID);

            string sText = null;
            try
            {
                sText = SelectedElement.getAttribute("innerText");
            }
            catch (System.Exception e)
            {
                DebugPrint("[Exception] " + e);
                sText = "[Exception]";
            }
            return sText;
        }
    }
}