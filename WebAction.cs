using mshtml;
using SHDocVw;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Collections;

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

        public void TotalWait(InternetExplorer IE, HTMLDocument doc, string ID = null, string ElemType = "ID", string WaitType = "Exist", int nTimer = 5000)
        {            
            LoadingWait(doc, nTimer);
            WaitPageLoading(IE, nTimer);
            if (ID != null)
            {
                WaitUntilID(doc, ID, ElemType, WaitType);
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
            TotalWait(IE, doc, ID, ElemType);

            if (ElemType == "ID" || ElemType == null)
            {
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
            TotalWait(IE, doc);
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
            TotalWait(IE, doc, ParentID);
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

            TotalWait(IE, doc, ListID);
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

                    TotalWait(IE, doc);
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

        public bool ReadComboItem(InternetExplorer IE, HTMLDocument doc,
                        string Depth1Txt, string Depth2Txt, string Depth3Txt, string Depth4Txt,
                        int nCurDepth, int nLastDepth,
                        string ID, string ID2, string ID3, string ID4, string Comment,
                        ref List<Dictionary<KPMReadInfo, List<string>>> ReadList)
        {
            bool bResult = true;

#if (DEBUG)
            DebugPrint("[ReadComboItem]");
#endif
            string[] IDAray = { ID, ID2, ID3, ID4 };
            string[] ItemIDAray = { ID + "_items", ID2 + "_items", ID3 + "_items", ID4 + "_items" };

            TotalWait(IE, doc, IDAray[nCurDepth]);
            IHTMLElement SelectedElement = doc.getElementById(IDAray[nCurDepth]);
            SelectedElement.click();

            TotalWait(IE, doc, ItemIDAray[nCurDepth]);
            IHTMLElementCollection elemcoll = doc.getElementById(ItemIDAray[nCurDepth]).children as IHTMLElementCollection;

            Dictionary<KPMReadInfo, List<string>> LastList = null;
            List<string> TextList = null;
            KPMReadInfo nInfo = null;
            int nLen = elemcoll.length;
            foreach (IHTMLElement elem in elemcoll)
            {
                string ElemText = elem.innerText;
                
                if (nCurDepth < nLastDepth - 1)
                {
                    if (nCurDepth == 0)
                        Depth1Txt = ElemText;
                    else if (nCurDepth == 1)
                        Depth2Txt = ElemText;
                    else if (nCurDepth == 2)
                        Depth3Txt = ElemText;
                    else
                        Depth4Txt = ElemText;

                    elem.click();
                    TotalWait(IE, doc);

                    int nNextdepth = nCurDepth + 1;
                    ReadComboItem(IE, doc,
                        Depth1Txt, Depth2Txt, Depth3Txt, Depth4Txt,
                        nNextdepth, nLastDepth,
                        ID, ID2, ID3, ID4, Comment,
                        ref ReadList);
                }
                else
                {   // End. Let's add all texts.
                    if (nInfo == null)
                    {
                        nInfo = new KPMReadInfo();
                        nInfo.Depth1 = Depth1Txt; nInfo.Depth2 = Depth2Txt;
                        nInfo.Depth3 = Depth3Txt; nInfo.Depth4 = Depth4Txt;
                        nInfo.nDepthCnt = nCurDepth;
                        nInfo.sDataType = Comment;
                    }
                    if (TextList == null)
                    {
                        TextList = new List<string>();
                    }
                    if (LastList == null)
                    {
                        LastList = new Dictionary<KPMReadInfo, List<string>>();
                    }

                    TextList.Add(ElemText);
                }
            }
            if (LastList != null)
            {
                LastList.Add(nInfo, TextList);
                ReadList.Add(LastList);
            }
            return bResult;
        }

        public bool SetTextBox(InternetExplorer IE, HTMLDocument doc, string ID, string InputText)
        {
            bool bResult = true;
#if (DEBUG)
            DebugPrint("[SetTextBox]" + ID + ", " + InputText);
#endif

            TotalWait(IE, doc, ID);
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

            TotalWait(IE, doc, ID);
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
            TotalWait(IE, doc, ID);

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