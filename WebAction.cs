using mshtml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SHDocVw;
using System;
using System.Collections.Generic;

namespace CS_KPMCreator
{
    internal class WebAction
    {
        private const int nTimeoutLimit = 10000;
        private Util g_Util = null;
        public WebDriverWait wait = null;

        public WebAction(ref Util util)
        {
            g_Util = util;
        }

        public void ManualWait(int nTimer = 3000)
        {
            int nCurTime = 0;
            int nTickGap = 1000;
            while (nCurTime < nTimer)
            {
#if (DEBUG)
                g_Util.DebugPrint("Manuel Wait. (" + nCurTime / 1000 + "/" + nTimer / 1000 + "sec)");
#endif
                System.Threading.Thread.Sleep(nTickGap);
                nCurTime += nTickGap;
            }
        }

        /*
        Wait Functions for IE
        */

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
                    g_Util.DebugPrint("ProgressDialog Wait (" + nCurTime / 1000 + "sec)");
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
            g_Util.DebugPrint("Wait for Page Loading.");
#endif
            while (IE.Busy == true || IE.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
            {
#if (DEBUG)
                g_Util.DebugPrint("Wait for Page Loading. (" + nCurTime / 1000 + "sec)");
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
                        g_Util.DebugPrint("Wait ID = " + ID);
#endif
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                    {
#if (DEBUG)
                        g_Util.DebugPrint("Find ID = " + ID);
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
                        g_Util.DebugPrint("Wait ID = " + ID);
#endif
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                    {
#if (DEBUG)
                        g_Util.DebugPrint("Find ID = " + ID);
#endif
                        break;
                    }
                }
            }
        }

        /*
        Wait Functions for Selenium
        */

        public void TotalWait(IWebDriver FF_driver, string ID = null, string ElemType = "ID", string WaitType = "Exist", int nTimer = 1000)
        {
            LoadingWait(FF_driver, nTimer);

            if (ID != null)
            {
                WaitUntilID(FF_driver, ID, ElemType, WaitType);
            }
        }

        public void LoadingWait(IWebDriver FF_driver, int nTimer)
        {
            string FindID = "progressDialog_modal";

            int nCurTime = 0;
            IWebElement element = null;
            while (true)
            {
                if (nCurTime > nTimeoutLimit)
                {
                    g_Util.DebugPrint("LoadingDialog Check Time Out. " + nCurTime + "ms");
                    break;
                }

                try
                {
                    element = FF_driver.FindElement(By.Id(FindID));
#if (DEBUG)
                    g_Util.DebugPrint("ProgressDialog Wait (" + nCurTime / 1000 + "sec)");
#endif
                    nCurTime += nTimer;
                    System.Threading.Thread.Sleep(nTimer);
                }
                catch (System.Exception e)
                {
                    // No 'progressDialog_modal' Element = Loading is finished.\
                    g_Util.DebugPrint("No progressDialog_modal! Loading is done.");
                    break;
                }
            }
        }

        private void WaitUntilID(IWebDriver FF_driver, string ID, string ElemType = "ID", string WaitType = "Exist")
        {
            // Wait until Element becomes clickable.
            int nCurTime = 0;
            int nTimer = 500;
            IWebElement element = null;
            while (true && element == null)
            {
                if (nCurTime > nTimeoutLimit)
                {
                    g_Util.DebugPrint("FindElement Check Time Out. " + nCurTime + "ms");
                    break;
                }

                try
                {
                    if (ElemType == "ID")
                    {
                        element = FF_driver.FindElement(By.Id(ID));
                    }
                    else if (ElemType == "NAME")
                    {
                        element = FF_driver.FindElement(By.Name(ID));
                    }
                }
                catch (System.Exception e)
                {
                    // No Element. Let's wait more.
#if (DEBUG)
                    g_Util.DebugPrint("WaitUntilID, Present (" + ID + ", " + nCurTime / 1000 + " sec)");
#endif
                    nCurTime += nTimer;
                    System.Threading.Thread.Sleep(nTimer);
                }
            }

            nCurTime = 0;
            element = null;
            while (true && element == null)
            {
                if(nCurTime > nTimeoutLimit)
                {
                    g_Util.DebugPrint("Clickable Check Time Out. " + nCurTime + "ms");
                    break;
                }
                try
                {
                    element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id(ID)));
                    g_Util.DebugPrint("Element became Clickable, " + ID);
                }
                catch (System.Exception e)
                {
                    // No Element. Let's wait more.
#if (DEBUG)
                    g_Util.DebugPrint("WaitUntilID, Clickable (" + ID + ", " + nCurTime / 1000 + " sec)");
#endif
                    nCurTime += nTimer;
                    System.Threading.Thread.Sleep(nTimer);
                }
            }
        }

        public bool ClickButton(InternetExplorer IE, HTMLDocument doc, string ID, string ElemType = "ID")
        {
            bool bResult = true;
#if (DEBUG)
            g_Util.DebugPrint("[ClickButton]" + ID);
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
                    g_Util.DebugPrint("Problem happen" + e);
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
                        break;
                    }
                    catch (System.Exception e)
                    {
                        g_Util.DebugPrint("[Exception] " + e);
                        bResult = false;
                    }
                }
            }

            return bResult;
        }

        public bool ClickButton(IWebDriver FF_driver, string ID, string ElemType = "ID")
        {
            bool bResult = true;
#if (DEBUG)
            g_Util.DebugPrint("[ClickButton]" + ID);
#endif
            TotalWait(FF_driver, ID, ElemType);

            IWebElement SelectedElement = null;

            if (ElemType == "NAME")
            {
                try
                {
                    SelectedElement = FF_driver.FindElement(By.Name(ID));
                }
                catch (System.Exception e)
                {
                    g_Util.DebugPrint("No Element despite Waiting. Skip this action.  " + ID );
                    return bResult;
                }
            }
            else   // if (ElemType == "ID" || ElemType == null)
            {
                try
                {
                    SelectedElement = FF_driver.FindElement(By.Id(ID));
                }
                catch (System.Exception e)
                {
                    g_Util.DebugPrint("No Element despite Waiting. Skip this action.  " + ID);
                    return bResult;
                }
            }

            try
            {
                //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id(ID)));
                SelectedElement.Click();
                g_Util.DebugPrint("Click is Called. (" + ID + ")");
            }
            catch (System.Exception e)
            {
                g_Util.DebugPrint("Problem happen" + e);
                bResult = false;
            }

            return bResult;
        }

        public bool ExecuteJS(InternetExplorer IE, HTMLDocument doc, string ID)
        {
            bool bResult = true;
#if (DEBUG)
            g_Util.DebugPrint("[ExecuteJS] " + ID);
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
                g_Util.DebugPrint("[Exception] " + e);
                bResult = false;
            }

            return bResult;
        }

        public bool ExecuteJS(IWebDriver FF_driver, string ID)
        {
            bool bResult = true;
#if (DEBUG)
            g_Util.DebugPrint("[ExecuteJS] " + ID);
#endif
            TotalWait(FF_driver);

            try
            {
                var jse = (IJavaScriptExecutor)FF_driver;
                jse.ExecuteScript(ID, "");
            }
            catch (System.Exception e)
            {
                g_Util.DebugPrint("[Exception] " + e);
                bResult = false;
            }

            return bResult;
        }

        public bool SetComboItem(InternetExplorer IE, HTMLDocument doc, string ParentID, string ListID, string InputValue)
        {
            bool bResult = true;
            if (InputValue != null)
            {
#if (DEBUG)
                g_Util.DebugPrint("[SetComboItem]" + ParentID + ", " + InputValue);
#endif
                TotalWait(IE, doc, ParentID);
                IHTMLElement SelectedElement = doc.getElementById(ParentID);
                try
                {
                    SelectedElement.click();
                }
                catch (System.Exception e)
                {
                    g_Util.DebugPrint("[Exception] " + e);
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
                            g_Util.DebugPrint("[Exception] " + e);
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
                    g_Util.DebugPrint("No Item in List. Input Value = " + InputValue);
                }
            }

            return bResult;
        }

        public bool SetComboItem(IWebDriver FF_driver, string ParentID, string ListID, string InputValue)
        {
            bool bResult = true;

#if (DEBUG)
            g_Util.DebugPrint("[SetComboItem]" + ParentID + ", " + InputValue);
#endif
            TotalWait(FF_driver, ParentID);
            IWebElement SelectedElement = FF_driver.FindElement(By.Id(ParentID));
                       
            try
            {
                SelectedElement.Click();
                g_Util.DebugPrint("Click called: " + ParentID);
            }
            catch (System.Exception e)
            {
                g_Util.DebugPrint("[Exception] " + e);
                bResult = false;
                return bResult;
            }

            TotalWait(FF_driver, ListID);

            bool bFind = false;
            SelectedElement = FF_driver.FindElement(By.Id(ListID));
            foreach (var Elem in SelectedElement.FindElements(By.TagName("li")))
            {
                if (Elem.Text == InputValue)
                {
                    Elem.Click();
                    bFind = true;
                }
            }

            if (bFind == false)
            {
                g_Util.DebugPrint("No Item in List. Input Value = " + InputValue);
            }

            return bResult;
        }

        public bool SetComboItemBySelect(IWebDriver FF_driver, string ParentID, string InputValue)
        {
            bool bResult = true;

#if (DEBUG)
            g_Util.DebugPrint("[SetComboItemBySelect]" + ParentID + ", " + InputValue);
#endif
            TotalWait(FF_driver, ParentID);
            IWebElement SelectedElement = FF_driver.FindElement(By.Id(ParentID));

            try
            {
                var Elements = new SelectElement(SelectedElement);
                Elements.SelectByText(InputValue);
            }
            catch (System.Exception e)
            {
                g_Util.DebugPrint("[Exception] " + e);
                bResult = false;
                return bResult;
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
            g_Util.DebugPrint("[ReadComboItem]");
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
                    g_Util.DebugPrint("\t[ReadComboItem] Text Add = (" + Depth1Txt + ")(" + Depth2Txt + ")(" + Depth3Txt + ")(" + Depth4Txt + ")-" + ElemText);
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
            if (InputText == null)
            {
                InputText = "";
            }
#if (DEBUG)
            g_Util.DebugPrint("[SetTextBox]" + ID + ", " + InputText);
#endif

            TotalWait(IE, doc, ID);
            IHTMLElement SelectedElement = doc.getElementById(ID);

            try
            {
                SelectedElement.setAttribute("innerText", InputText);
            }
            catch (System.Exception e)
            {
                g_Util.DebugPrint("[Exception] " + e);
                bResult = false;
            }

            return bResult;
        }

        public bool SetTextBox(IWebDriver FF_driver, string ID, string InputText)
        {
            bool bResult = true;
            if (InputText != null)
            {
#if (DEBUG)
                g_Util.DebugPrint("[SetTextBox]" + ID + ", " + InputText);
#endif

                TotalWait(FF_driver, ID);
                IWebElement SelectedElement = FF_driver.FindElement(By.Id(ID));

                try
                {
                    SelectedElement.SendKeys(InputText);
                }
                catch (System.Exception e)
                {
                    g_Util.DebugPrint("[Exception] " + e);
                    bResult = false;
                }
            }
            return bResult;
        }

        public bool CallEvent(InternetExplorer IE, HTMLDocument doc, string ID, string EventType)
        {
            bool bResult = true;

#if (DEBUG)
            g_Util.DebugPrint("[CallEvent]" + ID + ", " + EventType);
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
                    g_Util.DebugPrint("[Exception] " + e);
                    bResult = false;
                }
            }

            return bResult;
        }

        public bool CallEvent(IWebDriver FF_driver, string ID, string EventType)
        {
            bool bResult = true;

#if (DEBUG)
            g_Util.DebugPrint("[CallEvent]" + ID + ", " + EventType);
#endif

            TotalWait(FF_driver, ID);
            IWebElement SelectedElement = FF_driver.FindElement(By.Id(ID));

            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)FF_driver;
            if (EventType == "onchange")
            {
                jsExecutor.ExecuteScript("$(arguments[0]).change();", SelectedElement);
            }
            else if (EventType == "onkeyup")
            {
                jsExecutor.ExecuteScript("$(arguments[0]).keyup();", SelectedElement);
            }

            return bResult;
        }

        public string GetText(InternetExplorer IE, HTMLDocument doc, string ID)
        {
#if (DEBUG)
            g_Util.DebugPrint("[GetText]" + ID);
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
                g_Util.DebugPrint("[Exception] " + e);
                sText = "[Exception]";
            }
            return sText;
        }

        public string GetText(IWebDriver FF_driver, string ID)
        {
#if (DEBUG)
            g_Util.DebugPrint("[GetText]" + ID);
#endif
            TotalWait(FF_driver, ID);

            IWebElement SelectedElement = FF_driver.FindElement(By.Id(ID));

            string sText = null;
            try
            {
                sText = SelectedElement.Text;
            }
            catch (System.Exception e)
            {
                g_Util.DebugPrint("[Exception] " + e);
                sText = "[Exception]";
            }
            return sText;
        }
    }
}