using mshtml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SHDocVw;
using System.Collections.Generic;
using System.Net.Mail;

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

        public bool TotalWait(InternetExplorer IE, HTMLDocument doc, string ID = null, string ElemType = "ID", int nTimer = 5000)
        {
            bool bResult = LoadingWait(doc, nTimer);
            bResult = WaitPageLoading(IE, nTimer);
            if (ID != null)
            {
                bResult = WaitUntilID(doc, ID, ElemType);
            }
            return bResult;
        }

        public bool LoadingWait(HTMLDocument doc, int nTimer = 1000)
        {
            bool bResult = true;
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
                    bResult = false;
                }
                else
                {
                    bResult = true;
                    break;
                }
            }

            return bResult;
        }

        public bool WaitPageLoading(InternetExplorer IE, int nTimer = 1000)
        {
            bool bResult = true;
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

            return bResult;
        }

        private bool WaitUntilID(HTMLDocument doc, string ID, string ElemType = "ID")
        {
            bool bResult = true;
            int nTimer = 60;
            int nTry = 0;

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
                        bResult = false;
                    }
                    else
                    {
#if (DEBUG)
                        g_Util.DebugPrint("Find ID = " + ID);
#endif
                        bResult = true;
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
                        bResult = false;
                    }
                    else
                    {
#if (DEBUG)
                        g_Util.DebugPrint("Find ID = " + ID);
#endif
                        bResult = true;
                        break;
                    }
                }
            }

            return bResult;
        }

        public bool ClickButton(InternetExplorer IE, HTMLDocument doc, string ID, string ElemType = "ID")
        {
            bool bResult = true;
#if (DEBUG)
            g_Util.DebugPrint("[ClickButton]" + ID);
#endif
            if (TotalWait(IE, doc, ID, ElemType) == false)
            {
                if (ElemType != "NAME")
                {
                    g_Util.DebugPrint("Element Waiting False. Skip this action.  " + ID);
                    return true;
                }
            }

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

        public bool SetComboItem(InternetExplorer IE, HTMLDocument doc, string ParentID, string SubID, string InputValue)
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

                TotalWait(IE, doc, SubID);
                IHTMLElementCollection elemcoll = doc.getElementById(SubID).children as IHTMLElementCollection;

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

        /*
        Selenium Functions
        */

        //////////////////////////////////////////
        public bool TotalWait(IWebDriver FF_driver, string ID = null, string ElemType = "ID", int nTimer = 1000)
        {
            bool bResult = LoadingWait(FF_driver, nTimer);

            if (ID != null)
            {
                bResult = WaitUntilID(FF_driver, ID, ElemType);
            }

            return bResult;
        }

        public bool LoadingWait(IWebDriver FF_driver, int nTimer)
        {
            bool bResult = true;
            string FindID = "progressDialog_modal";

            int nCurTime = 0;
            IWebElement element = null;
            while (true)
            {
                if (nCurTime > nTimeoutLimit)
                {
                    g_Util.DebugPrint("LoadingDialog Check Time Out. " + nCurTime + "ms");
                    bResult = false;
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
                    bResult = false;
                }
                catch (System.Exception e)
                {
                    // No 'progressDialog_modal' Element = Loading is finished.\
                    g_Util.DebugPrint("No progressDialog_modal! Loading is done.");
                    bResult = true;
                    break;
                }
            }

            return bResult;
        }

        private bool WaitUntilID(IWebDriver FF_driver, string ID, string ElemType = "ID")
        {
            // Wait until Element becomes clickable.
            bool bResult = true;
            int nCurTime = 0;
            int nTimer = 500;
            IWebElement element = null;
            while (true && element == null)
            {
                if (ID != "problem_middle" && nCurTime > nTimeoutLimit)
                {
                    g_Util.DebugPrint("FindElement Check Time Out. " + nCurTime + "ms");
                    bResult = false;
                    break;
                }

                try
                {
                    if (ElemType == "NAME")
                    {
                        element = FF_driver.FindElement(By.Name(ID));
                    }
                    else if (ElemType == "XPATH_TEXT")
                    {
                        string path = "//*[contains(text(), '" + ID + "')]";
                        //"//*[contains(text(), 'Send to coordinator')]"
                        element = FF_driver.FindElement(By.XPath(path));
                    }
                    else
                    {
                        element = FF_driver.FindElement(By.Id(ID));
                    }

                    g_Util.DebugPrint("[WaitUntilID] Find Success. " + ID);
                }
                catch (System.Exception e)
                {
                    // No Element. Let's wait more.
#if (DEBUG)
                    g_Util.DebugPrint("WaitUntilID, Present (" + ID + ", " + nCurTime / 1000 + " sec)");
#endif
                    nCurTime += nTimer;
                    System.Threading.Thread.Sleep(nTimer);
                    bResult = false;
                }
            }
            //            nCurTime = 0;
            //            element = null;
            //            while (true && element == null)
            //            {
            //                if (ID != "problem_middle" && nCurTime > nTimeoutLimit)
            //                {
            //                    g_Util.DebugPrint("Clickable Check Time Out. " + nCurTime + "ms");
            //                    bResult = false;
            //                    break;
            //                }
            //                try
            //                {
            //                    element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id(ID)));
            //                    g_Util.DebugPrint("Element became Clickable, " + ID);
            //                    bResult = true;
            //                }
            //                catch (System.Exception e)
            //                {
            //                    // No Element. Let's wait more.
            //#if (DEBUG)
            //                    g_Util.DebugPrint("WaitUntilID, Clickable (" + ID + ", " + nCurTime / 1000 + " sec)");
            //#endif
            //                    nCurTime += nTimer;
            //                    System.Threading.Thread.Sleep(nTimer);
            //                    bResult = false;
            //                }
            //            }

            return bResult;
        }

        public void DoDefaultAction(IWebDriver FF_driver)
        {
            g_Util.DebugPrint("[DoDefaultAction]");
            ClickButton(FF_driver, "problemForm:bewertungAltKopf:bewertungAltKopf");
        }

        public IWebElement FindElemnt(IWebDriver FF_driver, string ID, string ElemType = "ID", int nTimer = 300)
        {
            int nCurTime = 0;
            IWebElement SelectedElement = null;

            while (true)
            {
                if (nCurTime > nTimeoutLimit)
                {
                    g_Util.DebugPrint("FindElement Check Time Out. " + nCurTime + "ms");
                    SelectedElement = null;
                    break;
                }

                try
                {
                    if (ElemType == "NAME")
                    {
                        SelectedElement = FF_driver.FindElement(By.Name(ID));
                    }
                    else if (ElemType == "XPATH_TEXT")
                    {
                        string path = "//*[contains(text(), '" + ID + "')]";
                        //"//*[contains(text(), 'Send to coordinator')]"
                        SelectedElement = FF_driver.FindElement(By.XPath(path)).FindElement(By.XPath("./.."));
                    }
                    else
                    {
                        SelectedElement = FF_driver.FindElement(By.Id(ID));
                    }
                    break;
                }
                catch (System.Exception e)
                {
                    g_Util.DebugPrint("No Element. Let's wait" + ID);
                    System.Threading.Thread.Sleep(nTimer);
                    nCurTime += nTimer;
                    SelectedElement = null;
                }
            }
            return SelectedElement;
        }

        public bool ClickElemnt(IWebElement SelectedElement, string ID, string ElemType = "ID", int nTimer = 300)
        {
            bool bResult = false;

            int nCurTime = 0;
            while (true)
            {
                if (nCurTime > nTimeoutLimit)
                {
                    g_Util.DebugPrint("ClickElemnt Check Time Out. " + nCurTime + "ms");
                    bResult = false;
                    break;
                }

                try
                {
                    SelectedElement.Click();

                    g_Util.DebugPrint("Click is Called. (" + ID + ")");
                    bResult = true;
                    break;
                }
                catch (System.Exception e)
                {
                    g_Util.DebugPrint("Element isn't Clickable. Let's wait. " + ID);
                    System.Threading.Thread.Sleep(nTimer);
                    nCurTime += nTimer;
                    bResult = false;
                }
            }
            return bResult;
        }

        public bool ClickButton(IWebDriver FF_driver, string ID, string ElemType = "ID", bool bSkipable = false)
        {
            bool bResult = true;
#if (DEBUG)
            g_Util.DebugPrint("[ClickButton]" + ID);
#endif
            if (TotalWait(FF_driver, ID, ElemType) == false)
            {
                if (bSkipable == true)
                {
                    g_Util.DebugPrint("Element Waiting False. Skip this action.  " + ID);
                    return true;
                }
                //else
                //{
                //    DoDefaultAction(FF_driver);
                //}
            }

            IWebElement SelectedElement = FindElemnt(FF_driver, ID, ElemType);
            if (SelectedElement == null)
            {
                g_Util.DebugPrint("No Element despite Waiting. Skip this action.  " + ID);
                bResult = true;
            }
            else
            {
                if (ClickElemnt(SelectedElement, ID) == false)
                {
                    g_Util.DebugPrint("Click Fail");
                    bResult = false;
                }
                else
                {
                    bResult = true;
                }
            }

            return bResult;
        }

        public bool ClickSendtoCoordinator(IWebDriver FF_driver, string ID, string ElemType = "ID", bool bSkipable = false)
        {
            bool bResult = true;
#if (DEBUG)
            g_Util.DebugPrint("[ClickSendtoCoordinator]" + ID);
#endif
            IWebElement SelectedElement = FindElemnt(FF_driver, ID, ElemType);
            if (SelectedElement == null)
            {
                g_Util.DebugPrint("No Element despite Waiting. Skip this action.  " + ID);
                bResult = false;
            }
            else
            {
                try
                {
                    SelectedElement = FF_driver.FindElement(By.XPath(".."));
                    SelectedElement.Click();
                    bResult = true;
                }
                catch (System.Exception e)
                {
                    g_Util.DebugPrint("XPath(..) Fail  " + ID);
                    bResult = false;
                }
            }
            return bResult;
        }

        public bool ExecuteScript(IWebDriver FF_driver, string ID, int nTimer = 300)
        {
            bool bResult = false;

            int nCurTime = 0;
            while (true)
            {
                if (nCurTime > nTimeoutLimit)
                {
                    g_Util.DebugPrint("ExecuteScript Time Out. " + nCurTime + "ms");
                    bResult = false;
                    break;
                }

                try
                {
                    var jse = (IJavaScriptExecutor)FF_driver;
                    jse.ExecuteScript(ID, "");
                    g_Util.DebugPrint("ExecuteScript is Called.   " + ID);
                    bResult = true;
                    break;
                }
                catch (System.Exception e)
                {
                    g_Util.DebugPrint("ExecuteScript isn't possible. Let's wait" + ID);
                    System.Threading.Thread.Sleep(nTimer);
                    nCurTime += nTimer;
                    bResult = false;
                }
            }
            return bResult;
        }

        public bool ExecuteJS(IWebDriver FF_driver, string ID)
        {
            bool bResult = true;
#if (DEBUG)
            g_Util.DebugPrint("[ExecuteJS] " + ID);
#endif
            if (TotalWait(FF_driver) == false)
            {
                g_Util.DebugPrint("Element Waiting False. Skip this action.  " + ID);
                bResult = true;
            }
            else
            {
                if (ExecuteScript(FF_driver, ID) == false)
                {
                    g_Util.DebugPrint("JS excution fail " + ID);
                    bResult = false;
                }
                else
                {
                    bResult = true;
                }
            }

            return bResult;
        }

        public bool SetComboItem(IWebDriver FF_driver, string ID, string InputValue)
        {
            bool bResult = true;

#if (DEBUG)
            g_Util.DebugPrint("[SetComboItem]" + ID + ", " + InputValue);
#endif
            if (TotalWait(FF_driver, ID) == false)
            {
                g_Util.DebugPrint("Element Waiting False. Skip this action.  " + ID);
                bResult = true;
            }
            else
            {
                IWebElement SelectedElement = FindElemnt(FF_driver, ID);
                if (SelectedElement == null)
                {
                    g_Util.DebugPrint("[Error] SetComboItem Element Find Fail ");
                    bResult = false;
                }
                else
                {
                    bool bFind = false;
                    foreach (var Elem in SelectedElement.FindElements(By.TagName("li")))
                    {
                        if (Elem.Text == InputValue)
                        {
                            Elem.Click();
                            g_Util.DebugPrint("Select: " + InputValue);

                            bFind = true;
                            bResult = true;
                        }
                    }

                    if (bFind == false)
                    {
                        g_Util.DebugPrint("[Warning] No Item in List. Input Value = " + InputValue);
                        bResult = false;
                    }
                    else
                    {
                        bResult = true;
                    }
                }
            }

            return bResult;
        }

        //        public bool SetComboItemBySelect(IWebDriver FF_driver, string ParentID, string InputValue)
        //        {
        //            bool bResult = true;

        //#if (DEBUG)
        //            g_Util.DebugPrint("[SetComboItemBySelect]" + ParentID + ", " + InputValue);
        //#endif
        //            if (TotalWait(FF_driver, ParentID) == false)
        //            {
        //                g_Util.DebugPrint("Element Waiting False. Skip this action.  " + ParentID);
        //                return true;
        //            }
        //            IWebElement SelectedElement = FF_driver.FindElement(By.Id(ParentID));

        //            try
        //            {
        //                var Elements = new SelectElement(SelectedElement);
        //                Elements.SelectByText(InputValue);
        //            }
        //            catch (System.Exception e)
        //            {
        //                g_Util.DebugPrint("[Exception] " + e);
        //                bResult = false;
        //                return bResult;
        //            }

        //            return bResult;
        //        }

        public bool SendKey(IWebElement SelectedElement, string InputText, int nTimer = 300)
        {
            bool bResult = false;

            int nCurTime = 0;
            while (true)
            {
                if (nCurTime > nTimeoutLimit)
                {
                    g_Util.DebugPrint("ExecuteScript Time Out. " + nCurTime + "ms");
                    bResult = false;
                    break;
                }

                try
                {
                    SelectedElement.Clear();
                    SelectedElement.SendKeys(InputText);
                    bResult = true;
                    break;
                }
                catch (System.Exception e)
                {
                    g_Util.DebugPrint("[Warning] Sendkey Fail" + InputText);
                    System.Threading.Thread.Sleep(nTimer);
                    nCurTime += nTimer;
                    bResult = false;
                }
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
                if (TotalWait(FF_driver, ID) == false)
                {
                    g_Util.DebugPrint("Element Waiting False. Skip this action.  " + ID);
                    bResult = true;
                }
                else
                {
                    IWebElement SelectedElement = FindElemnt(FF_driver, ID);
                    if (SelectedElement == null)
                    {
                        g_Util.DebugPrint("[Error] SetComboItem Element Find Fail " + ID);
                        bResult = false;
                    }
                    else
                    {
                        if (SendKey(SelectedElement, InputText) == false)
                        {
                            g_Util.DebugPrint("[Error] SendKeyFail. " + ID + "," + InputText);
                            bResult = false;
                        }
                        else
                        {
                            bResult = true;
                        }
                    }
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

            if (TotalWait(FF_driver, ID) == false)
            {
                g_Util.DebugPrint("Element Waiting False. Skip this action.  " + ID);
                bResult = true;
            }
            else
            {
                IWebElement SelectedElement = FindElemnt(FF_driver, ID);
                if (SelectedElement == null)
                {
                    g_Util.DebugPrint("[Error] SetComboItem Element Find Fail.  " + ID);
                    bResult = false;
                }
                else
                {
                    IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)FF_driver;
                    if (EventType == "onchange")
                    {
                        jsExecutor.ExecuteScript("$(arguments[0]).change();", SelectedElement);
                    }
                    else if (EventType == "onkeyup")
                    {
                        jsExecutor.ExecuteScript("$(arguments[0]).keyup();", SelectedElement);
                    }
                    bResult = true;
                }
            }

            return bResult;
        }

        public string GetText(IWebDriver FF_driver, string ID)
        {
            string sText = null;
#if (DEBUG)
            g_Util.DebugPrint("[GetText]" + ID);
#endif
            if (TotalWait(FF_driver, ID) == false)
            {
                g_Util.DebugPrint("Element Waiting False. Skip this action.  " + ID);
                sText = "No Elem";
            }
            else
            {
                IWebElement SelectedElement = FindElemnt(FF_driver, ID);
                if (SelectedElement == null)
                {
                    g_Util.DebugPrint("[Error] SetComboItem Element Find Fail.  " + ID);
                    sText = "No Elem";
                }
                else
                {

                    try
                    {
                        sText = SelectedElement.Text;
                    }
                    catch (System.Exception e)
                    {
                        g_Util.DebugPrint("[Error] " + e);
                        sText = "[No Number]";
                    }
                }
            }
            return sText;
        }
    }
}