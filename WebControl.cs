using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
//using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System.Windows.Forms;

namespace CS_KPMCreator 
{
    class WebControl
    {
        protected DriverService _driverSerivce = null;
        protected DriverOptions _options = null;
        //protected RemoteWebDriver _driver = null;
        protected IWebDriver _driver = null;

        RichTextBox g_richTB_Status = null;
        
        public void SetStatusBox(ref RichTextBox richTB_Status)
        {
            g_richTB_Status = richTB_Status;
        }
        public void SetBrowser(RadioButton rbFirefox, RadioButton rbChrome)
        {
            g_richTB_Status.Text = "I'm setting browser type...";

            if (rbFirefox.Checked == true)
            {
                _driverSerivce = FirefoxDriverService.CreateDefaultService(); ;
                _driverSerivce.HideCommandPromptWindow = true;
                _options = new FirefoxOptions();
                ((FirefoxOptions)_options).AddArgument("no-sandbox");
                _driver = new FirefoxDriver((FirefoxDriverService)_driverSerivce, (FirefoxOptions)_options);
            }
            else
            {
                if (rbChrome.Checked == true)
                {
                    _driverSerivce = ChromeDriverService.CreateDefaultService();
                    _driverSerivce.HideCommandPromptWindow = true;
                    _options = new ChromeOptions();
                    _driver = new ChromeDriver((ChromeDriverService)_driverSerivce, (ChromeOptions)_options);
                }
                else // Default = IE
                {
                    _driverSerivce = InternetExplorerDriverService.CreateDefaultService(); ;
                    _driverSerivce.HideCommandPromptWindow = true;
                    _options = new InternetExplorerOptions();
                    _driver = new InternetExplorerDriver((InternetExplorerDriverService)_driverSerivce, (InternetExplorerOptions)_options);
                }
            }
        }
        public void GoToTheSite(RadioButton rbB2B, RadioButton rbB2C, TextBox tB2BID, TextBox tB2BPW)
        {
            g_richTB_Status.Text = "I'm accessing KPM...";

            rbB2B.Checked = true;
            rbB2C.Checked = false;
            if (rbB2C.Checked == true)
            {
                string URL = "https://quasi.vw.vwg/kpm/kpmweb";
                _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                _driver.Navigate().GoToUrl(URL);
                MessageBox.Show("Wait!!! Login and Go to Main page. And then press OK", "Please Login KPM");
            }
            else
            {
                if (rbB2B.Checked == true)
                {
                    string sID = tB2BID.Text;
                    string sPW = tB2BPW.Text;
                    string URL = null;
                    if (sID != null && sPW != null)
                    {
                        URL = "https://" + sID + ":" + sPW + "@sso.volkswagen.de/kpmweb/Index.action";
                        //URL = "https://www.naver.com";
                        //URL = "https://quasi.vw.vwg/kpm/kpmweb";
                        //URL = "https://www.google.com";
                        //var wait = new WebDriverWait(_driver, new TimeSpan(0, 0, 1));
                    }
                    else
                    {
                        URL = "https://sso.volkswagen.de/kpmweb/Index.action";
                    }
                    _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
                    _driver.Navigate().GoToUrl(URL);
                }
                else
                {
                    MessageBox.Show("This case should happen.");
                }
            }
        }

        private void CreateOneTicket(Dictionary<string, string> dItem)
        {
            // open Browser

            // Log In

            // Action one by one
            //Action(Data, action);
        }

        public void CreateTickets(List<Dictionary<string, string>> TicketItemList, List<Dictionary<string, string>> LActionList)
        {
            g_richTB_Status.Text = "I'm creating KPM Ticket...";

            //Access each ticket items
            for (int nIdx = 0; nIdx < TicketItemList.Count; nIdx++)
            {
                Dictionary<string, string> dItem = TicketItemList[nIdx];
                CreateOneTicket(dItem); // create ticket
            }
        }
    }
}
