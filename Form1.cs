using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CS_KPMCreator
{
    public partial class Form1 : Form
    {
        private ExcelControl g_ExcelTool = new ExcelControl();
        private WebControl g_WebControl = new WebControl();
        private List<Dictionary<string, string>> g_TicketItemList = new List<Dictionary<string, string>>();

        protected String g_sWebtype = "";
        protected DriverService _driverSerivce = null;
        protected DriverOptions _options = null;
        protected RemoteWebDriver _driver = null;

        public Form1()
        {
            InitializeComponent();
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

                // Date read from Excel Files
                ReadExcelValue();
            }
        }

        private void ReadExcelValue()
        {
            // open KPM Doc
            Excel.Application ap = new Excel.Application();
            Excel.Workbook wb = ap.Workbooks.Open(tExcelPath.Text);
            Excel.Worksheet ws_KPMCreate = wb.Worksheets["kpmcreate"];
            ap.Visible = true;

            // Fill in ticketItemList with ticket items
            g_ExcelTool.FillDictionary(g_TicketItemList, ws_KPMCreate);
        }

        private void SetBrowser()
        {
            if (rbFirefox.Checked == true)
            {
                _driverSerivce = FirefoxDriverService.CreateDefaultService(); ;
                _driverSerivce.HideCommandPromptWindow = true;
                _options = new FirefoxOptions();                
                _driver = new FirefoxDriver((FirefoxDriverService)_driverSerivce, (FirefoxOptions)_options);
                g_sWebtype = "Firefox";
            }
            else
            {
                if (rbChrome.Checked == true)
                {
                    //_options = new ChromeOptions();
                    //_driver = new ChromeDriver("C:\\KPM_Creator\\Driver\\chromedriver.exe", (ChromeOptions)_options);
                    _driver = new ChromeDriver("C:\\KPM_Creator\\Driver\\chromedriver.exe");
                    g_sWebtype = "Chrome";
                }
                else // Default = IE
                {
                    _driverSerivce = InternetExplorerDriverService.CreateDefaultService(); ;
                    _driverSerivce.HideCommandPromptWindow = true;
                    _options = new InternetExplorerOptions();                    
                    _driver = new InternetExplorerDriver((InternetExplorerDriverService)_driverSerivce, (InternetExplorerOptions)_options);
                    g_sWebtype = "IE";
                }
            }
        }
        private void GoToTheSite()
        {
            rbB2B.Checked = true;
            rbB2C.Checked = false;
            if (rbB2C.Checked == true)
            {
                string URL = "https://quasi.vw.vwg/kpm/kpmweb";
                _driver.Navigate().GoToUrl(URL);
                //_driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                MessageBox.Show("Wait!!! Login and Go to Main page. And then press OK", "Please Login KPM");
            }
            else
            {
                if (rbB2B.Checked == true)
                {
                    string sID = tB2BID.Text;
                    string sPW = tB2BPW.Text;

                    if (sID != null && sPW != null)
                    {
                        //string URL = "https://" + sID + ":" + sPW + "@sso.volkswagen.de/kpmweb/Index.action";
                        string URL = "https://www.naver.com";
                        var wait = new WebDriverWait(_driver, new TimeSpan(0, 0, 30));

                        try
                        {
                            _driver.Navigate().GoToUrl(URL);
                        }
                        catch (WebDriverException e)
                        {
                            Console.WriteLine(e);
                        }
                        
                        MessageBox.Show("Wait!!! Login and Go to Main page. And then press OK", "Please Login KPM");
                        int awer = 1;
                    }
                    else
                    {
                        _driver.Navigate().GoToUrl("https://" + sID + ":" + sPW + "@sso.volkswagen.de/kpmweb/Index.action");
                        _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    }
                }
            }
        }

        private void bStartCreation_Click(object sender, EventArgs e)
        {
            SetBrowser();   // Get Browser

            GoToTheSite();  // Go to KPM site
                        
            g_WebControl.CreateTickets(g_TicketItemList);   // Start Ticket Creation
        }
    }
}