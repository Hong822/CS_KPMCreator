using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;

namespace CS_KPMCreator
{
    public partial class Form1 : Form
    {
        private ExcelControl g_ExcelTool = new ExcelControl();
        private WebControl g_WebControl = new WebControl();
        private List<Dictionary<string, string>> g_TicketItemList = new List<Dictionary<string, string>>();

        protected InternetExplorerDriverService _driverSerivce = null;
        protected InternetExplorerOptions _options = null;
        protected InternetExplorerDriver _driver = null;

        public Form1()
        {
            InitializeComponent();

            _driverSerivce = InternetExplorerDriverService.CreateDefaultService();
            _driverSerivce.HideCommandPromptWindow = true;

            _options = new InternetExplorerOptions();
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

        private void bStartCreation_Click(object sender, EventArgs e)
        {
            // Open Browser
            _driver = new InternetExplorerDriver(_driverSerivce, _options);

            rbB2B.Checked = true;
            rbB2C.Checked = false;

            if (rbB2C.Checked == true)
            {

            }
            else
            {
                if (rbB2B.Checked == true)
                {
                    string sID = tB2BID.Text;
                    string sPW = tB2BPW.Text;

                    if (sID != null && sPW != null)
                    {
                        _driver.Navigate().GoToUrl("https://sso.volkswagen.de/kpmweb/Index.action");
                        _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    }
                    else
                    {
                        _driver.Navigate().GoToUrl("https://" + sID + ":" + sPW + "@sso.volkswagen.de/kpmweb/Index.action");
                        _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    }                    
                }
            }
            
            // Start Ticket Creation 
            g_WebControl.CreateTickets(g_TicketItemList);
        }
    }
}