using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CS_KPMCreator
{
    public partial class Form1 : Form
    {
        private ExcelControl g_ExcelTool = new ExcelControl();
        private WebControl g_WebControl = new WebControl();

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
                tExcelPath.Text = ExcelOpenDialog.FileName;
                tExcelPath.Text = "D:\\25_C_Projects\\CS_KPMCreator\\KPM_Ticket_Creator_V1.xlsm";
                // Date read from Excel Files
                ReadExcelValue();

                // Start Creation
                StartTicketCreation();
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
            List<Dictionary<string, string>> TicketItemList = new List<Dictionary<string, string>>();
            g_ExcelTool.FillDictionary(TicketItemList, ws_KPMCreate);

            // Start Ticket Creation 
            g_WebControl.CreateTickets(TicketItemList);
        }

        private void StartTicketCreation()
        {
        }
    }
}