using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace CS_KPMCreator
{
    internal class ExcelControl
    {
        private RichTextBox g_richTB_Status = null;
        private Excel.Application g_KPMExcel = null;
        private Excel.Workbook g_KPMWB = null;

        public void SetStatusBox(ref RichTextBox richTB_Status)
        {
            g_richTB_Status = richTB_Status;
        }

        public void CloseExcelControl()
        {
            if (g_KPMWB != null)
            {
                DebugPrint("I'm saving Excel File...");
                g_KPMWB.Save();
                g_KPMWB = null;
            }
            if (g_KPMExcel != null)
            {
                DebugPrint("I'm closing Excel File...");
                g_KPMExcel.Quit();
                g_KPMExcel = null;
            }
        }

        public void DebugPrint(string sDebugString)
        {
            g_richTB_Status.Text = sDebugString;
            System.Diagnostics.Debug.WriteLine(sDebugString);
        }

        private int FindColumn(Range rRow, string ColName)
        {
            int nCurCol = 1;

            while (rRow.Cells[1, nCurCol].Value != null)
            {
                if (rRow.Cells[1, nCurCol].Value == ColName)
                {
                    break;
                }
                nCurCol++;
            }

            return nCurCol;
        }

        private int LastRowPerColumn(Range rCol)
        {
            int nLastRow = 1;

            while (rCol.Cells[nLastRow, 1].Value != null)
            {
                nLastRow++;
            }

            return nLastRow;
        }

        private int LastColumnPerRow(Range rRow)
        {
            int nLastCol = 1;

            while (rRow.Cells[1, nLastCol].Value != null)
            {
                nLastCol++;
            }

            return nLastCol;
        }

        public void FillDictionary(List<Dictionary<string, string>> TicketItemList, Worksheet ws, int nStartRow, int nEndRow, int nStartCol, int nEndCol)
        {
            // Read and setting relevant data
            for (int RowIdx = nStartRow; RowIdx < nEndRow; RowIdx++)
            {
                Dictionary<string, string> NewItem = new Dictionary<string, string>();
                for (int ColIdx = nStartCol; ColIdx < nEndCol; ColIdx++)
                {
                    string sKey = Convert.ToString(((Range)ws.Cells[1, ColIdx]).Value);
                    string sVal = Convert.ToString(((Range)ws.Cells[RowIdx, ColIdx]).Value);
                    NewItem.Add(sKey, sVal);
                }

                TicketItemList.Add(NewItem);
            }
        }

        public void ReadExcelValue(System.Windows.Forms.TextBox tExcelPath, RadioButton rbB2B, RadioButton rbB2C, RadioButton rbAudi, RadioButton rbPorsche, ref List<Dictionary<string, string>> LTicketItemList, ref List<Dictionary<string, string>> LActionList)
        {
            DebugPrint("I'm reading KPM Excel File...");

            // ***Setting KPM Items***
            g_KPMExcel = new Excel.Application();
            try
            {
                g_KPMWB = g_KPMExcel.Workbooks.Open(tExcelPath.Text);
            }
            catch
            {
                MessageBox.Show("Please Select Excel Path.");
            }
            Excel.Worksheet ws_KPMCreate = g_KPMWB.Worksheets["kpmcreate"];
            g_KPMExcel.Visible = true;
            //ap.Visible = false;

            int nStartRow, nEndRow, nStartCol, nEndCol;
            nStartRow = 2;
            nEndRow = LastRowPerColumn(ws_KPMCreate.get_Range("C:C"));
            nStartCol = 1;
            nEndCol = LastColumnPerRow(ws_KPMCreate.get_Range("1:1"));

            // Fill in ticketItemList with ticket items
            FillDictionary(LTicketItemList, ws_KPMCreate, nStartRow, nEndRow, nStartCol, nEndCol);

            // ***Setting Actions***
            DebugPrint("I'm reading Action Excel File...");

            Excel.Application ActionApp = new Excel.Application();
            //string ActionExcelPath = "E:\\VS_Project\\repos\\Hong822\\CS_KPMCreator\\KPM_Action_Description.xlsm";
            string ActionExcelPath = "D:\\25_C_Projects\\Repos\\Hong822\\CS_KPMCreator\\KPM_Action_Description.xlsm";

            Excel.Workbook ActionWB = null;
            try
            {
                ActionWB = ActionApp.Workbooks.Open(ActionExcelPath);
            }
            catch
            {
                MessageBox.Show("Please check " + ActionExcelPath + ".");
            }

            Excel.Worksheet ActionWS = null;
            if (rbB2B.Checked == true)
            {
                if (rbPorsche.Checked == true)
                {
                    ActionWS = ActionWB.Worksheets["PO_B2B"];
                }
                else
                {
                    ActionWS = ActionWB.Worksheets["AU_B2B"];
                }
            }
            else
            {
                if (rbB2C.Checked == true)
                {
                    if (rbPorsche.Checked == true)
                    {
                        ActionWS = ActionWB.Worksheets["PO_B2C"];
                    }
                    else
                    {
                        ActionWS = ActionWB.Worksheets["AU_B2C"];
                    }
                }
                else
                {
                    // it's invalid condition.
                }
            }
            ActionApp.Visible = false;

            // Fill in ticketItemList with ticket items
            nStartRow = 2;
            nEndRow = LastRowPerColumn(ActionWS.get_Range("A:A"));
            nStartCol = 1;
            nEndCol = LastColumnPerRow(ActionWS.get_Range("1:1"));
            FillDictionary(LActionList, ActionWS, nStartRow, nEndRow, nStartCol, nEndCol);

            if (ActionApp.Visible == false)
            {
                ActionApp.Quit();
            }
        }

        public void UpdateKPMDocument(List<Dictionary<string, string>> LTicketItemList)
        {
            DebugPrint("I'm updating Excel File...");

            Excel.Worksheet ws_KPMCreate = g_KPMWB.Worksheets["kpmcreate"];
            int nNumberCol = FindColumn(ws_KPMCreate.get_Range("1:1"), "Number");
            int nUploadCol = FindColumn(ws_KPMCreate.get_Range("1:1"), "Re-upload Attachment");

            int nCurKPMSheetRow = 2;
            for (int nIdx = 0; nIdx < LTicketItemList.Count; nIdx++)
            {
                Dictionary<string, string> TicketItem = LTicketItemList[nIdx];
                string nTicketNum = TicketItem["Number"];
                string Upload = TicketItem["Re-upload Attachment"];

                ((Range)ws_KPMCreate.Cells[nCurKPMSheetRow, nNumberCol]).Value = nTicketNum;
                ((Range)ws_KPMCreate.Cells[nCurKPMSheetRow, nUploadCol]).Value = Upload;
                nCurKPMSheetRow++;
            }

            DebugPrint("I'm saving Excel File...");
            g_KPMWB.Save();
            g_KPMWB = null;
            DebugPrint("I'm closing Excel File...");
            g_KPMExcel.Quit();
            g_KPMExcel = null;
        }
    }
}