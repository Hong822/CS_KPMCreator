using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace CS_KPMCreator
{
    internal class ExcelControl
    {
        RichTextBox g_richTB_Status = null;

        public void SetStatusBox(ref RichTextBox richTB_Status)
        {
            g_richTB_Status = richTB_Status;
        }

        public void ReadExcelValue(System.Windows.Forms.TextBox tExcelPath, RadioButton rbB2B, RadioButton rbB2C, RadioButton rbAudi, RadioButton rbPorsche, ref List<Dictionary<string, string>> LTicketItemList, ref List<Dictionary<string, string>> LActionList)
        {
            g_richTB_Status.Text = "I'm reading KPM Excel file...";

            // ***Setting KPM Items***
            Excel.Application ap = new Excel.Application();
            Excel.Workbook wb = null;
            try
            {
                wb = ap.Workbooks.Open(tExcelPath.Text);
            }
            catch
            {
                MessageBox.Show("Please Select Excel Path.");
            }
            Excel.Worksheet ws_KPMCreate = wb.Worksheets["kpmcreate"];
            ap.Visible = true;

            int nStartRow, nEndRow, nStartCol, nEndCol;
            nStartRow = 2;
            nEndRow = LastRowPerColumn(ws_KPMCreate.get_Range("C:C"));
            nStartCol = 1;
            nEndCol = LastColumnPerRow(ws_KPMCreate.get_Range("1:1"));

            // Fill in ticketItemList with ticket items
            FillDictionary(LTicketItemList, ws_KPMCreate, nStartRow, nEndRow, nStartCol, nEndCol);

            // ***Setting Actions***
            g_richTB_Status.Text = "I'm reading Action Excel file...";

            Excel.Application ActionApp = new Excel.Application();
            string ActionExcelPath = "E:\\VS_Project\\repos\\Hong822\\CS_KPMCreator\\KPM_Action_Description.xlsm";
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
    }
}