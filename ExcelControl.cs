using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace CS_KPMCreator
{
    internal class ExcelControl
    {
        private RichTextBox g_richTB_Status = null;
        private Excel.Application g_KPMExcelApp = null;
        private Excel.Workbook g_KPMWorkbook = null;
        private Excel.Worksheet g_KPMCreate_Worksheet = null;

        public void SetStatusBox(ref RichTextBox richTB_Status)
        {
            g_richTB_Status = richTB_Status;
        }

        public void CloseExcelControl()
        {
            if (g_KPMCreate_Worksheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(g_KPMCreate_Worksheet);
            }
            if (g_KPMWorkbook != null)
            {
                DebugPrint("I'm saving Excel File...");
                if (g_KPMWorkbook.ReadOnly == false)
                {
                    g_KPMWorkbook.Save();
                }
                g_KPMWorkbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(g_KPMWorkbook);
                g_KPMWorkbook = null;
            }
            if (g_KPMExcelApp != null)
            {
                DebugPrint("I'm closing Excel File...");
                g_KPMExcelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(g_KPMExcelApp);
                g_KPMExcelApp = null;
            }
        }

        public void CloseExcelControl(ref Excel.Worksheet ws, ref Excel.Workbook wb, ref Excel.Application app)
        {
            if (ws != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            }
            if (wb != null)
            {
                DebugPrint("I'm saving Excel File...");
                if (wb.ReadOnly == false)
                {
                    wb.Save();
                }
                wb.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                wb = null;
            }
            if (app != null)
            {
                DebugPrint("I'm closing Excel File...");
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
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

        public bool ReadExcelValue(System.Windows.Forms.TextBox tExcelPath, RadioButton rbB2B, RadioButton rbB2C, RadioButton rbAudi, RadioButton rbPorsche, RadioButton rbKPMRead, ref List<Dictionary<string, string>> LTicketItemList, ref List<Dictionary<string, string>> LActionList)
        {
            DebugPrint("I'm reading KPM Excel File...");

            // ***Setting KPM Items***
            g_KPMExcelApp = new Excel.Application();
            try
            {
                //tExcelPath.Text = "E:\\VS_Project\\repos\\Hong822\\CS_KPMCreator\\KPM_Ticket_Creator_V1.xlsm";
                g_KPMWorkbook = g_KPMExcelApp.Workbooks.Open(tExcelPath.Text);
                if (g_KPMWorkbook.ReadOnly == true && rbKPMRead.Checked == false)
                {
                    DebugPrint("Please Close KPM Excel File and Open with Write Autority...");
                    CloseExcelControl(ref g_KPMCreate_Worksheet, ref g_KPMWorkbook, ref g_KPMExcelApp);
                    return false;
                }
            }
            catch
            {
                DebugPrint("Please Select Excel Path.");
                CloseExcelControl(ref g_KPMCreate_Worksheet, ref g_KPMWorkbook, ref g_KPMExcelApp);
                return false;
            }

            g_KPMCreate_Worksheet = g_KPMWorkbook.Worksheets["KPM_Create"];
            g_KPMExcelApp.Visible = true;
            //g_KPMExcelApp.Visible = false;

            int nStartRow, nEndRow, nStartCol, nEndCol;
            nStartRow = 2;

            //int STCol = FindColumn(g_KPMCreate_Worksheet.get_Range("1:1"), "Short Text");
            //string tempRange = STCol + ":" + STCol;
            //nEndRow = LastRowPerColumn(g_KPMCreate_Worksheet.get_Range(tempRange));
            nEndRow = LastRowPerColumn(g_KPMCreate_Worksheet.get_Range("K:K"));

            nStartCol = 1;
            nEndCol = LastColumnPerRow(g_KPMCreate_Worksheet.get_Range("1:1"));

            // Fill in ticketItemList with ticket items
            FillDictionary(LTicketItemList, g_KPMCreate_Worksheet, nStartRow, nEndRow, nStartCol, nEndCol);

            // ***Setting Actions***
            DebugPrint("I'm reading Action Excel File...");

            Excel.Application ActionApp = new Excel.Application();
            Excel.Workbook ActionWB = null;
            Excel.Worksheet ActionWS = null;

            string ActionExcelPath = "C:\\KPM_Creator\\KPM_Action_Description.xlsm";
            //string ActionExcelPath = "E:\\VS_Project\\repos\\Hong822\\CS_KPMCreator\\KPM_Action_Description.xlsm";
            //string ActionExcelPath = "D:\\25_C_Projects\\Repos\\Hong822\\CS_KPMCreator\\KPM_Action_Description.xlsm";
            try
            {
                ActionWB = ActionApp.Workbooks.Open(ActionExcelPath);
            }
            catch
            {
                DebugPrint("Please check " + ActionExcelPath + ".");
                CloseExcelControl(ref ActionWS, ref ActionWB, ref ActionApp);
                return false;
            }

            string ActionSheet = "";
            if (rbKPMRead.Checked == true)
            {
                ActionSheet = "KPMRead";
            }
            else
            {
                if (rbPorsche.Checked == true)
                {
                    if (rbB2B.Checked == true)
                    {
                        ActionSheet = "PO_B2B";
                    }
                    else
                    {
                        ActionSheet = "PO_B2C";
                    }
                }
                else    // Audi
                {
                    if (rbB2B.Checked == true)
                    {
                        ActionSheet = "AU_B2B";
                    }
                    else
                    {
                        ActionSheet = "AU_B2C";
                    }
                }
            }

            ActionWS = ActionWB.Worksheets[ActionSheet];
            DebugPrint("I'm reading Action Excel File...\t" + ActionSheet);

            ActionApp.Visible = false;
            //ActionApp.Visible = true;

            // Fill in ticketItemList with ticket items
            nStartRow = 2;
            //STCol = FindColumn(ActionWS.get_Range("1:1"), "Step");
            //tempRange = STCol + ":" + STCol;
            //nEndRow = LastRowPerColumn(ActionWS.get_Range(tempRange));
            nEndRow = LastRowPerColumn(ActionWS.get_Range("A:A"));
            nStartCol = 1;
            nEndCol = LastColumnPerRow(ActionWS.get_Range("1:1"));
            FillDictionary(LActionList, ActionWS, nStartRow, nEndRow, nStartCol, nEndCol);

            if (ActionApp.Visible == false)
            {
                CloseExcelControl(ref ActionWS, ref ActionWB, ref ActionApp);
            }
            else
            {
                // TODO: Close Excel Objects
            }

            return true;
        }

        public void UpdateKPMDocument(List<Dictionary<string, string>> LTicketItemList)
        {
            DebugPrint("I'm updating Excel File...");

            int nNumberCol = FindColumn(g_KPMCreate_Worksheet.get_Range("1:1"), "Number");
            int nUploadCol = FindColumn(g_KPMCreate_Worksheet.get_Range("1:1"), "Re-upload Attachment");

            int nCurKPMSheetRow = 2;
            for (int nIdx = 0; nIdx < LTicketItemList.Count; nIdx++)
            {
                Dictionary<string, string> TicketItem = LTicketItemList[nIdx];
                string nTicketNum = TicketItem["Number"];
                string Upload = TicketItem["Re-upload Attachment"];

                ((Range)g_KPMCreate_Worksheet.Cells[nCurKPMSheetRow, nNumberCol]).Value = nTicketNum;
                ((Range)g_KPMCreate_Worksheet.Cells[nCurKPMSheetRow, nUploadCol]).Value = Upload;
                nCurKPMSheetRow++;
            }

            DebugPrint("I'm saving Excel File...");
            if (g_KPMWorkbook.ReadOnly == false)
            {
                g_KPMWorkbook.Save();
            }
        }

        public void UpdateKPMReadSheet(List<Dictionary<KPMReadInfo, List<string>>> ReadList)
        {
            DebugPrint("I'm updating KPM Read Excel Sheet...");


            Worksheet KPMSelection_Worksheet = g_KPMWorkbook.Worksheets["KPM_Selection"];

            int nCurRow = 3;
            for (int nIdx = 0; nIdx < ReadList.Count; nIdx++)
            {
                Dictionary<KPMReadInfo, List<string>> tempDic = ReadList[nIdx];
                List<KPMReadInfo> Keys = new List<KPMReadInfo>(tempDic.Keys);
                string sFunctionName = Keys[0].sDataType;

                int nCol = FindColumn(KPMSelection_Worksheet.get_Range("1:1"), sFunctionName);

                List<string> tempValueList = tempDic[Keys[0]];

                int nInfoDepth = Keys[0].nDepthCnt;
                for (int nCurDepth = 0; nCurDepth <= nInfoDepth; nCurDepth++)
                {
                    if (nCurDepth == nInfoDepth)
                    {
                        foreach (string curString in tempValueList)
                        {
                            ((Range)KPMSelection_Worksheet.Cells[nCurRow, nCol + nCurDepth]).Value = curString;
                            nCurRow++;
                        }
                    }
                    else
                    {
                        string DepthString = "";
                        if (nCurDepth == 0)
                        {
                            DepthString = Keys[0].Depth1;
                        }
                        else if (nCurDepth == 1)
                        {
                            DepthString = Keys[0].Depth2;
                        }
                        else if (nCurDepth == 2)
                        {
                            DepthString = Keys[0].Depth3;
                        }
                        ((Range)KPMSelection_Worksheet.Cells[nCurRow, nCol + nCurDepth]).Value = DepthString;
                    }
                }
            }

            if (g_KPMWorkbook.ReadOnly == false)
            {
                g_KPMWorkbook.Save();
            }
        }
    }
}