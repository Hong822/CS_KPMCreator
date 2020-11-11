using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace CS_KPMCreator
{
    internal class ExcelControl
    {
        private Util g_Util = null;
        private Excel.Application g_KPMExcelApp = null;
        private Excel.Workbook g_KPMWorkbook = null;
        private Excel.Worksheet g_KPMCreate_Worksheet = null;

        public ExcelControl(ref Util util)
        {
            g_Util = util;
        }

        public void CloseExcelControl(ref List<Process> processes)
        {
            CloseExcelControl(ref g_KPMCreate_Worksheet, ref g_KPMWorkbook, ref g_KPMExcelApp, ref processes);
        }

        public void CloseExcelControl(ref Excel.Worksheet ws, ref Excel.Workbook wb, ref Excel.Application app, ref List<Process> processes)
        {
            if (ws != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            }
            if (wb != null)
            {
                g_Util.DebugPrint("I'm saving Excel File...");
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
                g_Util.DebugPrint("I'm closing Excel File...");
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
            }

            if (processes != null && processes.Count > 0)
            {
                g_Util.DebugPrint("I'm killing Excel processes...");
                foreach (Process Iter in processes)
                {
                    if (Iter.ProcessName == "EXCEL")
                    {
                        try
                        {
                            Process.GetProcessById(Iter.Id);
                            Iter.Kill();
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
            }
        }

        public void FindNewProcess(Process[] processesBefore, Process[] processesAfter, ref List<Process> processes)
        {
            foreach (Process AfterIter in processesAfter)
            {
                bool bFind = false;
                foreach (Process BeforeIter in processesBefore)
                {
                    if (AfterIter.Id == BeforeIter.Id)
                    {
                        bFind = true;
                        break;
                    }
                }
                if (bFind == false)
                {
                    processes.Add(AfterIter);
                }
            }
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

        public bool ReadExcelValue(System.Windows.Forms.TextBox tExcelPath, RadioButton rbB2B, RadioButton rbB2C, RadioButton rbKPMRead, RadioButton rbTKCancel, ref List<Dictionary<string, string>> LTicketItemList, ref List<Dictionary<string, string>> LActionList, ref List<Process> processes)
        {
            g_Util.DebugPrint("I'm reading KPM Excel File...");

            Process[] processesBefore = Process.GetProcessesByName("EXCEL");

            // ***Setting KPM Items***
            g_KPMExcelApp = new Excel.Application();
            try
            {
                //Hard coding
                //tExcelPath.Text = "E:\\VS_Project\\repos\\Hong822\\CS_KPMCreator\\KPM_Ticket_Creator_V1.xlsm";
                //tExcelPath.Text = "D:\\25_C_Projects\\Repos\\Hong822\\CS_KPMCreator\\KPM_Ticket_Creator_V1.xlsm";

                g_KPMWorkbook = g_KPMExcelApp.Workbooks.Open(tExcelPath.Text);

                Process[] processesAfter = Process.GetProcessesByName("EXCEL");
                FindNewProcess(processesBefore, processesAfter, ref processes);

                if (g_KPMWorkbook.ReadOnly == true && rbKPMRead.Checked == false)
                {
                    g_Util.DebugPrint("Please Close KPM Excel File and Open with Write Autority...");
                    CloseExcelControl(ref g_KPMCreate_Worksheet, ref g_KPMWorkbook, ref g_KPMExcelApp, ref processes);
                    return false;
                }
            }
            catch
            {
                g_Util.DebugPrint("Please Select Excel Path.");
                CloseExcelControl(ref g_KPMCreate_Worksheet, ref g_KPMWorkbook, ref g_KPMExcelApp, ref processes);
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
            nEndRow = LastRowPerColumn(g_KPMCreate_Worksheet.get_Range("H:H"));

            nStartCol = 1;
            nEndCol = LastColumnPerRow(g_KPMCreate_Worksheet.get_Range("1:1"));

            // Fill in ticketItemList with ticket items
            FillDictionary(LTicketItemList, g_KPMCreate_Worksheet, nStartRow, nEndRow, nStartCol, nEndCol);

            // ***Setting Actions***
            g_Util.DebugPrint("I'm reading Action Excel File...");

            Excel.Application ActionApp = new Excel.Application();
            Excel.Workbook ActionWB = null;
            Excel.Worksheet ActionWS = null;

            var Dir = System.IO.Directory.GetCurrentDirectory();
            //Dir = Dir.Substring(0, Dir.LastIndexOf("\\"));

            //Hard coding
            string ActionExcelPath = Dir + "\\KPM_Action_Description.xlsm";
            //string ActionExcelPath = "E:\\VS_Project\\repos\\Hong822\\CS_KPMCreator\\KPM_Action_Description.xlsm";
            //ActionExcelPath = "D:\\25_C_Projects\\Repos\\Hong822\\CS_KPMCreator\\bin\\Debug\\KPM_Action_Description.xlsm";

            try
            {
                ActionWB = ActionApp.Workbooks.Open(ActionExcelPath);
                Process[] processesAfter = Process.GetProcessesByName("EXCEL");
                FindNewProcess(processesBefore, processesAfter, ref processes);
            }
            catch
            {
                g_Util.DebugPrint("Please check " + ActionExcelPath + ".");
                CloseExcelControl(ref ActionWS, ref ActionWB, ref ActionApp, ref processes);
                return false;
            }

            string ActionSheet = "";
            if (rbKPMRead.Checked == true)
            {
                ActionSheet = "KPMRead";
            }
            else if (rbTKCancel.Checked == true)
            {
                ActionSheet = "KPMDelete";
            }
            else
            {
                ActionSheet = (rbB2B.Checked == true) ? "B2B": "B2C"; 
            }

            ActionWS = ActionWB.Worksheets[ActionSheet];
            g_Util.DebugPrint("I'm reading Action Excel File...\t" + ActionSheet);

            //ActionApp.Visible = false;
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
                List<Process> dummyList = null;
                CloseExcelControl(ref ActionWS, ref ActionWB, ref ActionApp, ref dummyList);
            }
            else
            {
                // TODO: Close Excel Objects
            }

            return true;
        }

        public void UpdateKPMDocument(List<Dictionary<string, string>> LTicketItemList)
        {
            g_Util.DebugPrint("I'm updating Excel File...");

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

            g_Util.DebugPrint("I'm saving Excel File...");
            if (g_KPMWorkbook.ReadOnly == false)
            {
                g_KPMWorkbook.Save();
            }
        }

        public void UpdateKPMReadSheet(List<Dictionary<KPMReadInfo, List<string>>> ReadList)
        {
            g_Util.DebugPrint("I'm updating KPM Read Excel Sheet...");

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