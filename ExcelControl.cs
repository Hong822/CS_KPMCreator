using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace CS_KPMCreator
{
    internal class ExcelControl
    {
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

        public void FillDictionary(List<Dictionary<string, string>> TicketItemList, Worksheet ws)
        {
            // Read and setting relevant data
            int nLastRow = LastRowPerColumn(ws.get_Range("C:C"));
            int nLastCol = LastColumnPerRow(ws.get_Range("1:1"));

            for (int RowIdx = 2; RowIdx < nLastRow; RowIdx++)
            {
                Dictionary<string, string> NewItem = new Dictionary<string, string>();
                for (int ColIdx = 2; ColIdx < nLastCol; ColIdx++)
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