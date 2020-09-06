using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn_TableOfContents
{
    class GlobalFunction
    {


        //'Does the sheet exists in specific workbook?
        public static bool worksheetExists(Excel.Workbook WB, String sheetToFind)
        {
            foreach (Excel.Worksheet Sheet in WB.Worksheets)
            {
                if (sheetToFind.Equals(Sheet.Name)) return true;
            }
            return false;
        }

    }
}
