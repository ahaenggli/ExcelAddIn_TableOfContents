using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace ExcelAddIn_TableOfContents
{
    class GlobalFunction
    {


        //        'check whether Excel-GUI is german or not
        public static bool isGermanGUI() {
            return true;
            //ToDo: Sprache ermitteln...
        }

   

        //'Does the sheet exists in specific workbook?
        public static bool worksheetExists(Excel.Workbook WB, String sheetToFind ) {
            foreach (Excel.Worksheet Sheet in WB.Worksheets) {
                if(sheetToFind.Equals(Sheet.Name)) return true;                                
            }
                return false;
        }

    }
}
