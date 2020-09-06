using ExcelAddIn_TableOfContents.Properties;
using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddIn_TableOfContents
{
    class TocSheetExtension
    {

        public static void generateTocWorksheet()
        {
            Excel.Worksheet Newsh;
            Excel.Workbook Basebook;
            Excel.Worksheet Basesheet;

            int RwNum, ColNum;

            string[] TocColumns = { };
            string TableStyle = null;
            string TocSheetName = null;

            Globals.ThisAddIn.Application.ActiveWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            Globals.ThisAddIn.Application.ActiveWorkbook.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.ActiveWorkbook.Application.DisplayAlerts = false;

            Basebook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Basesheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            TocSheetName = getTocSheetName();
            TocColumns = getTocColumns();

            if (!GlobalFunction.worksheetExists(Basebook, TocSheetName))
            {
                Newsh = Basebook.Worksheets.Add(Basebook.Worksheets[1]);
                Newsh.Name = TocSheetName;
            }
            else
            {
                Newsh = Basebook.Worksheets[TocSheetName];
            }


            if (Newsh.ListObjects.Count > 0)
            {
                TableStyle = (string)Newsh.ListObjects[1].TableStyle.Name;
                Newsh.ListObjects[1].Delete();
            }


            Newsh.Cells.Clear();
            Newsh.Cells.Delete();


            setTocSheetFlag(Newsh);


            Globals.ThisAddIn.Application.ActiveWorkbook.Application.DisplayAlerts = true;



            // Add headers
            for (int i = 1; i < TocColumns.Length + 1; i++)
            {
                Newsh.Cells[1, i].Value = TocColumns[i - 1];
                Newsh.Cells[1, i].Font.Bold = true;
                Newsh.Cells[1, i].Font.Size = 12;
            }



            RwNum = 1;

            foreach (Excel.Worksheet Sh in Basebook.Worksheets)
            {
                if (!Sh.Name.Equals(Newsh.Name) && Sh.Visible == XlSheetVisibility.xlSheetVisible)
                {
                    ColNum = 1;
                    RwNum = RwNum + 1;

                    //  'Create a link to the sheet in the A column
                    Newsh.Hyperlinks.Add(Newsh.Cells[RwNum, 1], "", "'" + Sh.Name + "'!A1", "", Sh.Name);

                    foreach (String col in TocColumns)
                    {
                        if (!String.IsNullOrWhiteSpace(col) && !col.Equals(TocColumns[0]))
                        {
                            ColNum = ColNum + 1;
                            Newsh.Cells[RwNum, ColNum] = PropertyExtension.getProperty(Sh, col);
                        }
                    }


                }
            }

            Excel.ListObject tbl;
            Excel.Range trng;

            trng = Newsh.UsedRange;
            tbl = Newsh.ListObjects.Add(XlListObjectSourceType.xlSrcRange, trng, null, XlYesNoGuess.xlYes);
            tbl.TableStyle = TableStyle ?? Settings.Default.TocStyle;
            tbl.Name = TocSheetName;
            trng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            trng.Borders.Weight = Excel.XlBorderWeight.xlThin;
            trng.Borders.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;


            Newsh.UsedRange.ColumnWidth = 250;
            Newsh.UsedRange.RowHeight = 250;
            Newsh.UsedRange.HorizontalAlignment = Excel.Constants.xlLeft;
            Newsh.UsedRange.VerticalAlignment = Excel.Constants.xlTop;


            Newsh.UsedRange.Columns.AutoFit();

            foreach (Excel.Range rng in Newsh.UsedRange.Columns)
            {
                if (rng.ColumnWidth > 75)
                {
                    rng.ColumnWidth = 75;
                    rng.WrapText = true;
                }
            }


            Newsh.UsedRange.Rows.AutoFit();

            Globals.ThisAddIn.Application.ActiveWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Globals.ThisAddIn.Application.ActiveWorkbook.Application.ScreenUpdating = true;


            Basebook.Activate();
            Basesheet.Activate();
        }



        // 'get the name for the worksheet created field (custom property)
        public static String getWorksheetCreatedDatePropName()
        {
            String prop = "";

            try
            {
                if (String.IsNullOrWhiteSpace(prop)) prop = PropertyExtension.getProperty(getTocSheet(), "WorksheetCreatedDatePropName") ?? Settings.Default.WorksheetCreatedDatePropName;
                if (String.IsNullOrWhiteSpace(prop)) prop = "Datum";
                //if (String.IsNullOrWhiteSpace(prop) && !GlobalFunction.isGermanGUI()) prop = "Created";
            }
            catch (System.Exception e)
            {
                prop = "Created";
            }

            return prop;
        }

        //'get name of properties which are shown in the Toc sheet
        public static string[] getTocColumns()
        {

            String props = "";

            try
            {

                if (String.IsNullOrWhiteSpace(props)) props = PropertyExtension.getProperty(getTocSheet(), "TocColumns") ?? Settings.Default.TocColumns;
                if (String.IsNullOrWhiteSpace(props)) props = "Blatt;Datum;Beschreibung;Verantwortlich;ToDo;Status;Info";
                //if (String.IsNullOrWhiteSpace(props) && !GlobalFunction.isGermanGUI()) props = "Worksheet;Created;Description;Responsible;ToDo;Status;Info";

            }
            catch (System.Exception e)
            {

                props = "Worksheet;Created;Description;Responsible;ToDo;Status;Info";

            }

            //'first column has to be for the hyperlink to the other worksheets, first array entry should not be an existing custom property
            if (getTocCustomProperties().Contains(props.Split(';')[0]) || props.Split(';')[0].Equals(getWorksheetCreatedDatePropName()))
            {
                props = ";" + props;
            }

            return props.Split(';');
        }

        //'get name of custom proprties which should be created in all worksheets
        public static string[] getTocCustomProperties()
        {

            String props = "";

            try
            {

                if (String.IsNullOrWhiteSpace(props)) props = PropertyExtension.getProperty(getTocSheet(), "TocCustomProperties") ?? Settings.Default.TocCustomProperties;
                if (String.IsNullOrWhiteSpace(props)) props = "Beschreibung;Verantwortlich;ToDo;Status;Info;Datum";
                //if (String.IsNullOrWhiteSpace(props) && !GlobalFunction.isGermanGUI()) props = "Description;Responsible;ToDo;Status;Info;Created";

            }
            catch (System.Exception e)
            {

                props = "Description;Responsible;ToDo;Status;Info;Created";

            }

            return props.Split(';');
        }

        //set flag for Toc sheet
        public static void setTocSheetFlag(Excel.Worksheet ws)
        {
            Excel.Workbook ActiveWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            foreach (Excel.Worksheet Sheet in ActiveWorkbook.Worksheets)
            {
                PropertyExtension.setProperty(Sheet, "isToc", "0");
            }

            PropertyExtension.setProperty(ws, "isToc", "1");
        }



        //'get the defined name for the Toc worksheet
        public static String getTocSheetName()
        {
            Excel.Workbook ActiveWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            String sumsheet = "";


            foreach (Excel.Worksheet ws in ActiveWorkbook.Worksheets)
            {
                if (PropertyExtension.getProperty(ws, "isToc") == "1")
                {
                    return ws.Name;
                }
            }

            if (String.IsNullOrWhiteSpace(sumsheet)) sumsheet = Settings.Default.TocWorksheetName;
            if (String.IsNullOrWhiteSpace(sumsheet)) sumsheet = "Uebersicht";
            //if (String.IsNullOrWhiteSpace(sumsheet) && !GlobalFunction.isGermanGUI()) sumsheet = "Toc";

            return sumsheet;
        }



        //'returns ref to Toc sheet if exists, else nothing
        public static Excel.Worksheet getTocSheet()
        {

            Excel.Workbook ActiveWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            String idx = getTocSheetName();
            if (GlobalFunction.worksheetExists(ActiveWorkbook, idx))
            {
                return ActiveWorkbook.Worksheets[idx];
            }
            else
                return null;

        }

    }



}
