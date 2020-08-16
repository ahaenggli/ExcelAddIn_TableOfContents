﻿using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ExcelAddIn_TableOfContents.Properties;
using System.IO;
using System.Net;
using System.Threading;
using System.IO.Compression;

namespace ExcelAddIn_TableOfContents
{
    public partial class ThisAddIn
    {


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            this.Application.WorkbookNewSheet += Application_WorkbookNewSheet;
            this.Application.SheetActivate += Application_SheetActivate;

            //event and property do have the same name ....
            ((Excel.AppEvents_Event)this.Application).NewWorkbook += Application_NewWorkbook;

            //event for changes in TOC-Sheet
            this.Application.SheetChange += Application_SheetChange;


            Thread worker = new Thread(Update);
            worker.IsBackground = true;
            worker.Start();
        }

        private void Application_SheetChange(object Sh, Range Target)
        {
            // no worksheet
            if (Sh == null) return;
            // still no worksheet
            if (!(Sh is Excel.Worksheet)) return;
            // something is strage ... 
            if (!(((Excel.Worksheet)Sh).Parent is Excel.Workbook)) return;
            // no table, not data
            if (((Excel.Worksheet)Sh).ListObjects.Count == 0) return;
            // isTOC
            if (!PropertyExtension.getProperty(((Excel.Worksheet)Sh), "isToc").Equals("1")) return;


            //If Intersect(Target, Sh.Range(Sh.ListObjects(1).Range.Address)) Is Nothing And Sh.ListObjects(1).ListColumns.count = UBound(arrIdxCols) + 1 Then Exit Sub
            Excel.Worksheet ws = (Excel.Worksheet)Sh;
            string[] arrIdxCols = TocSheetExtension.getTocColumns();
            string[] arrCusProp = TocSheetExtension.getTocCustomProperties();
            if (arrIdxCols.Length <= 0) return;
            if (arrCusProp.Length <= 0) return;

            //'not in table-area (updated cell) and (number of table-columns is equal Toc-array)
            if ((Application.Intersect(Target, ws.Range[ws.ListObjects[1].Range.Address]) == null) && (ws.ListObjects[1].ListColumns.Count == arrIdxCols.Length)) return;

            string[] arrTblCols = ws.ListObjects[1].HeaderRowRange.Cells.Cast<Excel.Range>().Select(Selector).ToArray<string>();

            /* track all the changes... */
            // 'delete Toc props in available sheets
            //Dim rw As Range
            //Dim vl As Variant
            foreach (Excel.Range rw in ws.ListObjects[1].Range.Rows)
            {
                if (GlobalFunction.worksheetExists(ws.Parent, rw.Columns[1].Value))
                {

                    foreach (String missing in arrIdxCols.Where(x => !arrTblCols.Contains(x)))
                    {
                        PropertyExtension.setProperty(ws.Parent.Worksheets(rw.Columns[1].Value), missing, "");
                    }

                    foreach (Excel.Range cl in ws.ListObjects[1].HeaderRowRange.Cells)
                    {
                        if (cl.Column == 1) continue;
                        PropertyExtension.setProperty(ws.Parent.Worksheets(rw.Columns[1].Value), cl.Value, rw.Columns[cl.Column].Value);
                    }
                }
            }

            // war in altem idx, neu aber nicht mehr
            string[] newCusProp = arrCusProp.Where(x => !(arrIdxCols.Contains(x) && !arrTblCols.Contains(x))).ToArray();
            string str_arrTblCols = String.Join(";", arrTblCols);
            string str_newCusProp = String.Join(";", newCusProp);

            PropertyExtension.setProperty(ws, "TocColumns", str_arrTblCols);
            PropertyExtension.setProperty(ws, "TocCustomProperties", str_newCusProp);

        }

        public string Selector(Excel.Range cell)
        {
            if (cell.Value2 == null)
                return "";
            else return cell.Value2;
            //if (cell.Value2.GetType().ToString() == "System.Double")
            //    return ((double)cell.Value2).ToString();
            //else if (cell.Value2.GetType().ToString() == "System.String")
            //    return ((string)cell.Value2);
            //else if (cell.Value2.GetType().ToString() == "System.Boolean")
            //    return ((bool)cell.Value2).ToString();
            //else
            //    return "unknown";
        }

        private void Application_NewWorkbook(Excel.Workbook Wb)
        {
            if (!(Wb.Sheets[1] is Excel.Worksheet)) return;
            String cPrpNm = TocSheetExtension.getWorksheetCreatedDatePropName();
            if (!String.IsNullOrEmpty(cPrpNm)) PropertyExtension.setProperty(Wb.Sheets[1], cPrpNm, DateTime.Now.ToString());
            PropertyExtension.setProperty(Wb.Sheets[1], "isToc", "0");
        }



        private void Application_WorkbookNewSheet(Excel.Workbook Wb, object Sh)
        {
            if (!(Sh is Excel.Worksheet)) return;
            String cPrpNm = TocSheetExtension.getWorksheetCreatedDatePropName();
            if (!String.IsNullOrEmpty(cPrpNm)) PropertyExtension.setProperty((Excel.Worksheet)Sh, cPrpNm, DateTime.Now.ToString());
            PropertyExtension.setProperty((Excel.Worksheet)Sh, "isToc", "0");
        }

        private void Application_SheetActivate(object Sh)
        {
            if (!(Sh is Excel.Worksheet)) return;

            if (((Excel.Worksheet)Sh).Name.Equals(TocSheetExtension.getTocSheetName()))
                TocSheetExtension.generateTocWorksheet();

        }
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {

        }


        /*

Option Explicit

'event handler for application
Private TocSheetExtension_AppEventHandler As TocSheetAppEventHandler

' Sub is called whenever AddIn is loaded
Private Sub Workbook_Open()
 'To combine keys with    Precede the key code by
 'SHIFT   + (plus sign)
 'CTRL    ^ (caret)
 'ALT     % (percent sign)

 '' CTRL + Shift + A
 'Application.OnKey "^+{A}", "tstBox"
  Application.OnKey "{F5}", "handleF5Click"
  isF5 = False
 'init application wide event handler
  Set TocSheetExtension_AppEventHandler = New TocSheetAppEventHandler

End Sub

          */

        // handles click on F5-Key
        private void handleF5Click()
        {

            Excel.Workbook ActiveWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (ActiveWorkbook == null) return;


            if (!ActiveWorkbook.ActiveSheet.Name.Equals(TocSheetExtension.getTocSheetName()))
            {
                Form frm = new frmPropertyExtension();
                frm.Show();
            }
            else
                TocSheetExtension.generateTocWorksheet();



        }



        private void Update()
        {
            try
            {
                if (Settings.Default.LastUpdateCheck == null)
                {
                    Settings.Default.LastUpdateCheck = DateTime.Now;
                    Properties.Settings.Default.Save();
                }

                // once a day should be enougth....
                if (Settings.Default.LastUpdateCheck.AddMinutes(2) <= DateTime.Now)
                {

                    string ProgramData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\haenggli.NET\";
                    string AddInData = ProgramData + @"ExcelAddIn_TableOfContents\";
                    string StartFile = AddInData + @"TableOfContents.vsto";
                    string localFile = AddInData + @"TableOfContents.zip";
                    string DownloadUrl = Environment.GetEnvironmentVariable("TableOfContents_DownloadUrl", EnvironmentVariableTarget.Machine) ?? Settings.Default.UpdateUrl;

                    if (DownloadUrl.Equals("---"))
                    {
                        Settings.Default.LastUpdateCheck = DateTime.Now;
                        Properties.Settings.Default.Save();
                        return;
                    }

                    if (!Directory.Exists(AddInData)) Directory.CreateDirectory(AddInData);
                    foreach (System.IO.FileInfo file in new DirectoryInfo(AddInData).GetFiles()) file.Delete();
                    foreach (System.IO.DirectoryInfo subDirectory in new DirectoryInfo(AddInData).GetDirectories()) subDirectory.Delete(true);

                    if (DownloadUrl.StartsWith("http"))
                    {
                        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                        WebClient webClient = new WebClient();
                        webClient.DownloadFile(DownloadUrl, localFile);
                        webClient.Dispose();
                        DownloadUrl = localFile;
                    }

                    ZipFile.ExtractToDirectory(DownloadUrl, AddInData);

                    Settings.Default.LastUpdateCheck = DateTime.Now;
                    Properties.Settings.Default.Save();
                }
            }
            catch (System.Exception Ex)
            {
                MessageBox.Show(Ex.Message);
            }
        }


        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }



}
