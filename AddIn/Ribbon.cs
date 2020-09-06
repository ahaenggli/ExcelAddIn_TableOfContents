using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn_TableOfContents
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility-Member

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelAddIn_TableOfContents.Ribbon.xml");
        }

        #endregion

        #region Menübandrückrufe
        //Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.

        public void OnInfo(Office.IRibbonControl control)
        {
            Form frm = new AboutForm();
            frm.ShowDialog();
            frm.Close();
            frm.Dispose();
        }

        public void OnWertChange(Office.IRibbonControl control)
        {
            Form frm = new frmPropertyExtension();
            frm.ShowDialog();
            frm.Close();
            frm.Dispose();
        }

        public void OnSettings(Office.IRibbonControl control)
        {
            Form frm = new frmTocSheetExtension();
            frm.ShowDialog();
            frm.Close();
            frm.Dispose();
        }

        public void OnGenerieren(Office.IRibbonControl control)
        {
            TocSheetExtension.generateTocWorksheet();
        }


        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Hilfsprogramme

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
