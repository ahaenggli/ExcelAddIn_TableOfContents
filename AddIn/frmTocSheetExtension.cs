using ExcelAddIn_TableOfContents.Properties;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn_TableOfContents
{
    public partial class frmTocSheetExtension : Form
    {
        public frmTocSheetExtension()
        {
            InitializeComponent();
        }

        private void TocSheetExtensionForm_Load(object sender, EventArgs e)
        {
            Excel.Workbook ActiveWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            txtSumTitel.Text = TocSheetExtension.getTocSheetName();
            txtProperties.Text = String.Join(";", TocSheetExtension.getTocCustomProperties());
            txtSummaryColumns.Text = String.Join(";", TocSheetExtension.getTocColumns());
            txtWorkSheetCreatedDate.Text = TocSheetExtension.getWorksheetCreatedDatePropName();
            txtStyle.Text = Settings.Default.TocStyle;

            if (!GlobalFunction.worksheetExists(ActiveWorkbook, TocSheetExtension.getTocSheetName()))
            {
                cbSetDefault.Checked = true;
            }

        }

        private void OK_Click(object sender, EventArgs e)
        {
            Excel.Workbook ActiveWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (cbSetDefault.Checked)
            {
                //'save "global"-properties in ThisWorkbook.Worksheets(1)
                //' -> ThisWorkbook is where the code is saved (xlam-file)
                //' -> even a xlam file has at least one sheet
                //' -> here it's named "TocConfig"
                Settings.Default.TocWorksheetName = txtSumTitel.Text;
                Settings.Default.TocCustomProperties = txtProperties.Text;
                Settings.Default.TocColumns = txtSummaryColumns.Text;
                Settings.Default.WorksheetCreatedDatePropName = txtWorkSheetCreatedDate.Text;
                Settings.Default.TocStyle = txtStyle.Text;
                Settings.Default.Save();
            }

            if (!GlobalFunction.worksheetExists(ActiveWorkbook, TocSheetExtension.getTocSheetName()))
            {
                TocSheetExtension.generateTocWorksheet();
            }

            if (GlobalFunction.worksheetExists(ActiveWorkbook, TocSheetExtension.getTocSheetName()))
            {
                Excel.Worksheet toc = TocSheetExtension.getTocSheet();
                PropertyExtension.setProperty(toc, "TocWorksheetName", txtSumTitel.Text);
                PropertyExtension.setProperty(toc, "TocCustomProperties", txtProperties.Text);
                PropertyExtension.setProperty(toc, "TocColumns", txtSummaryColumns.Text);
                PropertyExtension.setProperty(toc, "WorksheetCreatedDatePropName", txtWorkSheetCreatedDate.Text);
                if (!toc.Name.Equals(txtSumTitel.Text)) toc.Name = TocSheetExtension.getTocSheetName();
            }

            Close();
        }
    }
}