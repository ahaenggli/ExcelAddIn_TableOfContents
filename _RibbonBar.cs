using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn_TableOfContents
{
    public partial class _RibbonBar
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

            Form frm = new frmPropertyExtension();
            frm.Show();

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Form frm = new frmTocSheetExtension();
            frm.Show();

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            TocSheetExtension.generateTocWorksheet();
        }
    }
}
