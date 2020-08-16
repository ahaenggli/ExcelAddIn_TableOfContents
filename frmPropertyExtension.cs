using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using ExcelAddIn_TableOfContents.Properties;

namespace ExcelAddIn_TableOfContents
{
    public partial class frmPropertyExtension : Form
    {

        //      'reference to opend sheet
        Excel.Worksheet ws;

        //'old value of combobox
        private String oldValue;


        public frmPropertyExtension()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form frm = new frmTocSheetExtension();
            frm.Show();
            Close();
        }

        private void PropertyExtensionForm_Load(object sender, EventArgs e)
        {




            oldValue = "";

            if (ws == null) ws = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            if (ws == null) return;

            ws.Parent.Activate();
            ws.Activate();

            cbProperty.Items.Clear();
            txtValue.Text = "";


            //'add existising properties to combobox.list
            foreach (Excel.CustomProperty xx in ws.CustomProperties)
            {
                if (!xx.Name.Equals("isToc")) cbProperty.Items.Add(xx.Name);
            }


            //'add default properties to combobox.list (if they are not already set)
            foreach (String tmp in TocSheetExtension.getTocCustomProperties())
            {
                if (!cbProperty.Items.Contains(tmp) && tmp != "isToc" && !String.IsNullOrWhiteSpace(tmp)) cbProperty.Items.Add(tmp);
            }

            //'add toc properties to combobox.list (if they are not already set)
            foreach (String tmp in TocSheetExtension.getTocColumns())
            {
                if (!cbProperty.Items.Contains(tmp) && tmp != "isToc" && !String.IsNullOrWhiteSpace(tmp)) cbProperty.Items.Add(tmp);
            }


            //'default property is first one defined
            cbProperty.Text = TocSheetExtension.getTocCustomProperties()[0] ?? "";


            //'caption of form (sheet/workbook in it)
            this.Text = "Bearbeite Felder von [" + ws.Parent.Name + "].[" + ws.Name + "]";
            txtValue.Focus();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PropertyExtension.setProperty(ws, cbProperty.Text, txtValue.Text);
            TocSheetExtension.generateTocWorksheet();
            Close();
        }

        private void cbProperty_TextChanged(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(oldValue)) PropertyExtension.setProperty(ws, oldValue, txtValue.Text);
            txtValue.Text = PropertyExtension.getProperty(ws, cbProperty.Text);
            oldValue = cbProperty.Text;
        }
    }
}
