using ExcelAddIn_TableOfContents.Properties;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelAddIn_TableOfContents
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
        }

        private void AboutForm_Load(object sender, EventArgs e)
        {
            lblProdukt.Text = AssemblyProduct;
            label3.Text = AssemblyCopyright;
            label5.Text = String.Format("Version {0}", AssemblyVersion);

            checkBox1.Checked = Settings.Default.IsAutoUpdate;
            txtVersion.Text = Settings.Default.VersionUrl;
            txtDownload.Text = Settings.Default.UpdateUrl;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Settings.Default.IsAutoUpdate = checkBox1.Checked;
            //Settings.Default.VersionUrl = txtVersion.Text;
            //Settings.Default.UpdateUrl = txtDownload.Text;

            Settings.Default.Save();
        }



        #region Assemblyattributaccessoren

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion

        private void lnkGitHub_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ProcessStartInfo website = new ProcessStartInfo(@"https://github.com/ahaenggli/ExcelAddIn_TableOfContents");
            Process.Start(website);
        }

        private void linkLizenz_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            ProcessStartInfo website = new ProcessStartInfo(@"https://github.com/ahaenggli/ExcelAddIn_TableOfContents/blob/master/LICENSE");
            Process.Start(website);
        }

        private void linkSpenden_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ProcessStartInfo website = new ProcessStartInfo(@"https://github.com/ahaenggli/ExcelAddIn_TableOfContents");
            Process.Start(website);
        }
    }
}