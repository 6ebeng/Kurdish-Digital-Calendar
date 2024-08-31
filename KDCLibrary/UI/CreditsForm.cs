using System;
using System.Runtime.InteropServices;
using KDCLibrary.Helpers;

namespace KDCLibrary
{
    [ComVisible(false)]
    public partial class CreditsForm : DevExpress.XtraEditors.XtraForm
    {
        public CreditsForm()
        {
            InitializeComponent();
            string version = new RegistryHelper().GetRegistryValueFromX6432Path(
                @"SOFTWARE\Rekbin Devs\Kurdish Digital Calendar",
                "Version"
            );
            label10.Text = "Version " + version;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/6ebeng/Kurdish-Digital-Calendar");
        }
    }
}
