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

            label10.Text =
                "Version "
                + new RegistryHelper().LoadSetting(
                    "Version",
                    "Unknown",
                    "Kurdish Digital Calendar"
                );
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/6ebeng/Kurdish-Digital-Calendar");
        }
    }
}
