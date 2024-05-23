using System;
using System.Runtime.InteropServices;

namespace KDCLibrary
{
    [ComVisible(false)]
    public partial class CreditsForm : DevExpress.XtraEditors.XtraForm
    {
        public CreditsForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/6ebeng/Kurdish-Digital-Calendar");
        }
    }
}
