using System.Reflection;
using KDCLibrary;
using Office = Microsoft.Office.Core;

namespace KDC_Viso
{
    public partial class ThisAddIn
    {
        private Ribbon ribbon;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Ribbon.AppName = Assembly.GetExecutingAssembly().GetName().Name;
            ribbon = new Ribbon();
            return ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Ribbon.VisioApp = Globals.ThisAddIn.Application;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
