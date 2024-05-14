using System.Reflection;
using KDCLibrary;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace KDC_PowerPoint
{
    public partial class ThisAddIn
    {
        private Ribbon ribbon;

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Ribbon.AppName = Assembly.GetExecutingAssembly().GetName().Name;
            ribbon = new Ribbon();
            return ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ribbon.PowerPointApp = Globals.ThisAddIn.Application;
            this.Application.PresentationOpen += Application_PresentationOpen;
            this.Application.AfterNewPresentation += Application_AfterNewPresentation;
        }

        private void Application_PresentationOpen(Presentation Pres)
        {
            HandlePresentationLoad(Pres);
        }

        private void Application_AfterNewPresentation(Presentation Pres)
        {
            HandlePresentationLoad(Pres);
        }

        private void HandlePresentationLoad(Presentation Pres)
        {
            if (Ribbon.IsAutoUpdateOnLoadDoc)
                ribbon.UpdateDatesFromCustomXmlPartsForPowerPoint(Pres);
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
