using System.Reflection;
using KDCLibrary;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace KDC_Excel
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
            ribbon.ExcelApp = Globals.ThisAddIn.Application;
            // Explicitly cast Application to AppEvents_Event to resolve ambiguity
            ((Excel.AppEvents_Event)this.Application).WorkbookOpen +=
                new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            ((Excel.AppEvents_Event)this.Application).NewWorkbook +=
                new Excel.AppEvents_NewWorkbookEventHandler(Application_NewWorkbook);
        }

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            HandleWorkbookLoad(Wb);
        }

        private void Application_NewWorkbook(Excel.Workbook Wb)
        {
            HandleWorkbookLoad(Wb);
        }

        private void HandleWorkbookLoad(Excel.Workbook Wb)
        {
            if (Ribbon.IsAutoUpdateOnLoadDoc)
                ribbon.UpdateDatesFromCustomXmlPartsForExcel(Wb);
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
