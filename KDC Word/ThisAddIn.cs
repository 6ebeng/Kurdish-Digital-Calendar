using System.Reflection;
using KDCLibrary;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace KDC_Word
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
            ribbon.WordApp = Globals.ThisAddIn.Application;

            if (Globals.ThisAddIn.Application.Documents.Count > 0)
            {
                HandleDocumentLoad(Globals.ThisAddIn.Application.ActiveDocument);
            }
            // currrent document open event
            this.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(
                Application_DocumentOpen
            );
            ((ApplicationEvents4_Event)this.Application).NewDocument +=
                new ApplicationEvents4_NewDocumentEventHandler(Application_NewDocument);
        }

        private void Application_DocumentOpen(Document Doc)
        {
            HandleDocumentLoad(Doc);
        }

        private void Application_NewDocument(Document Doc)
        {
            HandleDocumentLoad(Doc);
        }

        private void Application_Startup(Document Doc)
        {
            HandleDocumentLoad(Doc);
        }

        private void HandleDocumentLoad(Document Doc)
        {
            if (Ribbon.IsAutoUpdateOnLoadDoc)
                ribbon.UpdateDatesFromCustomXmlPartsForWordOutlook(Doc);
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
