using System;
using System.Reflection;
using System.Windows.Forms;
using KDCLibrary;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace KDC_Outlook
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

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ribbon.OutlookApp = Globals.ThisAddIn.Application;

            this.Application.ItemLoad += new ApplicationEvents_11_ItemLoadEventHandler(
                Application_ItemLoad
            );
        }

        private void Application_ItemLoad(object Item)
        {
            if (Item is ItemEvents_10_Event itemEvent)
            {
                itemEvent.Open += new ItemEvents_10_OpenEventHandler(
                    (ref bool Cancel) => Item_Open(Item, ref Cancel)
                );
            }
        }

        private void Item_Open(object Item, ref bool Cancel)
        {
            switch (Item)
            {
                case MailItem mailItem:
                    var document = mailItem.GetInspector.WordEditor as Word.Document;
                    if (document != null && Ribbon.IsAutoUpdateOnLoadDoc)
                    {
                        ribbon.UpdateDatesFromCustomXmlPartsForWordOutlook(document);
                    }
                    break;
                case AppointmentItem appointmentItem:
                    appointmentItem.Body = appointmentItem.Body;
                    appointmentItem.Save();
                    break;
                case TaskItem taskItem:
                    taskItem.Body = taskItem.Body;
                    taskItem.Save();
                    break;
                case ContactItem contactItem:
                    contactItem.Body = contactItem.Body;
                    contactItem.Save();
                    break;
                default:
                    MessageBox.Show(
                        "This item type does not support updates.",
                        "Info",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    break;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Cleanup if needed
        }

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
