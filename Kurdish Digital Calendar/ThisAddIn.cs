using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Kurdish_Digital_Calendar.DateConversionLibrary;
using Microsoft.Office.Tools.Ribbon;

namespace Kurdish_Digital_Calendar
{
    public partial class ThisAddIn
    {

        public void InsertKurdishDate(int formatChoice, string dialect, bool isAddSuffix)
        {
            DateTime todayGregorian = DateTime.Today; // Today's Gregorian date
            string todayKurdish = KurdishDate.fromGregorianToKurdish(todayGregorian, formatChoice, dialect, isAddSuffix);

            // Ensure there's a selection or a place to insert the text
            if (this.Application.Selection != null)
            {
                this.Application.Selection.TypeText(todayKurdish);
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
