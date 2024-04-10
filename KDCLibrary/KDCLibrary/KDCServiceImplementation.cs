using System.Collections.Generic;
using KDCLibrary.Calendars;
using KDCLibrary.Helpers;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Project = Microsoft.Office.Interop.MSProject;
using Visio = Microsoft.Office.Interop.Visio;
using Word = Microsoft.Office.Interop.Word;

namespace KDCLibrary
{
    public class KDCServiceImplementation : IKDCService
    {
        public string Kurdish(int formatChoice, string dialect, bool isAddSuffix)
        {
            return new DateInsertion().Kurdish(formatChoice, dialect, isAddSuffix);
        }

        public string ConvertDateBasedOnUserSelection(
            string selectedText,
            bool isReverse,
            string targetDialect,
            string fromFormat,
            string toFormat,
            string targetCalendar,
            bool isAddSuffix
        )
        {
            return new DateConversion().ConvertDateBasedOnUserSelection(
                selectedText,
                isReverse,
                targetDialect,
                fromFormat,
                toFormat,
                targetCalendar,
                isAddSuffix
            );
        }

        public void Credits()
        {
            new CreditsForm().Show();
        }

        public string GetRibbonXml()
        {
            return new Helper().GetResourceText("KDCLibrary.UI.Ribbon.xml");
        }

        public void SaveSetting(string keyName, string value)
        {
            new RegistryHelper().SaveSetting(keyName, value);
        }

        public string LoadSetting(string keyName, string defaultValue)
        {
            return new RegistryHelper().LoadSetting(keyName, defaultValue);
        }

        public int DetermineFormatChoiceFromCheckbox(string checkboxLabel)
        {
            return new Helper().DetermineFormatChoiceFromCheckbox(checkboxLabel);
        }
    }
}
