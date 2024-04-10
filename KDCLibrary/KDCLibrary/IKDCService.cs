using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Project = Microsoft.Office.Interop.MSProject;
using Visio = Microsoft.Office.Interop.Visio;
using Word = Microsoft.Office.Interop.Word;

namespace KDCLibrary
{
    public interface IKDCService
    {
        // Business Logic methods
        string Kurdish(int formatChoice, string dialect, bool isAddSuffix);
        string ConvertDateBasedOnUserSelection(
            string selectedText,
            bool isReverse,
            string targetDialect,
            string fromFormat,
            string toFormat,
            string targetCalendar,
            bool isAddSuffix
        );

        void Credits();

        string GetRibbonXml();

        void SaveSetting(string keyName, string value);

        string LoadSetting(string keyName, string defaultValue);

        int DetermineFormatChoiceFromCheckbox(string checkboxLabel);
    }
}
