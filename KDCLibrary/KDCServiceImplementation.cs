using KDCLibrary.Calendars;
using KDCLibrary.Helpers;

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

        public void SaveSetting(string keyName, string value, string appName)
        {
            new RegistryHelper().SaveSetting(keyName, value, appName);
        }

        public string LoadSetting(string keyName, string defaultValue, string appName)
        {
            return new RegistryHelper().LoadSetting(keyName, defaultValue, appName);
        }

        public int DetermineFormatChoiceFromCheckbox(string checkboxLabel)
        {
            return new Helper().DetermineFormatChoiceFromCheckbox(checkboxLabel);
        }
    }
}
