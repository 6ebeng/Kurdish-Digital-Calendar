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

        void SaveSetting(string keyName, string value, string appName);

        string LoadSetting(string keyName, string defaultValue, string appName);

        int DetermineFormatChoiceFromCheckbox(string checkboxLabel);
    }
}
