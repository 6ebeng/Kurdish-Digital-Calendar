using Microsoft.Win32;

namespace KDCLibrary.Helpers
{
    internal class RegistryHelper
    {
        private const string RegistryPath = @"SOFTWARE\6ebeng\KurdishDigitalCalendar";

        public void SaveSetting(string keyName, string value)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath))
            {
                key?.SetValue(keyName, value);
            }
        }

        public string LoadSetting(string keyName, string defaultValue)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
            {
                if (key != null)
                {
                    return key.GetValue(keyName, defaultValue).ToString();
                }
                return defaultValue;
            }
        }
    }
}
