using Microsoft.Win32;

namespace KDCLibrary.Helpers
{
    internal class RegistryHelper
    {
        private const string RegistryPath = @"SOFTWARE\Rekbin Devs\Kurdish Digital Calendar";

        public void SaveSetting(string keyName, string value, string appName)
        {
            using (
                RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath + "\\" + appName)
            )
            {
                key?.SetValue(keyName, value);
            }
        }

        public string LoadSetting(string keyName, string defaultValue, string appName)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath + "\\" + appName))
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
