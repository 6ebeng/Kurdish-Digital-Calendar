using System;
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

        private string GetRegistryValue(string keyName, string valueName, RegistryView registryView)
        {
            try
            {
                using (
                    RegistryKey baseKey = RegistryKey.OpenBaseKey(
                        RegistryHive.LocalMachine,
                        registryView
                    )
                )
                using (RegistryKey key = baseKey.OpenSubKey(keyName))
                {
                    return key?.GetValue(valueName)?.ToString();
                }
            }
            catch (Exception)
            {
                try
                {
                    using (
                        RegistryKey baseKey = RegistryKey.OpenBaseKey(
                            RegistryHive.CurrentUser,
                            registryView
                        )
                    )
                    using (RegistryKey key = baseKey.OpenSubKey(keyName))
                    {
                        return key?.GetValue(valueName)?.ToString();
                    }
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        public string GetRegistryValueFromX6432Path(string keyName, string valueName)
        {
            if (string.IsNullOrEmpty(keyName) || string.IsNullOrEmpty(valueName))
            {
                return null;
            }

            string value =
                GetRegistryValue(keyName, valueName, RegistryView.Registry64)
                ?? GetRegistryValue(
                    keyName.Replace("SOFTWARE", "SOFTWARE\\Wow6432Node"),
                    valueName,
                    RegistryView.Registry64
                );

            return value;
        }
    }
}
