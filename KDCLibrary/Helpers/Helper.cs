using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace KDCLibrary.Helpers
{
    internal class Helper
    {
        public string GetCultureInfoFromLanguage(string language)
        {
            switch (language.ToLower())
            {
                case "arabic":
                    return "ar-US"; // CultureInfo for Saudi Arabia (Arabic)
                case "english":
                    return "en-US"; // CultureInfo for United States (English)
                case "kurdish (Central)":
                    return "ku-IQ";
                case "kurdish (northern)":
                    return "ku-TR";
                default:
                    return "en-US"; // Default to English if unknown language
            }
        }

        public string GetHijriUmmalquraSuffix(string language)
        {
            switch (language)
            {
                case "Arabic":
                    return "هـ ";
                case "Kurdish (Central)":
                    return "ی كۆچی";
                case "Kurdish (Northern)":
                    return " Koçî";
                case "English":
                    return " AH";
                default:
                    return "";
            }
        }

        public string GetGregorianSuffix(string language)
        {
            switch (language)
            {
                case "Arabic":
                    return "م ";
                case "Kurdish (Central)":
                    return "ی زایینی";
                case "Kurdish (Northern)":
                    return " Zayînî";
                case "English":
                    return " AD";
                default:
                    return "";
            }
        }

        public string GetKurdishSuffix(string language)
        {
            switch (language)
            {
                case "Arabic":
                    return "كردی ";
                case "Kurdish (Central)":
                    return "ی كوردی";
                case "Kurdish (Northern)":
                    return " Kurdî";
                case "English":
                    return " Kurdish";
                default:
                    return "";
            }
        }

        public int SelectFormatChoice(string format)
        {
            switch (format)
            {
                case "dddd, dd MMMM, yyyy":
                    return 1;
                case "dddd, dd/MM/yyyy":
                    return 2;
                case "dd MMMM, yyyy":
                    return 3;
                case "MMMM dd, yyyy":
                    return 15;
                case "dd/MM/yyyy":
                    return 4;
                case "MM/dd/yyyy":
                    return 10;
                case "yyyy/MM/dd":
                    return 11;
                case "MMMM yyyy":
                    return 16;
                case "MM/yyyy":
                    return 5;
                case "MMMM":
                    return 8;
                case "yyyy":
                    return 6;
                default:
                    MessageBox.Show(
                        "Unsupported format selected.",
                        "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation
                    );
                    return -1;
            }
        }

        public string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (
                    string.Compare(
                        resourceName,
                        resourceNames[i],
                        StringComparison.OrdinalIgnoreCase
                    ) == 0
                )
                {
                    using (
                        StreamReader resourceReader = new StreamReader(
                            asm.GetManifestResourceStream(resourceNames[i])
                        )
                    )
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }
}
