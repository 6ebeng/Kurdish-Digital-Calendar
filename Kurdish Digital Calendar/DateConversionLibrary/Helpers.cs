using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kurdish_Digital_Calendar.DateConversionLibrary
{
    internal class Helpers
    {

        public static string GetCultureInfoFromLanguage(string language)
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

        public static string GetHijriUmmalquraSuffix(string language)
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

        public static string GetGregorianSuffix(string language)
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

        public static string GetKurdishSuffix(string language)
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

        public static int SelectFormatChoice(string format)
        {
            switch (format)
            {
                case "dddd, dd MMMM, yyyy":
                    return 1;
                case "dddd, dd/MM/yyyy":
                    return 2;
                case "dd MMMM, yyyy":
                    return 3;
                case "dd/MM/yyyy":
                    return 4;
                case "MM/dd/yyyy":
                    return 10;
                case "yyyy/MM/dd":
                    return 11;
                default:
                    MessageBox.Show("Unsupported format selected.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return -1;
            }
        }

    }
}
