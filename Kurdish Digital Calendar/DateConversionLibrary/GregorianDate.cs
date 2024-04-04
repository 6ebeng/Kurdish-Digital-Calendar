using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kurdish_Digital_Calendar.DateConversionLibrary
{
    internal class GregorianDate
    {

        public static string FormatGregorian(DateTime gDate, int formatChoice, string language, bool isAddSuffix)
        {
            string suffix = isAddSuffix ? Helpers.GetGregorianSuffix(language) : "";

            string saperator = language == "Arabic" ? "، " : ", ";

            var cultureInfo = new CultureInfo(Helpers.GetCultureInfoFromLanguage(language));
            string nameMonth;
            string dayName;

            switch (language)
            {
                case "Arabic":
                    nameMonth = GregorianDate.GregorianMonthNameArabic(gDate.Month);
                    dayName = gDate.ToString("dddd", cultureInfo);
                    break;
                case "English":
                    nameMonth = gDate.ToString("MMMM", cultureInfo);
                    dayName = gDate.ToString("dddd", cultureInfo);
                    break;

                case "Kurdish (Central)":
                    nameMonth = GregorianDate.GregorianMonthNameKurdishCentral(gDate.Month);
                    dayName = KurdishDate.KurdishWeekdayNameCentral((int)gDate.DayOfWeek + 1);
                    break;
                case "Kurdish (Northern)":
                    nameMonth = GregorianDate.GregorianMonthNameKurdishNorthern(gDate.Month);
                    dayName = KurdishDate.KurdishWeekdayNameNorthern((int)gDate.DayOfWeek + 1);
                    break;
                default:
                    nameMonth = gDate.ToString("MMMM", cultureInfo);
                    dayName = gDate.ToString("dddd", cultureInfo);
                    break;
            }
            return DateFormatting.FormatDate(gDate.Day, gDate.Month, gDate.Year, (int)gDate.DayOfWeek, formatChoice, nameMonth, dayName);
        }

        public static string GregorianMonthNameArabic(int index)
        {
            string[] months = new string[]
            {
                "يناير", // January
                "فبراير", // February
                "مارس", // March
                "أبريل", // April
                "مايو", // May
                "يونيو", // June
                "يوليو", // July
                "أغسطس", // August
                "سبتمبر", // September
                "أكتوبر", // October
                "نوفمبر", // November
                "ديسمبر" // December
            };
            return months[index - 1];
        }

        public static string GregorianMonthNameKurdishCentral(int index)
        {
            string[] months = new string[]
            {
                "کانونی دووەم",
                "شوبات",
                "ئازار",
                "نیسان",
                "ئایار",
                "حوزەیران",
                "تەمموز",
                "ئاب",
                "ئەیلوول",
                "تشرینی یەكەم",
                "تشرینی دووەم",
                "كانونی یەکەم"
            };

            return months[index - 1];
        }

        public static string GregorianMonthNameKurdishNorthern(int index)
        {
            string[] months = new string[]
            {
                "Çile",
                "Şibat",
                "Adar",
                "Nîsan",
                "Gulan",
                "Pûşper",
                "Tîrmeh",
                "Tebax",
                "Îlon",
                "Cotmeh",
                "Mijdar",
                "Kanûn"
            };

            return months[index - 1];
        }


    }
}
