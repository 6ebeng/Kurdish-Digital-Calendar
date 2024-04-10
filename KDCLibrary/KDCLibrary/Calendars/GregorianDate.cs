using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KDCLibrary.Calendars
{
    internal class GregorianDate
    {
        public string FormatGregorian(
            DateTime gDate,
            int formatChoice,
            string language,
            bool isAddSuffix
        )
        {
            string suffix = isAddSuffix ? new Helper().GetGregorianSuffix(language) : "";

            string saperator = language == "Arabic" ? "، " : ", ";

            var cultureInfo = new CultureInfo(new Helper().GetCultureInfoFromLanguage(language));
            string nameMonth;
            string dayName;

            switch (language)
            {
                case "Arabic":
                    nameMonth = new GregorianDate().GregorianMonthNameArabic(gDate.Month);
                    dayName = gDate.ToString("dddd", cultureInfo);
                    break;
                case "English":
                    nameMonth = gDate.ToString("MMMM", cultureInfo);
                    dayName = gDate.ToString("dddd", cultureInfo);
                    break;
                case "Kurdish (Central)":
                    nameMonth = new GregorianDate().GregorianMonthNameKurdishCentral(gDate.Month);
                    dayName = new KurdishDate().KurdishWeekdayNameCentral((int)gDate.DayOfWeek + 1);
                    break;
                case "Kurdish (Northern)":
                    nameMonth = new GregorianDate().GregorianMonthNameKurdishNorthern(gDate.Month);
                    dayName = new KurdishDate().KurdishWeekdayNameNorthern(
                        (int)gDate.DayOfWeek + 1
                    );
                    break;
                default:
                    nameMonth = gDate.ToString("MMMM", cultureInfo);
                    dayName = gDate.ToString("dddd", cultureInfo);
                    break;
            }
            return new DateFormatting().FormatDate(
                gDate.Day,
                gDate.Month,
                gDate.Year,
                (int)gDate.DayOfWeek,
                formatChoice,
                nameMonth,
                dayName,
                saperator,
                suffix
            );
        }

        public string GregorianMonthNameArabic(int index)
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

        public string GregorianMonthNameKurdishCentral(int index)
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

        public string GregorianMonthNameKurdishNorthern(int index)
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
