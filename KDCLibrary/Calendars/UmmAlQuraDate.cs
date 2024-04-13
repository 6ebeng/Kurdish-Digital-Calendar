using System;
using System.Globalization;

namespace KDCLibrary.Calendars
{
    internal class UmmAlQuraDate
    {
        // Convert Gregorian date to Umm Al-Qura date
        public string FromGregorianToUmmAlQura(
            DateTime gDate,
            int formatChoice,
            string language,
            bool isAddSuffix
        )
        {
            UmAlQuraCalendar ummAlQuraCalendar = new UmAlQuraCalendar();
            int year = ummAlQuraCalendar.GetYear(gDate);
            int month = ummAlQuraCalendar.GetMonth(gDate);
            int day = ummAlQuraCalendar.GetDayOfMonth(gDate);

            string suffix = isAddSuffix ? GetArabicSuffix(language) : "";

            string saperator = language == "Arabic" ? "، " : ", ";

            var cultureInfo = new CultureInfo(new Helper().GetCultureInfoFromLanguage(language));
            string nameMonth;
            string dayName;

            switch (language)
            {
                case "Arabic":
                    nameMonth = HijriUmmalquraMonthNameArabic(month);
                    dayName = ArabicWeekdayName((int)gDate.DayOfWeek + 1);
                    break;
                case "English":
                    nameMonth = HijriUmmalquraMonthNameEnglish(month);
                    dayName = gDate.ToString("dddd", cultureInfo);
                    break;

                case "Kurdish (Central)":
                    nameMonth = HijriUmmalquraMonthNameKurdishCentral(month);
                    dayName = new KurdishDate().KurdishWeekdayNameCentral((int)gDate.DayOfWeek + 1);
                    break;
                case "Kurdish (Northern)":
                    nameMonth = HijriUmmalquraMonthNameKurdishNorthern(month);
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
                day,
                month,
                year,
                (int)gDate.DayOfWeek + 1,
                formatChoice,
                nameMonth,
                dayName,
                saperator,
                suffix
            );
        }

        public DateTime FromUmmAlQuraToGregorian(DateTime date)
        {
            HijriCalendar ummAlQuraCalendar = new HijriCalendar();
            ummAlQuraCalendar.HijriAdjustment = -1; // Adjusting the Hijri calendar to match the Umm Al-Qura calendar

            return ummAlQuraCalendar.ToDateTime(date.Year, date.Month, date.Day, 0, 0, 0, 0);
        }

        public string GetArabicSuffix(string language)
        {
            switch (language)
            {
                case "Arabic":
                    return "هـ";
                case "Kurdish (Central)":
                    return "ی كۆچی";
                case "Kurdish (Northern)":
                    return "Koçî";
                case "English":
                    return "AH";
                default:
                    return "";
            }
        }

        public string HijriUmmalquraMonthNameEnglish(int index)
        {
            string[] months = new string[]
            {
                "Muharram",
                "Safar",
                "Rabi' al-awwal",
                "Rabi' al-thani",
                "Jumada al-awwal",
                "Jumada al-thani",
                "Rajab",
                "Sha'ban",
                "Ramadan",
                "Shawwal",
                "Dhu al-Qi'dah",
                "Dhu al-Hijjah"
            };
            return months[index - 1]; // Adjusting for zero-based index
        }

        public string ArabicWeekdayName(int index)
        {
            string[] weekdays = new string[]
            {
                "الأحد",
                "الاثنين",
                "الثلاثاء",
                "الأربعاء",
                "الخميس",
                "الجمعة",
                "السبت"
            };
            return weekdays[index - 1];
        }

        public string HijriUmmalquraMonthNameArabic(int index)
        {
            string[] months = new string[]
            {
                "محرم",
                "صفر",
                "ربيع الأول",
                "ربيع الثاني",
                "جمادى الأولى",
                "جمادى الآخرة",
                "رجب",
                "شعبان",
                "رمضان",
                "شوال",
                "ذو القعدة",
                "ذو الحجة"
            };
            return months[index - 1]; // Adjusting for zero-based index
        }

        public string HijriUmmalquraMonthNameKurdishCentral(int index)
        {
            string[] months = new string[]
            {
                "موحەڕەم",
                "سەفەر",
                "ڕەبیعی یه‌كه‌م ",
                "ڕەبیعی دووه‌م",
                "جه‌مادی یه‌كه‌م",
                "جه‌مادی دووه‌م",
                "ڕەجەب",
                "شەعبان",
                "ڕەمەزان",
                "شەوال",
                "زولقەعدە",
                "زولحیججە"
            };
            return months[index - 1]; // Adjusting for zero-based index
        }

        public string HijriUmmalquraMonthNameKurdishNorthern(int index)
        {
            string[] months = new string[]
            {
                "Muherem",
                "Sefer",
                "Rebî'ulewel",
                "Rebî'usanî",
                "Cumadalûla",
                "Cumadasaniye",
                "Receb",
                "Şeban",
                "Remezan",
                "Şewel",
                "Zîlqe'de",
                "Zîlhice"
            };
            return months[index - 1]; // Adjusting for zero-based index
        }
    }
}
