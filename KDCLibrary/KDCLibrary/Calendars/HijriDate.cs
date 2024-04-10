using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KDCLibrary.Calendars
{
    internal class HijriDate
    {
        // Convert Gregorian date to Hijri date
        public string FromGregorianToHijri(
            DateTime gDate,
            int formatChoice,
            string language,
            bool isAddSuffix
        )
        {
            HijriCalendar hijriCalendar = new HijriCalendar();
            int year = hijriCalendar.GetYear(gDate);
            int month = hijriCalendar.GetMonth(gDate);
            int day = hijriCalendar.GetDayOfMonth(gDate);

            string suffix = isAddSuffix ? new Helper().GetHijriUmmalquraSuffix(language) : "";

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
                    nameMonth = month.ToString("MMMM", cultureInfo);
                    dayName = day.ToString("dddd", cultureInfo);
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

        public DateTime FromHijriToGregorian(DateTime date)
        {
            HijriCalendar hijriCalendar = new HijriCalendar();
            hijriCalendar.HijriAdjustment = 0; // Adjusting the Hijri calendar to match the Umm Al-Qura calendar

            return hijriCalendar.ToDateTime(date.Year, date.Month, date.Day, 0, 0, 0, 0);
        }

        public string ArabicWeekdayName(int index)
        {
            string[] weekdays = new string[]
            {
                "الأحد", // Sunday
                "الاثنين", // Monday
                "الثلاثاء", // Tuesday
                "الأربعاء", // Wednesday
                "الخميس", // Thursday
                "الجمعة", // Friday
                "السبت" // Saturday
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
