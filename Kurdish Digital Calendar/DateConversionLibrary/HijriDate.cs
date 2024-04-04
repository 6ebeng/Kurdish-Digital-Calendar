using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.AxHost;

namespace Kurdish_Digital_Calendar.DateConversionLibrary
{
    internal class HijriDate
    {
        public DateTime Date { get; private set; }

        public HijriDate(int year, int month, int day)
        {
            HijriCalendar hijriCalendar = new HijriCalendar();
            this.Date = hijriCalendar.ToDateTime(year, month, day, 0, 0, 0, 0);
        }

        // Convert Hijri date to Gregorian date
        // Since the DateTime object is inherently in the Gregorian calendar, this method correctly returns the Date property.
        public DateTime ToGregorian()
        {
            return Date;
        }

        // Convert Gregorian date to Hijri date
        public static string FromGregorianToHijri(DateTime gDate,int formatChoice, string language, bool isAddSuffix)
        {
            HijriCalendar hijriCalendar = new HijriCalendar();
            int year = hijriCalendar.GetYear(gDate);
            int month = hijriCalendar.GetMonth(gDate);
            int day = hijriCalendar.GetDayOfMonth(gDate);

            string suffix = isAddSuffix ? Helpers.GetHijriUmmalquraSuffix(language) : "";

            string saperator = language == "Arabic" ? "، " : ", ";

            var cultureInfo = new CultureInfo(Helpers.GetCultureInfoFromLanguage(language));
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
                    dayName = KurdishDate.KurdishWeekdayNameCentral((int)gDate.DayOfWeek + 1);
                    break;
                case "Kurdish (Northern)":
                    nameMonth = HijriUmmalquraMonthNameKurdishNorthern(month);
                    dayName = KurdishDate.KurdishWeekdayNameNorthern((int)gDate.DayOfWeek + 1);
                    break;
                default:
                    nameMonth = month.ToString("MMMM", cultureInfo);
                    dayName = day.ToString("dddd", cultureInfo);
                    break;
            }

            return DateFormatting.FormatDate(day, month, year, (int)gDate.DayOfWeek + 1, formatChoice, nameMonth, dayName, saperator, suffix);
        }

        public static DateTime FromHijriToGregorian(DateTime date)
        {
            HijriCalendar hijriCalendar = new HijriCalendar();
            hijriCalendar.HijriAdjustment = 0; // Adjusting the Hijri calendar to match the Umm Al-Qura calendar

            return hijriCalendar.ToDateTime(date.Year, date.Month, date.Day, 0, 0, 0, 0); 
        }



        public static string ArabicWeekdayName(int index)
        {
            string[] weekdays = new string[]
            {
                "الأحد", // Sunday
                "الاثنين", // Monday
                "الثلاثاء", // Tuesday
                "الأربعاء", // Wednesday
                "الخميس", // Thursday
                "الجمعة", // Friday
                "السبت"  // Saturday
            };
            return weekdays[index - 1];
        }

        public static string HijriUmmalquraMonthNameArabic(int index)
        {
            string[] months = new string[]{
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

        public static string HijriUmmalquraMonthNameEnglish(int index)
        {
            string[] months = new string[]{
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

        public static string HijriUmmalquraMonthNameKurdishCentral(int index)
        {
            string[] months = new string[]{
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



        public static string HijriUmmalquraMonthNameKurdishNorthern(int index)
        {
            string[] months = new string[]{
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
