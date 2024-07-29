using System;
using System.Runtime.InteropServices;
using KDCLibrary.Helpers;

namespace KDCLibrary.Calendars
{
    [ComVisible(false)]
    public class KurdishDate
    {
        public static readonly int ADEra = 1;

        // Convert Gregorian date to Kurdish date
        public string FromGregorianToKurdish(
            DateTime gDate,
            int formatChoice,
            string dialect,
            bool isAddSuffix
        )
        {
            var gYear = gDate.Year;
            int[] daysInMonth =
            {
                31,
                31,
                31,
                31,
                31,
                31,
                30,
                30,
                30,
                30,
                30,
                DateTime.IsLeapYear(gYear) ? 30 : 29 //The kurdish leap year occur in last months which is almost new gregorian year (I followed the logic every greorian leap year is the previous year of kurdish leap year + 700)
            };

            DateTime referenceDate = new DateTime(gYear, 3, 21);
            int daysOffset;

            if (gDate >= referenceDate)
            {
                daysOffset = (gDate - referenceDate).Days;
                gYear++; // Adjusting for the current year in the Kurdish calendar
            }
            else
            {
                referenceDate = new DateTime(gYear - 1, 3, 21);
                daysOffset = (gDate - referenceDate).Days;
            }

            int kYear = gYear + 699; // Adjusting for the Kurdish year
            int kMonth = 1;
            while (daysOffset >= daysInMonth[kMonth - 1])
            {
                daysOffset -= daysInMonth[kMonth - 1];
                kMonth++;
            }
            int kDay = daysOffset + 1;

            int kWeekDay = (int)gDate.DayOfWeek + 1; // Adjusting to match Kurdish week starting with Sunday as 1

            string suffix = isAddSuffix ? new Helper().GetKurdishSuffix(dialect) : "";
            string saperator = dialect == "Kurdish (Central)" ? "، " : ", ";

            // Utilize the KurdishDate class methods based on the dialect
            string monthName =
                (dialect == "Kurdish (Central)")
                    ? new KurdishDate().KurdishMonthNameCentral(kMonth)
                    : new KurdishDate().KurdishMonthNameNorthern(kMonth);
            string weekDayName =
                (dialect == "Kurdish (Central)")
                    ? new KurdishDate().KurdishWeekdayNameCentral(kWeekDay)
                    : new KurdishDate().KurdishWeekdayNameNorthern(kWeekDay);

            // This assumes a FormatKurdishDateDialect method exists to handle the formatting.
            // You'll need to adjust this call to match your actual implementation, which might involve creating a new method or adjusting existing logic.
            return new DateFormatting().FormatDate(
                kDay,
                kMonth,
                kYear,
                kWeekDay,
                formatChoice,
                monthName,
                weekDayName,
                saperator,
                suffix
            );
        }

        // Convert Kurdish date to Gregorian date
        public DateTime FromKurdishToGregorian(int kDay, int kMonth, int kYear)
        {
            int gYear = kYear - 700; // Adjusting the Kurdish year to the Gregorian year
            int[] daysInMonth =
            {
                31,
                31,
                31,
                31,
                31,
                31,
                30,
                30,
                30,
                30,
                30,
                DateTime.IsLeapYear(gYear) ? 30 : 29
            };

            DateTime referenceDate = new DateTime(gYear, 3, 21);
            int daysBeforeNewYear = kDay - 1; // Days in the current month

            for (int i = 0; i < kMonth - 1; i++)
            {
                daysBeforeNewYear += daysInMonth[i];
            }

            // Calculate the Gregorian date by adding the days before the New Year to the reference date
            return referenceDate.AddDays(daysBeforeNewYear);
        }

        public string KurdishWeekdayNameCentral(int index)
        {
            return KurdishWeekdayNameCentralArray()[index - 1];
        }

        public string KurdishMonthNameCentral(int index)
        {
            return KurdishMonthNameCentralArray()[index - 1];
        }

        public string KurdishWeekdayNameNorthern(int index)
        {
            return KurdishWeekdayNameNorthernArray()[index - 1];
        }

        public string KurdishMonthNameNorthern(int index)
        {
            return KurdishMonthNameNorthernArray()[index - 1];
        }

        public string[] KurdishMonthNameCentralArray()
        {
            return new string[]
            {
                "نەورۆز",
                "گوڵان",
                "جۆزەردان",
                "پووشپەڕ",
                "خەرمانان",
                "گەلاوێژ",
                "رەزبەر",
                "گەڵاڕێزان",
                "سەرماوەز",
                "بەفرانبار",
                "ڕێبەندان",
                "ڕه‌شه‌مێ",
                ""
            };
        }

        public string[] KurdishWeekdayNameCentralArray()
        {
            return new string[]
            {
                "یەکشەممە",
                "دووشەممە",
                "سێشەممە",
                "چوارشەممە",
                "پێنجشەممە",
                "هەینی",
                "شەممە"
            };
        }

        public string[] KurdishMonthNameNorthernArray()
        {
            return new string[]
            {
                "Nêwroz",
                "Gullan",
                "Avrêl",
                "Pusper",
                "Tîrmeh",
                "Gelawêj",
                "Rezber",
                "Kewçêr",
                "Sermawez",
                "Berfanbar",
                "Rêbendan",
                "Resheme",
                ""
            };
        }

        public string[] KurdishWeekdayNameNorthernArray()
        {
            return new string[] { "Yekşem", "Duşem", "Sêşem", "Çarşem", "Pêncşem", "Înê", "Şemî" };
        }

        public int GetDaysInKurdishMonth(int year, int month)
        {
            switch (month)
            {
                case 1:
                case 2:
                case 3:
                case 4:
                case 5:
                case 6:
                    return 31;
                case 7:
                case 8:
                case 9:
                case 10:
                case 11:
                    return 30;
                case 12:
                    return DateTime.IsLeapYear(year) ? 30 : 29; //The kurdish leap year occur in last months which is almost new gregorian year (I followed the logic every greorian leap year is the previous year of kurdish leap year + 700)
                default:
                    return 0;
            }
        }

        public DateTime GetFirstDayOfYear(int year)
        {
            return new DateTime(year, 3, 21);
        }
    }
}
