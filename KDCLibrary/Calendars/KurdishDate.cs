using System;

namespace KDCLibrary.Calendars
{
    internal class KurdishDate
    {
        public string FromGregorianToKurdish(
            DateTime gDate,
            int formatChoice,
            string dialect,
            bool isAddSuffix
        )
        {
            var gYear = gDate.Year;
            bool isGregorianLeapYear = DateTime.IsLeapYear(gYear);
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
                isGregorianLeapYear ? 30 : 29
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

        // Implementation of ConvertKurdishToGregorian method
        public DateTime FromKurdishToGregorian(int kDay, int kMonth, int kYear)
        {
            int gYear = kYear - 700; // Adjusting the Kurdish year to the Gregorian year
            bool isGregorianLeapYear = DateTime.IsLeapYear(gYear);
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
                isGregorianLeapYear ? 30 : 29
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
            string[] weekdays = new string[]
            {
                "یەکشەممە", // Sunday
                "دووشەممە", // Monday
                "سێشەممە", // Tuesday
                "چوارشەممە", // Wednesday
                "پێنجشەممە", // Thursday
                "هەینی", // Friday
                "شەممە" // Saturday
            };
            return weekdays[index - 1];
        }

        public string KurdishMonthNameCentral(int index)
        {
            string[] months = new string[]
            {
                "نەورۆز", // March
                "گوڵان", // April
                "جۆزەردان", // May
                "پووشپەر", // June
                "گەلاوێژ", // July
                "خەرمانان", // August
                "رەزبەر", // September
                "گەلاڕێزان", // October
                "سەرماوەز", // November
                "بەفرانبار", // December
                "رێبەندان", // January
                "رەشەمە" // February
            };
            return months[index - 1];
        }

        public string KurdishWeekdayNameNorthern(int index)
        {
            string[] weekdays = new string[]
            {
                "Yekşem", // Sunday
                "Duşem", // Monday
                "Sêşem", // Tuesday
                "Çarşem", // Wednesday
                "Pêncşem", // Thursday
                "Înê", // Friday
                "Şemî" // Saturday
            };
            return weekdays[index - 1];
        }

        public string KurdishMonthNameNorthern(int index)
        {
            string[] months = new string[]
            {
                "Nêwroz", // March
                "Gullan", // April
                "Avrêl", // May (Note: "Avrêl" is not traditionally a Kurdish name but often used for May in absence of a widely accepted Kurdish equivalent, reflecting April in Gregorian. Adjust as per actual usage or replace with the traditional Kurdish name for May if applicable.)
                "Pusper", // June
                "Tîrmeh", // July
                "Gelawêj", // August
                "Rezber", // September
                "Kewçêr", // October
                "Sermawez", // November
                "Berfanbar", // December
                "Rêbendan", // January
                "Resheme" // February
            };
            return months[index - 1];
        }
    }
}
