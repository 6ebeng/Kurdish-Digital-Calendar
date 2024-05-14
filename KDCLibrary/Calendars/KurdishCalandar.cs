using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace KDCLibrary.Calendars
{
    [ComVisible(false)]
    public class KurdishCalendar : Calendar
    {
        public const int ADEra = 1;

        public const int DatePartYear = 0;

        public const int DatePartDayOfYear = 1;

        public const int DatePartMonth = 2;

        public const int DatePartDay = 3;

        public const int MaxYear = 9999;

        public const int AdjustedKurdishYear = 700;

        public static int[] DaysToMonth365 = new int[13]
        {
            0,
            31,
            62,
            93,
            124,
            155,
            186,
            216,
            246,
            276,
            306,
            336,
            365
        };

        public static int[] DaysToMonth366 = new int[13]
        {
            0,
            31,
            62,
            93,
            124,
            155,
            186,
            216,
            246,
            276,
            306,
            336,
            366
        };

        public override CalendarAlgorithmType AlgorithmType => CalendarAlgorithmType.SolarCalendar;

        public override int[] Eras => new int[1] { 1 };

        public override DateTime MinSupportedDateTime => new DateTime(700, 3, 21);
        public override DateTime MaxSupportedDateTime => new DateTime(9999, 3, 20);

        public override DateTime AddMonths(DateTime time, int months)
        {
            if (months < -120000 || months > 120000)
            {
                throw new ArgumentOutOfRangeException(
                    "months",
                    string.Format(
                        CultureInfo.CurrentCulture,
                        "ArgumentOutOfRange_Range -120000, 120000"
                    )
                );
            }

            //time.GetDatePart(out var year, out var month, out var day);

            var year = time.Year;
            var month = time.Month;
            var day = time.Day;

            int num = month - 1 + months;
            if (num >= 0)
            {
                month = num % 12 + 1;
                year += num / 12;
            }
            else
            {
                month = 12 + (num + 1) % 12;
                year += (num - 11) / 12;
            }

            int[] array = (
                (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0))
                    ? DaysToMonth366
                    : DaysToMonth365
            );
            int num2 = array[month] - array[month - 1];
            if (day > num2)
            {
                day = num2;
            }

            long ticks = DateToTicks(year, month, day) + time.Ticks % 864000000000L;
            //Calendar.CheckAddResult(ticks, MinSupportedDateTime, MaxSupportedDateTime);
            return new DateTime(ticks);
        }

        public override DateTime AddYears(DateTime time, int years)
        {
            return AddMonths(time, years * 12);
        }

        public override int GetDayOfMonth(DateTime time)
        {
            return time.Day;
        }

        public override DayOfWeek GetDayOfWeek(DateTime time)
        {
            return (DayOfWeek)((int)(time.Ticks / 864000000000L + 1) % 7);
        }

        public override int GetDayOfYear(DateTime time)
        {
            return time.DayOfYear;
        }

        public override int GetDaysInMonth(int year, int month, int era)
        {
            if (era == 0 || era == 1)
            {
                if (year < 1 || year > 9999)
                {
                    throw new ArgumentOutOfRangeException(
                        "year",
                        "ArgumentOutOfRange_Range 1, 9999"
                    );
                }

                if (month < 1 || month > 12)
                {
                    throw new ArgumentOutOfRangeException("month", "ArgumentOutOfRange_Month");
                }

                int[] array = (
                    (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0))
                        ? DaysToMonth366
                        : DaysToMonth365
                );
                return array[month] - array[month - 1];
            }

            throw new ArgumentOutOfRangeException("era", "ArgumentOutOfRange_InvalidEraValue");
        }

        public override int GetDaysInYear(int year, int era)
        {
            if (era == 0 || era == 1)
            {
                if (year >= 1 && year <= 9999)
                {
                    if (year % 4 != 0 || (year % 100 == 0 && year % 400 != 0))
                    {
                        return 365;
                    }

                    return 366;
                }

                throw new ArgumentOutOfRangeException(
                    "year",
                    string.Format(CultureInfo.CurrentCulture, "ArgumentOutOfRange_Range 1,9999")
                );
            }

            throw new ArgumentOutOfRangeException("era", "ArgumentOutOfRange_InvalidEraValue");
        }

        public override int GetEra(DateTime time)
        {
            return 1;
        }

        public override int GetMonth(DateTime time)
        {
            return time.Month;
        }

        public override int GetMonthsInYear(int year, int era)
        {
            if (era == 0 || era == 1)
            {
                if (year >= 1 && year <= 9999)
                {
                    return 12;
                }

                throw new ArgumentOutOfRangeException(
                    "year",
                    string.Format(CultureInfo.CurrentCulture, "ArgumentOutOfRange_Range 1, 9999")
                );
            }

            throw new ArgumentOutOfRangeException("era", "ArgumentOutOfRange_InvalidEraValue");
        }

        public override int GetYear(DateTime time)
        {
            return time.Year;
        }

        public override bool IsLeapDay(int year, int month, int day, int era)
        {
            if (month < 1 || month > 12)
            {
                throw new ArgumentOutOfRangeException("month", "ArgumentOutOfRange_Range , 1, 12");
            }

            if (era != 0 && era != 1)
            {
                throw new ArgumentOutOfRangeException("era", "ArgumentOutOfRange_InvalidEraValue");
            }

            if (year < 1 || year > 9999)
            {
                throw new ArgumentOutOfRangeException("year", "ArgumentOutOfRange_Range 1, 9999");
            }

            if (day < 1 || day > GetDaysInMonth(year, month))
            {
                throw new ArgumentOutOfRangeException(
                    "day",
                    $"ArgumentOutOfRange_Range 1,{GetDaysInMonth(year, month)}"
                );
            }

            if (!IsLeapYear(year))
            {
                return false;
            }

            if (month == 12 && day == 30)
            {
                return true;
            }

            return false;
        }

        public override bool IsLeapMonth(int year, int month, int era)
        {
            if (era != 0 && era != 1)
            {
                throw new ArgumentOutOfRangeException("era", "ArgumentOutOfRange_InvalidEraValue");
            }

            if (year < 1 || year > 9999)
            {
                throw new ArgumentOutOfRangeException(
                    "year",
                    string.Format(CultureInfo.CurrentCulture, "ArgumentOutOfRange_Range", 1, 9999)
                );
            }

            if (month < 1 || month > 12)
            {
                throw new ArgumentOutOfRangeException("month", "ArgumentOutOfRange_Range 1, 12");
            }

            return false;
        }

        public override bool IsLeapYear(int year, int era)
        {
            if (era == 0 || era == 1)
            {
                if (year >= 1 && year <= 9999)
                {
                    if (year % 4 == 0)
                    {
                        if (year % 100 == 0)
                        {
                            return year % 400 == 0;
                        }

                        return true;
                    }

                    return false;
                }

                throw new ArgumentOutOfRangeException(
                    "year",
                    string.Format(CultureInfo.CurrentCulture, "ArgumentOutOfRange_Range", 1, 9999)
                );
            }

            throw new ArgumentOutOfRangeException("era", "ArgumentOutOfRange_InvalidEraValue");
        }

        public override DateTime ToDateTime(
            int year,
            int month,
            int day,
            int hour,
            int minute,
            int second,
            int millisecond,
            int era
        )
        {
            if (era == 0 || era == 1)
            {
                return new DateTime(year, month, day, hour, minute, second, millisecond);
            }

            throw new ArgumentOutOfRangeException("ArgumentOutOfRange_InvalidEraValue");
        }

        public override int ToFourDigitYear(int year)
        {
            if (year < 0) { }

            if (year > 9999) { }

            return base.ToFourDigitYear(year);
        }

        public override int GetLeapMonth(int year, int era)
        {
            if (era != 0 && era != 1)
            {
                throw new ArgumentOutOfRangeException("era", "ArgumentOutOfRange_InvalidEraValue");
            }

            if (year < 1 || year > 9999)
            {
                throw new ArgumentOutOfRangeException(
                    "year",
                    string.Format(CultureInfo.CurrentCulture, "ArgumentOutOfRange_Range", 1, 9999)
                );
            }

            return 0;
        }

        public static long GetAbsoluteDate(int year, int month, int day)
        {
            if (year >= 1 && year <= 9999 && month >= 1 && month <= 12)
            {
                int[] array = (
                    (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0))
                        ? DaysToMonth366
                        : DaysToMonth365
                );
                if (day >= 1 && day <= array[month] - array[month - 1])
                {
                    int num = year - 1;
                    int num2 =
                        num * 365 + num / 4 - num / 100 + num / 400 + array[month - 1] + day - 1;
                    return num2;
                }
            }

            throw new ArgumentOutOfRangeException(null, "ArgumentOutOfRange_BadYearMonthDay");
        }

        public long DateToTicks(int year, int month, int day)
        {
            return GetAbsoluteDate(year, month, day) * 864000000000L;
        }
    }
}
