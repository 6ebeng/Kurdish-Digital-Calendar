namespace KDCLibrary.Calendars
{
    internal class DateFormatting
    {
        public string FormatDate(
            int Day,
            int Month,
            int Year,
            int aWeekDay,
            int formatChoice,
            string monthName,
            string weekDayName,
            string separator,
            string suffix
        )
        {
            string dayFormatted = Day.ToString("00");
            string monthFormatted = Month.ToString("00");

            switch (formatChoice)
            {
                case 1:
                    return $"{weekDayName}{separator}{dayFormatted} {monthName}{separator}{Year}{suffix}";
                case 2:
                    return $"{weekDayName}{separator}{dayFormatted}/{monthFormatted}/{Year}{suffix}";
                case 3:
                    return $"{dayFormatted} {monthName}{separator}{Year}{suffix}";
                case 4:
                    return $"{dayFormatted}/{monthFormatted}/{Year}{suffix}";
                case 5:
                    return $"{monthFormatted}/{Year}{suffix}";
                case 6:
                    return $"{Year}{suffix}";
                case 7:
                    return $"{dayFormatted}/{monthFormatted}";
                case 8:
                    return monthName;
                case 9:
                    return dayFormatted;
                case 10:
                    return $"{monthFormatted}/{dayFormatted}/{Year}{suffix}";
                case 11:
                    return $"{Year}/{monthFormatted}/{dayFormatted}{suffix}";
                default:
                    return "Invalid Format Choice";
            }
        }
    }
}
