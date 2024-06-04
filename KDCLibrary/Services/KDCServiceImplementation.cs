﻿using System.Numerics;
using System.Runtime.InteropServices;
using KDCLibrary.Calendars;
using KDCLibrary.Services;

namespace KDCLibrary
{
    [ComVisible(true)]
    public class KDCServiceImplementation : IKDCService
    {
        public string toKurdish(int formatChoice, string dialect, bool isAddSuffix)
        {
            return new DateInsertion().Kurdish(formatChoice, dialect, isAddSuffix);
        }

        public string ConvertDateBasedOnUserSelection(
            string selectedText,
            bool isReverse,
            string targetDialect,
            string fromFormat,
            string toFormat,
            string targetCalendar,
            bool isAddSuffix
        )
        {
            return new DateConversion().ConvertDateBasedOnUserSelection(
                selectedText,
                isReverse,
                targetDialect,
                fromFormat,
                toFormat,
                targetCalendar,
                isAddSuffix
            );
        }

        public string ConvertNumberToKurdishText(long number)
        {
            return new NumberToWordText().ToKurdish(number);
        }
    }
}
