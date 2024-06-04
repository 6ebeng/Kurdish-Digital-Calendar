using System;
using System.Numerics;
using System.Runtime.InteropServices;

namespace KDCLibrary
{
    [ComVisible(true)]
    public interface IKDCService
    {
        string toKurdish(int formatChoice, string dialect, bool isAddSuffix);
        string ConvertDateBasedOnUserSelection(
            string selectedText,
            bool isReverse,
            string targetDialect,
            string fromFormat,
            string toFormat,
            string targetCalendar,
            bool isAddSuffix
        );

        string ConvertNumberToKurdishText(long number);
    }
}
