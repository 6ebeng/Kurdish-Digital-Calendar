using System.Collections.Specialized;
using System.Runtime.InteropServices;

namespace KDCLibrary
{
    [ComVisible(true)]
    public interface IKDCService
    {
        string ToKurdish(int formatChoice, string dialect, bool isAddSuffix);
        string ConvertDateBasedOnUserSelection(
            string selectedText,
            bool isReverse,
            string targetDialect,
            string fromFormat,
            string toFormat,
            string targetCalendar,
            bool isAddSuffix
        );

        string ConvertNumberToKurdishCentralText(long number);
        string ConvertNumberToKurdishNorthernText(long number);
    }
}
