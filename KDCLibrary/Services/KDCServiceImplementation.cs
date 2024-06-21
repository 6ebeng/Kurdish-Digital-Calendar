using System.Runtime.InteropServices;
using KDCLibrary.Calendars;
using KDCLibrary.Services;

namespace KDCLibrary
{
    [ComVisible(true)]
    public class KDCServiceImplementation : IKDCService
    {
        public string ToKurdish(int formatChoice, string dialect, bool isAddSuffix)
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

        public string ConvertNumberToKurdishCentralText(long number)
        {
            return new NumberToWordText().KurdishCentral(number);
        }

        public string ConvertNumberToKurdishNorthernText(long number)
        {
            return new NumberToWordText().KurdishNorthern(number);
        }
    }
}
