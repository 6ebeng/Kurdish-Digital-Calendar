using System.Runtime.InteropServices;
using ExcelDna.Integration;
using KDCLibrary;

namespace KDC_Excel
{
    [ComVisible(true)]
    public class UDFs
    {
        private static readonly IKDCService kdcService = new KDCServiceImplementation();

        [ExcelFunction(Description = "Converts a number to Kurdish words")]
        public static string ConvertNumberToKurdishText(double number)
        {
            long longNumber = (long)number;
            return kdcService.ConvertNumberToKurdishText(longNumber);
        }

        [ExcelFunction(Description = "Converts a date to Kurdish")]
        public static string ConvertDateToKurdish(
            string date,
            bool isKurdishCentral,
            string fromFormat,
            string toFormat,
            string targetCalendar,
            bool isAddSuffix
        )
        {
            if (isKurdishCentral)
                return kdcService.ConvertDateBasedOnUserSelection(
                    date,
                    false,
                    "Kurdish (Central)",
                    fromFormat,
                    toFormat,
                    targetCalendar,
                    isAddSuffix
                );
            else
                return kdcService.ConvertDateBasedOnUserSelection(
                    date,
                    false,
                    "Kurdish (Northern)",
                    fromFormat,
                    toFormat,
                    targetCalendar,
                    isAddSuffix
                );
        }
    }
}
