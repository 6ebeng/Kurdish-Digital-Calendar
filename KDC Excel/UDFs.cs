using System.Runtime.InteropServices;
using ExcelDna.Integration;
using KDCLibrary;

namespace KDC_Excel_UDFs
{
    [ComVisible(true)]
    public static class UDFs
    {
        private static readonly IKDCService kdcService = new KDCServiceImplementation();

        [ExcelFunction(
            Description = "Converts a number to Kurdish words",
            HelpTopic = "Converts a number to Kurdish words",
            Name = "ConvertNumberToKurdishText",
            Category = "Number to Word"
        )]
        public static string ConvertNumberToKurdishText(double number)
        {
            long longNumber = (long)number;
            return kdcService.ConvertNumberToKurdishText(longNumber);
        }

        [ExcelFunction(
            Description = "Converts a date to Kurdish",
            HelpTopic = "Converts a date to Kurdish",
            Name = "ConvertDateToKurdish",
            Category = "KDC"
        )]
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
