using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using KDCLibrary;

namespace KDC_Excel
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
        public static string ConvertNumberToKurdishText(double number, string langcode)
        {
            long longNumber = (long)number;

            if (langcode == "" || langcode == "ckb")
                return kdcService.ConvertNumberToKurdishCentralText(longNumber);

            if (langcode == "ku")
                return kdcService.ConvertNumberToKurdishNorthernText(longNumber);

            MessageBox.Show(
                "Invalid language code. Please enter langcode (ckb) for Kurdish Central (Optional) or (ku) for Kurdish Northern.",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return "";
        }

        [ExcelFunction(
            Description = "Converts a date to Kurdish",
            HelpTopic = "Converts a date to Kurdish",
            Name = "ConvertDateToKurdish",
            Category = "KDC"
        )]
        public static string ConvertDateToKurdish(
            string date,
            string targetDialect,
            string fromFormat,
            string toFormat,
            string targetCalendar,
            bool isAddSuffix
        )
        {
            if (targetDialect == "" || targetDialect == "Kurdish (Central)")
                return kdcService.ConvertDateBasedOnUserSelection(
                    date,
                    false,
                    targetDialect,
                    fromFormat,
                    toFormat,
                    targetCalendar,
                    isAddSuffix
                );

            if (targetDialect == "Kurdish (Northern)")
                return kdcService.ConvertDateBasedOnUserSelection(
                    date,
                    false,
                    targetDialect,
                    fromFormat,
                    toFormat,
                    targetCalendar,
                    isAddSuffix
                );

            if (targetDialect == "")
                MessageBox.Show(
                    "Invalid target dialect. Please enter targetDialect \"Kurdish (Central)\" or \"Kurdish (Northern)\".",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

            return "";
        }
    }
}
