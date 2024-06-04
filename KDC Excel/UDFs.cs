using System.Runtime.InteropServices;
using KDCLibrary;

namespace KDC_Excel
{
    [ComVisible(true)]
    public class UDFs
    {
        private static readonly IKDCService kdcService = new KDCServiceImplementation();

        public static string ConvertNumberToKurdishText(double number)
        {
            long longNumber = (long)number;
            return kdcService.ConvertNumberToKurdishText(longNumber);
        }
    }
}
