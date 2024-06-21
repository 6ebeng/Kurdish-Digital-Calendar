using System;
using System.Linq;
using KDCLibrary.Conversions;

namespace KDCLibrary.Services
{
    internal class NumberToWordText
    {
        public string KurdishCentral(long number)
        {
            if (number > 9223372036854775807)
                return "نەتوانرا نزمتر/به‌رزتر لە 9,223,372,036,854,775,807 شیكار بێت";

            if (number < -9223372036854775807)
                return "نەتوانرا نزمتر/به‌رزتر لە 9,223,372,036,854,775,807 شیكار بێت";

            if (number == 0)
                return "سفر";

            if (number < 0)
                return "كه‌م " + new NumberToKurdishCentral().Convert(Math.Abs(number));

            return new NumberToKurdishCentral().Convert(number).Trim();
        }

        public string KurdishNorthern(long number)
        {
            if (number > 9223372036854775807)
                return "nikare ji 9,223,372,036,854,775,807 kêmtir/bilindtir be";

            if (number < -9223372036854775807)
                return "nikare ji 9,223,372,036,854,775,807 kêmtir/bilindtir be";

            if (number == 0)
                return "sifir";

            if (number < 0)
                return "kêm " + new NumberToKurdishNorthern().Convert(Math.Abs(number));

            return new NumberToKurdishNorthern().Convert(number).Trim();
        }
    }
}
