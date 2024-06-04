using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KDCLibrary.Services
{
    internal class NumberToWordText
    {
        private readonly string[] Units =
        {
            "",
            "یەک",
            "دوو",
            "سێ",
            "چوار",
            "پێنج",
            "شەش",
            "حەوت",
            "هەشت",
            "نۆ"
        };

        private readonly string[] Tens =
        {
            "",
            "دە",
            "بیست",
            "سی",
            "چل",
            "پەنجا",
            "شەست",
            "حەفتا",
            "هەشتا",
            "نەوەد"
        };

        private readonly string[] Teens =
        {
            "دە",
            "یازده‌",
            "دوازده‌",
            "سێزده‌",
            "چواردە",
            "پازدە",
            "شازدە",
            "حەڤدە",
            "هەژدە",
            "نۆزدە"
        };

        public string ToKurdish(long number)
        {
            if (number > 9223372036854775807)
                return "نەتوانرا بەرزتر لە 9,223,372,036,854,775,807 بێت";

            if (number < -9223372036854775807)
                return "نەتوانرا نزمتر لە 9,223,372,036,854,775,807- بێت";

            if (number == 0)
                return "سفر";

            if (number < 0)
                return "كه‌م " + ConvertNumberToKurdishText(Math.Abs(number));

            return ConvertNumberToKurdishText(number).Trim();
        }

        private string ConvertNumberToKurdishText(long number)
        {
            var parts = new string[]
            {
                ProcessSegment((number / 1000000000000000000) % 1000, " کوینتیلیۆن", true),
                ProcessSegment((number / 1000000000000000) % 1000, " کوادریلیۆن", true),
                ProcessSegment((number / 1000000000000) % 1000, " تریلیۆن", true),
                ProcessSegment((number / 1000000000) % 1000, " ملیار", true),
                ProcessSegment((number / 1000000) % 1000, " ملیۆن", true),
                ProcessSegment((number / 1000) % 1000, " هەزار", true),
                ProcessSegment(number % 1000)
            };

            var result = string.Join(" و ", parts.Where(part => !string.IsNullOrWhiteSpace(part)));

            if (result.Contains("ملیار و یەک ملیۆن"))
            {
                result = result.Replace("ملیار و یەک ملیۆن", "ملیار و ملیۆنێك");
            }

            if (result.Contains("تریلیۆن و یەک ملیار"))
            {
                result = result.Replace("تریلیۆن و یەک ملیار", "تریلیۆن و ملیارێك");
            }

            if (result.Contains("کوادریلیۆن و یەک تریلیۆن"))
            {
                result = result.Replace("کوادریلیۆن و یەک تریلیۆن", "کوادریلیۆن و تریلیۆنێك");
            }

            if (result.Contains("کوینتیلیۆن و یەک کوادریلیۆن"))
            {
                result = result.Replace("کوینتیلیۆن و یەک کوادریلیۆن", "کوینتیلیۆن و کوادریلیۆنێك");
            }

            return result;
        }

        private string ProcessSegment(long number, string suffix = "", bool isNumberSegment = false)
        {
            if (number == 0)
                return "";

            var parts = new string[]
            {
                ProcessHundreds(number / 100),
                ProcessTensAndUnits(number % 100)
            };

            var result =
                string.Join(" و ", parts.Where(part => !string.IsNullOrWhiteSpace(part))) + suffix;

            if (isNumberSegment && result.StartsWith("یەک هەزار"))
            {
                result = result.Replace("یەک هەزار", "هەزار");
            }

            return result;
        }

        private string ProcessHundreds(long number)
        {
            if (number == 0)
                return "";

            if (number == 1)
                return "سەد";

            return Units[(int)number] + " سەد";
        }

        private string ProcessTensAndUnits(long number)
        {
            if (number == 0)
                return "";

            if (number < 10)
                return Units[(int)number];

            if (number < 20)
                return Teens[(int)number - 10];

            var tens = number / 10;
            var units = number % 10;

            if (units == 0)
                return Tens[(int)tens];

            return Tens[(int)tens] + " و " + Units[(int)units];
        }
    }
}
