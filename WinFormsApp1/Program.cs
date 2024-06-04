using System.Diagnostics;
using System.Numerics;

namespace WinFormsApp1
{
    internal static class Program
    {
        private static readonly string[] Units =
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

        private static readonly string[] Tens =
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

        private static readonly string[] Teens =
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

        public static string NumberToKurdishText(BigInteger number)
        {
            if (number > 9223372036854775807)
                return "نەتوانرا بەرزتر لە 9,223,372,036,854,775,807 بێت";

            if (number < -9223372036854775807)
                return "نەتوانرا نزمتر لە 9,223,372,036,854,775,807- بێت";

            if (number == 0)
                return "سفر";

            if (number < 0)
                return "كه‌م " + ConvertNumberToKurdishText(BigInteger.Abs(number));

            return ConvertNumberToKurdishText(number).Trim();
        }

        private static string ConvertNumberToKurdishText(BigInteger number)
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

        private static string ProcessSegment(
            BigInteger number,
            string suffix = "",
            bool isNumberSegment = false
        )
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

        private static string ProcessHundreds(BigInteger number)
        {
            if (number == 0)
                return "";

            if (number == 1)
                return "سەد";

            return Units[(int)number] + " سەد";
        }

        private static string ProcessTensAndUnits(BigInteger number)
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

        [STAThread]
        static void Main()
        {
            Debug.WriteLine(NumberToKurdishText(20)); // بیست
            Debug.WriteLine(NumberToKurdishText(21)); // بیست و یەک
            Debug.WriteLine(NumberToKurdishText(24)); // بیست و چوار
            Debug.WriteLine(NumberToKurdishText(100)); // سەد
            Debug.WriteLine(NumberToKurdishText(101)); // سەد و یەک
            Debug.WriteLine(NumberToKurdishText(102)); // سەد و دوو
            Debug.WriteLine(NumberToKurdishText(107)); // سەد و حەوت
            Debug.WriteLine(NumberToKurdishText(108)); // سەد و هەشت
            Debug.WriteLine(NumberToKurdishText(109)); // سەد و نۆ
            Debug.WriteLine(NumberToKurdishText(110)); // سەد و دە
            Debug.WriteLine(NumberToKurdishText(111)); // سەد و یازده‌
            Debug.WriteLine(NumberToKurdishText(112)); // سەد و دوازده‌
            Debug.WriteLine(NumberToKurdishText(119)); // سەد و نۆزدە
            Debug.WriteLine(NumberToKurdishText(120)); // سەد و بیست
            Debug.WriteLine(NumberToKurdishText(121)); // سەد و بیست و یەک
            Debug.WriteLine(NumberToKurdishText(122)); // سەد و بیست و دوو
            Debug.WriteLine(NumberToKurdishText(123)); // سەد و بیست و سێ

            Debug.WriteLine(NumberToKurdishText(324)); // سێ سەد و بیست و چوار

            Debug.WriteLine(NumberToKurdishText(924)); // نۆ سەد و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(904)); // نۆ سەد و چوار
            Debug.WriteLine(NumberToKurdishText(914)); // نۆ سەد و چواردە

            Debug.WriteLine(NumberToKurdishText(1001)); // هەزار و یەک
            Debug.WriteLine(NumberToKurdishText(1012)); // هەزار و دوازده‌
            Debug.WriteLine(NumberToKurdishText(1022)); // هەزار و بیست و دوو

            Debug.WriteLine(NumberToKurdishText(1101)); // هەزار و سەد و یەک
            Debug.WriteLine(NumberToKurdishText(1115)); // هەزار و سەد و پازدە

            Debug.WriteLine(NumberToKurdishText(1124)); // هەزار و سەد و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(4124)); // چوار هەزار و سەد و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(8002)); // هەشت هەزار و دوو
            Debug.WriteLine(NumberToKurdishText(9016)); // نۆ هەزار و شازدە

            Debug.WriteLine(NumberToKurdishText(10001)); // ده‌ هه‌زار و یەک
            Debug.WriteLine(NumberToKurdishText(10002)); // ده‌ هه‌زار و دوو

            Debug.WriteLine(NumberToKurdishText(10012)); // ده‌ هه‌زار و دوازده‌

            Debug.WriteLine(NumberToKurdishText(10101)); // ده‌ هه‌زار و سەد و یەک
            Debug.WriteLine(NumberToKurdishText(10111)); // ده‌ هه‌زار و سەد و یازده‌

            Debug.WriteLine(NumberToKurdishText(10201)); // ده‌ هه‌زار و دوو سەد و یەک
            Debug.WriteLine(NumberToKurdishText(10213)); // ده‌ هه‌زار و دوو سەد و سێزده‌
            Debug.WriteLine(NumberToKurdishText(11224)); // یازده‌ هه‌زار و دوو سەد و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(400213)); // چوار سەد هەزار و دوو سەد و سێزده‌
            Debug.WriteLine(NumberToKurdishText(859213)); // هەشت سەد و پەنجا و نۆ هەزار و دوو سەد و سێزده‌
            Debug.WriteLine(NumberToKurdishText(999999)); // نۆ سەد و نه‌وه‌د و نۆ هەزار و نۆ سەد و نه‌وه‌د و نۆ
            Debug.WriteLine(NumberToKurdishText(1000000)); // یەک ملیۆن

            Debug.WriteLine(NumberToKurdishText(1000001)); // یەک ملیۆن و یەک
            Debug.WriteLine(NumberToKurdishText(1000002)); // یەک ملیۆن و دوو
            Debug.WriteLine(NumberToKurdishText(1000011)); // یەک ملیۆن و یازده‌
            Debug.WriteLine(NumberToKurdishText(1000012)); // یەک ملیۆن و دوازده‌
            Debug.WriteLine(NumberToKurdishText(1000024)); // یەک ملیۆن و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(1000100)); // یەک ملیۆن و سەد
            Debug.WriteLine(NumberToKurdishText(1000101)); // یەک ملیۆن و سەد و یەک
            Debug.WriteLine(NumberToKurdishText(1000102)); // یەک ملیۆن و سەد و دوو
            Debug.WriteLine(NumberToKurdishText(1000113)); // یەک ملیۆن و سەد و دووازده‌
            Debug.WriteLine(NumberToKurdishText(1000124)); // یەک ملیۆن و سەد و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(1000245)); // یەک ملیۆن و دوو سەد و چل و پێنج
            Debug.WriteLine(NumberToKurdishText(1001000)); // یەک ملیۆن و هەزار
            Debug.WriteLine(NumberToKurdishText(1001001)); // یەک ملیۆن و هەزار و یەک
            Debug.WriteLine(NumberToKurdishText(1001002)); // یەک ملیۆن و هەزار و دوو
            Debug.WriteLine(NumberToKurdishText(1001012)); // یەک ملیۆن و هەزار و دوازده‌
            Debug.WriteLine(NumberToKurdishText(1001024)); // یەک ملیۆن و هەزار و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(1001100)); // یەک ملیۆن و هەزار و سەد
            Debug.WriteLine(NumberToKurdishText(1001101)); // یەک ملیۆن و هەزار و سەد و یەک
            Debug.WriteLine(NumberToKurdishText(1001102)); // یەک ملیۆن و هەزار و سەد و دوو
            Debug.WriteLine(NumberToKurdishText(1001113)); // یەک ملیۆن و هەزار و سەد و سێزده‌‌
            Debug.WriteLine(NumberToKurdishText(1001124)); // یەک ملیۆن و هەزار و سەد و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(1001245)); // یەک ملیۆن و هەزار و دوو سەد و چل و پێنج
            Debug.WriteLine(NumberToKurdishText(1002000)); // یەک ملیۆن و دوو هەزار
            Debug.WriteLine(NumberToKurdishText(1002001)); // یەک ملیۆن و دوو هەزار و یەک
            Debug.WriteLine(NumberToKurdishText(1002002)); // یەک ملیۆن و دوو هەزار و دوو
            Debug.WriteLine(NumberToKurdishText(1002012)); // یەک ملیۆن و دوو هەزار و دوازده‌
            Debug.WriteLine(NumberToKurdishText(1002024)); // یەک ملیۆن و دوو هەزار و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(1002100)); // یەک ملیۆن و دوو هەزار و سەد
            Debug.WriteLine(NumberToKurdishText(1002101)); // یەک ملیۆن و دوو هەزار و سەد و یەک
            Debug.WriteLine(NumberToKurdishText(1002102)); // یەک ملیۆن و دوو هەزار و سەد و دوو
            Debug.WriteLine(NumberToKurdishText(1002113)); // یەک ملیۆن و دوو هەزار و سەد و سێزده‌
            Debug.WriteLine(NumberToKurdishText(1002124)); // یەک ملیۆن و دوو هەزار و سەد و بیست و چوار
            Debug.WriteLine(NumberToKurdishText(1002245)); // یەک ملیۆن و دوو هەزار و دوو سەد و چل و پێنج
            Debug.WriteLine(NumberToKurdishText(1010001)); // یەک ملیۆن و ده‌ هه‌زار و یەک
            Debug.WriteLine(NumberToKurdishText(1020001)); // یەک ملیۆن و بیست و هه‌زار و یەک
            Debug.WriteLine(NumberToKurdishText(1120001)); // یەک ملیۆن و سه‌د و بیست هه‌زار و یەک
            Debug.WriteLine(NumberToKurdishText(1220001)); // یەک ملیۆن و دوو سه‌د و بیست هه‌زار و یەک
            Debug.WriteLine(NumberToKurdishText(2220001)); // دوو ملیۆن و دوو سه‌د و بیست هه‌زار و یەک
            Debug.WriteLine(NumberToKurdishText(3220001)); // سێ ملیۆن و دوو سه‌د و بیست هه‌زار و یەک
            Debug.WriteLine(NumberToKurdishText(4220001)); // چوار ملیۆن و دوو سه‌د و بیست هه‌زار و یەک
            Debug.WriteLine(NumberToKurdishText(10220001)); // ده‌ ملیۆن و دوو سه‌د و بیست هه‌زار و یەک
            Debug.WriteLine(NumberToKurdishText(100000000)); // سەد ملیۆن
            Debug.WriteLine(NumberToKurdishText(300000001)); // سێ سەد ملیۆن و یەک
            Debug.WriteLine(NumberToKurdishText(500000002)); // پێنج سەد ملیۆن و دوو

            Debug.WriteLine(NumberToKurdishText(999999999)); // نۆ سەد و نه‌وه‌د و نۆ ملیۆن و نۆ سەد و نه‌وه‌د و نۆ هەزار و نۆ سەد و نه‌وه‌د و نۆ
            Debug.WriteLine(NumberToKurdishText(99999999)); // نه‌وه‌د و نۆ ملیۆن و نۆ سەد و نه‌وه‌د و نۆ هەزار و نۆ سەد و نه‌وه‌د و نۆ
            Debug.WriteLine(NumberToKurdishText(9999999)); // نۆ ملیۆن و نۆ سەد و نه‌وه‌د و نۆ هەزار و نۆ سەد و نه‌وه‌د و نۆ
            Debug.WriteLine(NumberToKurdishText(999999)); // نۆ سەد و نه‌وه‌د و نۆ هەزار و نۆ سەد و نه‌وه‌د و نۆ
            Debug.WriteLine(NumberToKurdishText(99999)); // نه‌وه‌د و نۆ هەزار و نۆ سەد و نه‌وه‌د و نۆ
            Debug.WriteLine(NumberToKurdishText(9999)); // نۆ هەزار و نۆ سەد و نه‌وه‌د و نۆ
            Debug.WriteLine(NumberToKurdishText(999)); // نۆ سەد و نه‌وه‌د و نۆ
            Debug.WriteLine(NumberToKurdishText(99)); // نه‌وه‌د و نۆ
            Debug.WriteLine(NumberToKurdishText(9)); // نۆ

            Debug.WriteLine(NumberToKurdishText(1000000000)); // یەک ملیار
            Debug.WriteLine(NumberToKurdishText(1000000001)); // یەک ملیار و یەک
            Debug.WriteLine(NumberToKurdishText(1000000002)); // یەک ملیار و دوو
            Debug.WriteLine(NumberToKurdishText(1000000033)); // یەک ملیار و سی و سێ
            Debug.WriteLine(NumberToKurdishText(1000000333)); // یەک ملیار و سی سەد و سی و سێ
            Debug.WriteLine(NumberToKurdishText(1000003333)); // یەک ملیار و سێ هەزار و سی سەد و سی و سێ
            Debug.WriteLine(NumberToKurdishText(1000033333)); // یەک ملیار و سی و سێ هه‌زار و سی سەد و سی و سێ
            Debug.WriteLine(NumberToKurdishText(1000333333)); // یەک ملیار و سێ سەد و سی و سێ هه‌زار و سی سەد و سی و سێ
            Debug.WriteLine(NumberToKurdishText(1003333333)); // یەک ملیار و سێ ملیۆن و سێ سەد و سی و سێ هه‌زار و سی سەد و سی و سێ
            Debug.WriteLine(NumberToKurdishText(1033333333)); // یەک ملیار و سی و سێ ملیۆن و سێ سەد و سی و سێ هه‌زار و سی سەد و سی و سێ
            Debug.WriteLine(NumberToKurdishText(1333333333)); // یەک ملیار و سێ سەد و سی و سێ ملیۆن و سێ سەد و سی و سێ هه‌زار و سی سەد و سی و سێ
            Debug.WriteLine(NumberToKurdishText(3333333333)); // سێ ملیار و سێ سەد و سی و سێ ملیۆن و سێ سەد و سی و سێ هه‌زار و سی سەد و سی و سێ

            Debug.WriteLine(NumberToKurdishText(1000000000)); // یەک ملیار
            Debug.WriteLine(NumberToKurdishText(1000000001)); // یەک ملیار و یەک
            Debug.WriteLine(NumberToKurdishText(1000000011)); // یه‌ك ملیار و یازده‌
            Debug.WriteLine(NumberToKurdishText(1000000111)); // یه‌ك ملیار و سەد و یازده‌
            Debug.WriteLine(NumberToKurdishText(1000001111)); // یه‌ك ملیار و هه‌زار و سەد و یازده‌
            Debug.WriteLine(NumberToKurdishText(1000011111)); // یه‌ك ملیار و یازده‌ هه‌زار و سەد و یازده‌
            Debug.WriteLine(NumberToKurdishText(1000111111)); // یه‌ك ملیار و سەد و یازده‌ هه‌زار و سەد و یازده‌
            Debug.WriteLine(NumberToKurdishText(1001111111)); // یه‌ك ملیار و ملیۆنێك و سه‌د و یازده‌ هه‌زار و سەد و یازده‌
            Debug.WriteLine(NumberToKurdishText(1011111111)); // یه‌ك ملیار و یازده‌ ملیۆن و سه‌د و یازده‌ هه‌زار و سەد و یازده‌
            Debug.WriteLine(NumberToKurdishText(1111111111)); // یه‌ك ملیار و سه‌د و یازده‌ ملیۆن و سه‌د و یازده‌ هه‌زار و سەد و یازده‌

            Debug.WriteLine(NumberToKurdishText(241258569));

            Debug.WriteLine(NumberToKurdishText(1000000000000));
            Debug.WriteLine(NumberToKurdishText(1000000000001));
            Debug.WriteLine(NumberToKurdishText(1000000000011));
            Debug.WriteLine(NumberToKurdishText(1000000000111));
            Debug.WriteLine(NumberToKurdishText(1000000001111));
            Debug.WriteLine(NumberToKurdishText(1000000011111));
            Debug.WriteLine(NumberToKurdishText(1000000111111));
            Debug.WriteLine(NumberToKurdishText(1000001111111));
            Debug.WriteLine(NumberToKurdishText(1000011111111));
            Debug.WriteLine(NumberToKurdishText(1000111111111));
            Debug.WriteLine(NumberToKurdishText(1001111111111));
            Debug.WriteLine(NumberToKurdishText(1011111111111));
            Debug.WriteLine(NumberToKurdishText(1001111111111));

            Debug.WriteLine(NumberToKurdishText(9223372036854775808));
        }
    }
}
