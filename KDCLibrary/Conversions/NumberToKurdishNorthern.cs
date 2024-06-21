using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KDCLibrary.Conversions
{
    internal class NumberToKurdishNorthern
    {
        private readonly string[] Units =
        {
            "",
            "yek",
            "du",
            "sê",
            "çar",
            "pênc",
            "şeş",
            "heft",
            "heşt",
            "neh"
        };

        private readonly string[] Tens =
        {
            "",
            "deh",
            "bîst",
            "sih",
            "çel",
            "pêncî",
            "şêst",
            "heftê",
            "heştê",
            "nehvêd"
        };

        private readonly string[] Teens =
        {
            "deh",
            "Yanzdeh",
            "Duwanzdeh",
            "Sêzdeh",
            "çardeh",
            "panzdeh",
            "şanzdeh",
            "hevdeh",
            "hejdeh",
            "nozdeh"
        };

        public string Convert(long number)
        {
            var parts = new string[]
            {
                ProcessSegment((number / 1000000000000000000) % 1000, "quintillion ", true),
                ProcessSegment((number / 1000000000000000) % 1000, "qadrillion ", true),
                ProcessSegment((number / 1000000000000) % 1000, "trîlyon ", true),
                ProcessSegment((number / 1000000000) % 1000, "milyar ", true),
                ProcessSegment((number / 1000000) % 1000, "milyon ", true),
                ProcessSegment((number / 1000) % 1000, "hezar ", true),
                ProcessSegment(number % 1000)
            };

            var result = string.Join(" û ", parts.Where(part => !string.IsNullOrWhiteSpace(part)));

            if (result.Contains("milyar û yek milyon"))
            {
                result = result.Replace("milyar û yek milyon", "milyar û milyonek");
            }

            if (result.Contains("trîlyon û yek milyar"))
            {
                result = result.Replace("trîlyon û yek milyar", "trîlyon û milyarek");
            }

            if (result.Contains("quadrillion û yek trîlyon"))
            {
                result = result.Replace("quadrillion û yek trîlyon", "quadrillion û trîlyonek");
            }

            if (result.Contains("quintillion û yek quadrilyon"))
            {
                result = result.Replace(
                    "quintillion û yek quadrilyon",
                    "quintillion û quadrilyonek"
                );
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

            if (isNumberSegment && result.StartsWith("yek hezar"))
            {
                result = result.Replace("yek hezar", "hezar");
            }

            return result;
        }

        private string ProcessHundreds(long number)
        {
            if (number == 0)
                return "";

            if (number == 1)
                return "sed";

            return Units[(int)number] + "sed ";
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

            return Tens[(int)tens] + " û " + Units[(int)units];
        }
    }
}
