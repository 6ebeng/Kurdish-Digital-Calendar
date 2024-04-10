using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KDCLibrary
{
    public class KDCConstants
    {
        public static class KeyNames
        {
            public const string IsReverse = "IsReverse";
            public const string SelectedDialect = "SelectedDialect";
            public const string SelectedFormat1 = "SelectedFormat1";
            public const string SelectedFormat2 = "SelectedFormat2";
            public const string LastSelectionGroup1 = "LastSelectionGroup1";
            public const string LastSelectionGroup2 = "LastSelectionGroup2";
            public const string CheckBoxStates = "CheckBoxStates";
            public const string IsAddSuffix = "IsAddSuffix";
        }

        public class DefaultValues
        {
            public static readonly List<string> Dialects = new List<string> { "Kurdish (Central)", "Kurdish (Northern)" };
            public static readonly List<string> Formats = new List<string>
            {
                "dddd, dd MMMM, yyyy",
                "dddd, dd/MM/yyyy",
                "dd MMMM, yyyy",
                "dd/MM/yyyy",
                "MM/dd/yyyy",
                "yyyy/MM/dd"
            };

            public static readonly List<string> CalendarGroup1 = new List<string> { "Gregorian", "Hijri", "Umm al-Qura" };
            public static readonly List<string> CalendarGroup2 = new List<string>
            {
                "Gregorian (English)", "Gregorian (Arabic)", "Gregorian (Kurdish Central)", "Gregorian (Kurdish Northern)",
                "Hijri (English)", "Hijri (Arabic)", "Hijri (Kurdish Central)", "Hijri (Kurdish Northern)",
                "Umm al-Qura (English)", "Umm al-Qura (Arabic)", "Umm al-Qura (Kurdish Central)", "Umm al-Qura (Kurdish Northern)",
                "Kurdish (Central)", "Kurdish (Northern)"
            };

        }

    }
}
