using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KDC_Word.Models
{
    internal class CalendarModel
    {
        public List<string> Dialects { get; private set; }
        public List<string> Formats { get; private set; }
        public List<string> CalendarGroup1 { get; private set; }
        public List<string> CalendarGroup2 { get; private set; }

        public CalendarModel()
        {
            InitializeDialects();
            InitializeFormats();
            InitializeCalendarGroups();
        }

        private void InitializeDialects()
        {
            Dialects = new List<string> { "Kurdish (Central)", "Kurdish (Northern)" };
        }

        private void InitializeFormats()
        {
            Formats = new List<string>
            {
                "dddd, dd MMMM, yyyy", "dddd, dd/MM/yyyy",
                "dd MMMM, yyyy", "dd/MM/yyyy",
                "MM/dd/yyyy", "yyyy/MM/dd"
            };
        }

        private void InitializeCalendarGroups()
        {
            CalendarGroup1 = new List<string> { "Gregorian", "Hijri", "Umm al-Qura" };
            CalendarGroup2 = new List<string>
            {
                "Gregorian (English)", "Gregorian (Arabic)", "Gregorian (Kurdish Central)", "Gregorian (Kurdish Northern)",
                "Hijri (English)", "Hijri (Arabic)", "Hijri (Kurdish Central)", "Hijri (Kurdish Northern)",
                "Umm al-Qura (English)", "Umm al-Qura (Arabic)", "Umm al-Qura (Kurdish Central)", "Umm al-Qura (Kurdish Northern)",
                "Kurdish (Central)", "Kurdish (Northern)"
            };
        }
    }
}
