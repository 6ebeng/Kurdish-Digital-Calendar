using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KDCLibrary.Calendars
{
    internal class DateInsertion
    {
        public string Kurdish(int formatChoice, string dialect, bool isAddSuffix)
        {
            DateTime todayGregorian = DateTime.Today; // Today's Gregorian date
            string todayKurdish = new KurdishDate().FromGregorianToKurdish(todayGregorian, formatChoice, dialect, isAddSuffix);

            return todayKurdish;
        }
    }
}
