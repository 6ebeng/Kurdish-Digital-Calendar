using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kurdish_Digital_Calendar.DateConversionLibrary
{
    internal class InsertDate
    {
        public static string Kurdish(int formatChoice, string dialect, bool isAddSuffix)
        {
            DateTime todayGregorian = DateTime.Today; // Today's Gregorian date
            string todayKurdish = KurdishDate.fromGregorianToKurdish(todayGregorian, formatChoice, dialect, isAddSuffix);

            return todayKurdish;
        }
    }
}
