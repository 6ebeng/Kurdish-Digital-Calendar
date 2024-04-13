using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace KDCLibrary.Calendars
{
    internal class DateConversion
    {
        public string ConvertDateBasedOnUserSelection(
            string selectedText,
            bool isReverse,
            string targetDialect,
            string fromFormat,
            string toFormat,
            string targetCalendar,
            bool isAddSuffix
        )
        {
            string fromDateFormat = isReverse ? toFormat : fromFormat;
            string toDateFormat = isReverse ? fromFormat : toFormat;
            string calendarType = isReverse ? targetDialect : targetCalendar;
            DateTime parsedDate;
            DateTime targetDate;
            int formatChoice;
            string resultDate;
            string suffixString = "";

            String selectedTextCleaned = CleanSelectedText(selectedText);

            if (string.IsNullOrEmpty(selectedTextCleaned))
            {
                MessageBox.Show(
                    "No text selected.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return selectedText;
            }

            // Extract and remove suffix from selectedTextCleaned
            suffixString = ExtractSuffix(
                selectedTextCleaned,
                GetSuffixesBasedCalendarType(calendarType)
            );
            if (!string.IsNullOrEmpty(suffixString))
            {
                selectedTextCleaned = selectedTextCleaned.Replace(suffixString, "");
            }
            ;

            if (
                fromDateFormat == "dd/MM/yyyy"
                || fromDateFormat == "MM/dd/yyyy"
                || fromDateFormat == "yyyy/MM/dd"
            )
            {
                parsedDate = ExtractSimpleDate(selectedTextCleaned, fromDateFormat, calendarType);
            }
            else
            {
                parsedDate = ExtractComplexDate(selectedTextCleaned, fromDateFormat, calendarType);
            }

            if (parsedDate == DateTime.MinValue)
            {
                MessageBox.Show(
                    "Invalid date or format.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return selectedText;
            }

            targetDate = ConvertTargetCalendarToGregorian(parsedDate, calendarType);

            formatChoice = new Helper().SelectFormatChoice(toDateFormat);

            resultDate = isReverse
                ? ConvertGregorianToTargetCalendar(
                    targetDate,
                    formatChoice,
                    targetCalendar,
                    isAddSuffix
                )
                : new KurdishDate().FromGregorianToKurdish(
                    targetDate,
                    formatChoice,
                    targetDialect,
                    isAddSuffix
                );

            return resultDate;
        }

        public string ExtractSuffix(string text, List<string> suffixString)
        {
            string orignalSuffix = "";
            foreach (string suffix in suffixString)
            {
                if (!string.IsNullOrEmpty(text))
                {
                    if (text.Contains(suffix))
                    {
                        // check if the suffix is at the end of the string
                        if (text.Substring(text.Length - suffix.Length) == suffix)
                        {
                            orignalSuffix = suffix;
                            break;
                        }
                    }
                }
            }
            return orignalSuffix;
        }

        public List<string> GetSuffixesBasedCalendarType(string calendarType)
        {
            List<string> suffixStrings = new List<string> { };
            switch (calendarType)
            {
                case "Gregorian (English)":
                case "Gregorian (Arabic)":
                case "Gregorian (Kurdish Central)":
                case "Gregorian (Kurdish Northern)":
                case "Gregorian":
                    suffixStrings = new List<string>
                    {
                        " AD",
                        " م",
                        "ی زایینی",
                        " Zayînî",
                        " میلادی"
                    };
                    break;
                case "Hijri (English)":
                case "Hijri (Arabic)":
                case "Hijri (Kurdish Central)":
                case "Hijri (Kurdish Northern)":
                case "Umm al-Qura (English)":
                case "Umm al-Qura (Arabic)":
                case "Umm al-Qura (Kurdish Central)":
                case "Umm al-Qura (Kurdish Northern)":
                case "Umm al-Qura":
                case "Hijri":
                    suffixStrings = new List<string> { " AH", " هـ", "ی كۆچی", " Koçî", " هجری" };
                    break;
                case "Kurdish (Central)":
                case "Kurdish (Northern)":
                    suffixStrings = new List<string>
                    {
                        " كردی",
                        " ك",
                        "ی كوردی",
                        " Kurdî",
                        " Kurdi",
                        " Kurdish"
                    };
                    break;
            }
            return suffixStrings;
        }

        public DateTime ExtractComplexDate(
            string dateString,
            string dateFormat,
            string targetCalendarType
        )
        {
            int dayPart = 0,
                monthPart = 0,
                yearPart = 0;
            bool isDayParsed = false,
                isMonthParsed = false,
                isYearParsed = false;
            Dictionary<string, int> monthNames = new Dictionary<string, int>(
                StringComparer.OrdinalIgnoreCase
            );

            // This method will populate the dictionary based on the calendar type.
            PopulateMonthNames(monthNames, targetCalendarType);

            // Normalize Arabic comma to English comma in dateString
            dateString = dateString.Replace('،', ',');

            // Handle different complex formats
            switch (dateFormat)
            {
                case "dddd, dd MMMM, yyyy":
                    var tempParts = dateString.Split(
                        new string[] { ", " },
                        StringSplitOptions.None
                    );
                    if (tempParts.Length != 3)
                        return DateTime.MinValue; // Invalid format if not split into 3 parts

                    dateString = tempParts[1].Trim(); // Assume format includes day name prefix, which we ignore

                    var tempPartMD = dateString.Split(' ');
                    if (tempPartMD.Length != 2)
                        return DateTime.MinValue; // Invalid format if not split into 2 parts

                    isDayParsed = int.TryParse(tempPartMD[0], out dayPart);

                    if (monthNames.TryGetValue(FormatMonthName(tempPartMD[1]), out var monthValue))
                    {
                        monthPart = monthValue;
                        isMonthParsed = true; // Successfully retrieved month from dictionary
                    }

                    isYearParsed = int.TryParse(tempParts[2], out yearPart);
                    break;

                case "dddd, dd/MM/yyyy":
                    tempParts = dateString.Split(new string[] { ", " }, StringSplitOptions.None);
                    if (tempParts.Length != 2)
                        return DateTime.MinValue; // Invalid format if not split into 2 parts

                    dateString = tempParts[1]; // Remove day name

                    tempParts = dateString.Split('/');
                    if (tempParts.Length != 3)
                        return DateTime.MinValue; // Invalid format if not split into 3 parts

                    isDayParsed = int.TryParse(tempParts[0], out dayPart);
                    isMonthParsed = int.TryParse(tempParts[1], out monthPart);
                    isYearParsed = int.TryParse(tempParts[2], out yearPart);
                    break;

                case "dd MMMM, yyyy":
                    tempParts = dateString.Replace(",", "").Split(' ');
                    if (tempParts.Length != 3)
                        return DateTime.MinValue; // Invalid format if not split into 3 parts

                    isDayParsed = int.TryParse(tempParts[0], out dayPart);

                    if (monthNames.TryGetValue(FormatMonthName(tempParts[1]), out monthValue))
                    {
                        monthPart = monthValue;
                        isMonthParsed = true; // Successfully retrieved month from dictionary
                    }

                    isYearParsed = int.TryParse(tempParts[2], out yearPart);
                    break;

                default:
                    return DateTime.MinValue; // Unsupported format
            }

            if (!isDayParsed || !isMonthParsed || !isYearParsed)
            {
                // One or more components couldn't be parsed as integers
                return DateTime.MinValue;
            }

            if (ValidateDateComponents(dayPart, monthPart, yearPart, targetCalendarType))
            {
                return new DateTime(yearPart, monthPart, dayPart);
            }
            else
            {
                return DateTime.MinValue;
            }
        }

        private string FormatMonthName(string monthName)
        {
            // Ensure the first letter is uppercase and the rest are lowercase for consistent dictionary lookup
            if (monthName != null && monthName.Length > 1)
            {
                // char.ToUpper does not need CultureInfo
                return char.ToUpper(monthName[0])
                    + monthName.Substring(1).ToLower(CultureInfo.InvariantCulture);
            }
            else
            {
                return monthName?.ToUpper(CultureInfo.InvariantCulture);
            }
        }

        public DateTime ExtractSimpleDate(
            string dateString,
            string dateFormat,
            string targetCalendarType
        )
        {
            int dayPart,
                monthPart,
                yearPart;
            string[] dateParts;

            // Identify delimiter and split the date string
            string delimiter = IdentifyDelimiter(dateString);
            if (string.IsNullOrEmpty(delimiter))
            {
                return DateTime.MinValue; // Equivalent to '0' in VBA
            }

            dateParts = dateString.Split(new string[] { delimiter }, StringSplitOptions.None);

            if (dateParts.Length != 3) // Ensure array has three components
            {
                return DateTime.MinValue;
            }

            // Using int.TryParse for safer parsing
            bool isMonthParsed,
                isDayParsed,
                isYearParsed;
            switch (dateFormat)
            {
                case "MM/dd/yyyy":
                    isMonthParsed = int.TryParse(dateParts[0], out monthPart);
                    isDayParsed = int.TryParse(dateParts[1], out dayPart);
                    isYearParsed = int.TryParse(dateParts[2], out yearPart);
                    break;
                case "dd/MM/yyyy":
                    isDayParsed = int.TryParse(dateParts[0], out dayPart);
                    isMonthParsed = int.TryParse(dateParts[1], out monthPart);
                    isYearParsed = int.TryParse(dateParts[2], out yearPart);
                    break;
                case "yyyy/MM/dd":
                    isYearParsed = int.TryParse(dateParts[0], out yearPart);
                    isMonthParsed = int.TryParse(dateParts[1], out monthPart);
                    isDayParsed = int.TryParse(dateParts[2], out dayPart);
                    break;
                default:
                    // For unsupported formats, return DateTime.MinValue
                    return DateTime.MinValue;
            }

            if (!isDayParsed || !isMonthParsed || !isYearParsed)
            {
                // One or more components couldn't be parsed as integers
                return DateTime.MinValue;
            }

            if (ValidateDateComponents(dayPart, monthPart, yearPart, targetCalendarType))
            {
                return new DateTime(yearPart, monthPart, dayPart);
            }
            else
            {
                return DateTime.MinValue;
            }
        }

        private string IdentifyDelimiter(string text)
        {
            // Define an array of known delimiters
            char[] delimiters = new char[] { '/', '-', '\\', '_', '\t', '\u060C' }; // Including tab and Arabic comma

            foreach (var delimiter in delimiters)
            {
                if (text.Contains(delimiter.ToString()))
                {
                    return delimiter.ToString();
                }
            }

            return ""; // Return empty if no delimiter found
        }

        public bool ValidateDateComponents(int day, int month, int year, string calendarType)
        {
            int daysInMonth = 0;
            bool isLeapYear = false;

            switch (calendarType)
            {
                case "Gregorian":
                case "Gregorian (English)":
                case "Gregorian (Arabic)":
                case "Gregorian (Kurdish Central)":
                case "Gregorian (Kurdish Northren)":
                    if (month < 1 || month > 12 || year < 1)
                        return false;

                    switch (month)
                    {
                        case 1:
                        case 3:
                        case 5:
                        case 7:
                        case 8:
                        case 10:
                        case 12:
                            daysInMonth = 31;
                            break;
                        case 4:
                        case 6:
                        case 9:
                        case 11:
                            daysInMonth = 30;
                            break;
                        case 2:
                            isLeapYear = (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
                            daysInMonth = isLeapYear ? 29 : 28;
                            break;
                        default:
                            return false;
                    }
                    return day >= 1 && day <= daysInMonth;

                case "Hijri (English)":
                case "Hijri (Arabic)":
                case "Hijri (Kurdish Central)":
                case "Hijri (Kurdish Northern)":
                case "Umm al-Qura (English)":
                case "Umm al-Qura (Arabic)":
                case "Umm al-Qura (Kurdish Central)":
                case "Umm al-Qura (Kurdish Northern)":
                case "Umm al-Qura":
                case "Hijri":
                    if (month < 1 || month > 12 || year < 1)
                        return false;

                    isLeapYear = new[] { 2, 5, 7, 10, 13, 16, 18, 21, 24, 26, 29 }.Contains(
                        year % 30
                    );
                    daysInMonth =
                        month == 12 ? (isLeapYear ? 30 : 29) : ((month % 2 == 0) ? 29 : 30);
                    return day >= 1 && day <= daysInMonth;

                case "Kurdish (Central)":
                case "Kurdish (Northern)":
                    if (month < 1 || month > 12 || year < 2622)
                        return false; // Kurdish calendar year for Gregorian year 1

                    isLeapYear =
                        ((year + 1) % 4 == 0 && (year + 1) % 100 != 0) || ((year + 1) % 400 == 0);
                    switch (month)
                    {
                        case 1:
                        case 2:
                        case 3:
                        case 4:
                        case 5:
                        case 6:
                            daysInMonth = 31;
                            break;
                        case 7:
                        case 8:
                        case 9:
                        case 10:
                        case 11:
                            daysInMonth = 30;
                            break;
                        case 12:
                            daysInMonth = isLeapYear ? 29 : 28;
                            break;
                    }
                    return day >= 1 && day <= daysInMonth;

                default:
                    return false;
            }
        }

        private void PopulateMonthNames(
            Dictionary<string, int> monthNames,
            string targetCalendarType
        )
        {
            // Clearing existing entries to avoid duplicates
            monthNames.Clear();

            switch (targetCalendarType)
            {
                case "Gregorian (English)":
                case "Gregorian (Arabic)":
                case "Gregorian (Kurdish Central)":
                case "Gregorian (Kurdish Northern)":
                case "Gregorian":

                    // English Months
                    monthNames.Add("January", 1);
                    monthNames.Add("February", 2);
                    monthNames.Add("March", 3);
                    monthNames.Add("April", 4);
                    monthNames.Add("May", 5);
                    monthNames.Add("June", 6);
                    monthNames.Add("July", 7);
                    monthNames.Add("August", 8);
                    monthNames.Add("September", 9);
                    monthNames.Add("October", 10);
                    monthNames.Add("November", 11);
                    monthNames.Add("December", 12);

                    // Arabic Months
                    monthNames.Add("يناير", 1); // January
                    monthNames.Add("فبراير", 2); // February
                    monthNames.Add("مارس", 3); // March
                    monthNames.Add("أبريل", 4); // April
                    monthNames.Add("مايو", 5); // May
                    monthNames.Add("يونيو", 6); // June
                    monthNames.Add("يوليو", 7); // July
                    monthNames.Add("أغسطس", 8); // August
                    monthNames.Add("سبتمبر", 9); // September
                    monthNames.Add("أكتوبر", 10); // October
                    monthNames.Add("نوفمبر", 11); // November
                    monthNames.Add("ديسمبر", 12); // December

                    // Kurdish Central Months
                    monthNames.Add("کانونی دووەم", 1); // January
                    monthNames.Add("شوبات", 2); // February
                    monthNames.Add("ئازار", 3); // March
                    monthNames.Add("نیسان", 4); // April
                    monthNames.Add("ئایار", 5); // May
                    monthNames.Add("حوزەیران", 6); // June
                    monthNames.Add("تەمموز", 7); // July
                    monthNames.Add("ئاب", 8); // August
                    monthNames.Add("ئەیلوول", 9); // September
                    monthNames.Add("تشرینی یەكەم", 10); // October
                    monthNames.Add("تشرینی دووەم", 11); // November
                    monthNames.Add("كانونی یەكەم", 12); // December

                    // Kurdish Northern Months
                    monthNames.Add("Kanûna duyem", 1); // January
                    monthNames.Add("Şubat", 2); // February
                    monthNames.Add("Adar", 3); // March
                    monthNames.Add("Nîsan", 4); // April
                    monthNames.Add("Ayar", 5); // May
                    monthNames.Add("Hezîran", 6); // June
                    monthNames.Add("Temûz", 7); // July
                    monthNames.Add("Ab", 8); // August
                    monthNames.Add("Eylol", 9); // September
                    monthNames.Add("Çiriya yekem", 10); // October
                    monthNames.Add("Çiriya duyem", 11); // November
                    monthNames.Add("Kanûna yekem", 12); // December

                    // extra cases for Kurdish Latin
                    monthNames.Add("Kanûnî Duyem", 1);
                    monthNames.Add("Kanûna Duyê", 1);

                    //monthNames.Add("Şubat", 2);
                    monthNames.Add("Sibat", 2);

                    //monthNames.Add("Nîsan", 4);

                    //monthNames.Add("Hezîran", 6);

                    monthNames.Add("Tîrmeh", 7);

                    monthNames.Add("Tebax", 8);

                    monthNames.Add("Eylûl", 9);
                    monthNames.Add("Êlon", 9);
                    monthNames.Add("Îlon", 9);

                    monthNames.Add("Cotmeh", 10);
                    monthNames.Add("Çiriya Êkê", 10);
                    monthNames.Add("Çiriya pêşîn", 10);

                    monthNames.Add("Çiriya Duyê", 11);
                    monthNames.Add("Çiriya paşîn", 11);
                    monthNames.Add("Teşrîn", 11);

                    monthNames.Add("Kanûnî Paşîn", 12);
                    monthNames.Add("Kanûnî Yekem", 12);
                    break;

                case "Kurdish (Central)":
                case "Kurdish (Northern)":
                    // Kurdish Latin months in Unicode characters (Northern dialect)
                    monthNames.Add("Nêwroz", 1); // March
                    monthNames.Add("Gullan", 2); // April
                    monthNames.Add("Avrêl", 3); // May
                    monthNames.Add("Pusper", 4); // June
                    monthNames.Add("Tîrmeh", 5); // July
                    monthNames.Add("Gelawêj", 6); // August
                    monthNames.Add("Rezber", 7); // September
                    monthNames.Add("Kewçêr", 8); // October
                    monthNames.Add("Sermawez", 9); // November
                    monthNames.Add("Bafranbar", 10); // December
                    monthNames.Add("Rêbendan", 11); // January
                    monthNames.Add("Reşemî", 12); // February

                    // Kurdish non-Latin months in Unicode characters (Central dialect)
                    monthNames.Add("نەورۆز", 1); // March
                    monthNames.Add("گوڵان", 2); // April
                    monthNames.Add("جۆزەردان", 3); // May
                    monthNames.Add("پووشپەڕ", 4); // June
                    monthNames.Add("گەلاوێژ", 5); // July
                    monthNames.Add("خەرمانان", 6); // August
                    monthNames.Add("ڕەزبەر", 7); // September
                    monthNames.Add("گەڵاڕێزان", 8); // October
                    monthNames.Add("سەرماوەز", 9); // November
                    monthNames.Add("بەفرانبار", 10); // December
                    monthNames.Add("ڕێبەندان", 11); // January
                    monthNames.Add("رەشەمێ", 12); // February

                    // Extra cases for Kurdish months Romanized
                    monthNames.Add("Xakelêwe", 1);
                    //monthNames.Add("Gullan", 2); // Also listed above, for completion
                    monthNames.Add("Zerdan", 3);
                    monthNames.Add("Cozerdan", 3); // Duplicate entry, consider your logic for handling duplicates
                    monthNames.Add("Pûşper", 4);
                    //monthNames.Add("Gelawêj", 5); // Also listed above
                    monthNames.Add("Xermanan", 6);
                    monthNames.Add("Beran", 7);
                    monthNames.Add("Razbar", 7); // Duplicate entry
                    monthNames.Add("Xezan", 8);
                    monthNames.Add("Khazalawar", 8); // Duplicate entry, consider logic for handling
                    monthNames.Add("Saran", 9);
                    monthNames.Add("Befran", 10);
                    monthNames.Add("Befranbar", 10); // Duplicate entry
                    //monthNames.Add("Rêbendan", 11);


                    //Extra cases for Kurdish months non-Latin
                    monthNames.Add("ئادار", 1);
                    monthNames.Add("ئاڤدار", 1);
                    monthNames.Add("خاكه‌ لێوه‌", 1);
                    monthNames.Add("ئاخه‌لێو", 1);
                    monthNames.Add("هه‌رمێپشكوان", 1);
                    monthNames.Add("ئاخلیڤه‌", 1);

                    monthNames.Add("بانه‌مه‌ڕ‌", 2);
                    monthNames.Add("شه‌سته‌بارانه‌‌", 2);
                    monthNames.Add("شه‌سته‌باران‌", 2);

                    monthNames.Add("بارانبڕان", 3);
                    monthNames.Add("به‌خته‌باران", 3);
                    monthNames.Add("جووتان", 3);

                    monthNames.Add("خێڤه‌", 4);
                    monthNames.Add("پووشكاڵ‌", 4);

                    monthNames.Add("خزیران", 5);
                    monthNames.Add("جۆخینان", 5);
                    monthNames.Add("مێوه‌گه‌نان", 5);
                    monthNames.Add("زه‌خیران", 5);
                    monthNames.Add("تیرمه‌هـ", 5);

                    monthNames.Add("ته‌باخ", 6);
                    monthNames.Add("گه‌ڵاڤێژ", 6);

                    monthNames.Add("مشتاخان", 7);
                    monthNames.Add("هه‌وه‌ڵپایز", 7);
                    monthNames.Add("كه‌وچه‌ڕێن", 7);

                    monthNames.Add("خه‌زانان", 8);
                    monthNames.Add("خه‌زه‌ڵوه‌ر", 8);
                    monthNames.Add("گه‌ڵاخه‌زان", 8);
                    monthNames.Add("سه‌رپه‌له‌", 8);
                    monthNames.Add("چڕیائێكی", 8);
                    monthNames.Add("به‌روبه‌ز", 8);

                    monthNames.Add("كه‌وبه‌دار", 9);
                    monthNames.Add("چریا دووێ", 9);

                    monthNames.Add("به‌فرا ئێكێ", 10);
                    monthNames.Add("هه‌وه‌ڵزستان", 10);

                    monthNames.Add("به‌فربارادووێ", 11);
                    monthNames.Add("چله‌", 11);

                    monthNames.Add("ڕه‌شه‌مه‌", 12);
                    monthNames.Add("ڕه‌شه‌ما‌", 12);
                    monthNames.Add("ڕه‌شه‌مه‌هـ‌", 12);
                    monthNames.Add("بازه‌به‌ران", 12);
                    monthNames.Add("گوڵه‌مانگ", 12);

                    // Extra cases for Kurdish months Latin
                    monthNames.Add("Adar", 3);
                    //monthNames.Add("Tîrmeh", 7);
                    //monthNames.Add("Kewçêr", 10);
                    monthNames.Add("Marrêşan", 11);
                    monthNames.Add("Berfanbar", 12);

                    break;

                case "Hijri (English)":
                case "Hijri (Arabic)":
                case "Hijri (Kurdish Central)":
                case "Hijri (Kurdish Northern)":
                case "Umm al-Qura (English)":
                case "Umm al-Qura (Arabic)":
                case "Umm al-Qura (Kurdish Central)":
                case "Umm al-Qura (Kurdish Northern)":
                case "Umm al-Qura":
                case "Hijri":
                    // English Hijri/Umm al-Qura Months
                    monthNames.Add("Muharram", 1);
                    monthNames.Add("Safar", 2);
                    monthNames.Add("Rabi' al-Awwal", 3);
                    monthNames.Add("Rabi' al-Thani", 4);
                    monthNames.Add("Jumada al-Awwal", 5);
                    monthNames.Add("Jumada al-Thani", 6);
                    monthNames.Add("Rajab", 7);
                    monthNames.Add("Sha'ban", 8);
                    monthNames.Add("Ramadan", 9);
                    monthNames.Add("Shawwal", 10);
                    monthNames.Add("Dhu al-Qi'dah", 11);
                    monthNames.Add("Dhu al-Hijjah", 12);

                    // Arabic Hijri/Umm al-Qura Months
                    monthNames.Add("محرم", 1);
                    monthNames.Add("صفر", 2);
                    monthNames.Add("ربيع الأول", 3);
                    monthNames.Add("ربيع الثاني", 4);
                    monthNames.Add("جمادى الأولى", 5);
                    monthNames.Add("جمادى الآخرة", 6);
                    monthNames.Add("رجب", 7);
                    monthNames.Add("شعبان", 8);
                    monthNames.Add("رمضان", 9);
                    monthNames.Add("شوال", 10);
                    monthNames.Add("ذو القعدة", 11);
                    monthNames.Add("ذو الحجة", 12);

                    // Kurdish Central Hijri/Umm al-Qura Months
                    monthNames.Add("موحەڕەم", 1);
                    monthNames.Add("سەفەر", 2);
                    monthNames.Add("ڕەبیعی یه‌كه‌م ", 3);
                    monthNames.Add("ڕەبیعی دووه‌م", 4);
                    monthNames.Add("جه‌مادی یه‌كه‌م", 5);
                    monthNames.Add("جه‌مادی دووه‌م", 6);
                    monthNames.Add("ڕەجەب", 7);
                    monthNames.Add("شەعبان", 8);
                    monthNames.Add("ڕەمەزان", 9);
                    monthNames.Add("شەوال", 10);
                    monthNames.Add("زولقەعدە", 11);
                    monthNames.Add("زولحیججە", 12);

                    // Kurdish Northern Hijri/Umm al-Qura Months
                    monthNames.Add("Muherem", 1);
                    monthNames.Add("Sefer", 2);
                    monthNames.Add("Rebî'ulewel", 3);
                    monthNames.Add("Rebî'usanî", 4);
                    monthNames.Add("Cumadalûla", 5);
                    monthNames.Add("Cumadasaniye", 6);
                    monthNames.Add("Receb", 7);
                    monthNames.Add("Şeban", 8);
                    monthNames.Add("Remezan", 9);
                    monthNames.Add("Şewel", 10);
                    monthNames.Add("Zîlqe'de", 11);
                    monthNames.Add("Zîlhice", 12);

                    // Extra cases for Arabic months
                    monthNames.Add("ٱلْمُحَرَّم", 1);
                    monthNames.Add("صَفَر", 2);
                    monthNames.Add("رَبِيع ٱلْأَوَّل", 3);
                    monthNames.Add("رَبِيع ٱلثَّانِي", 4);
                    monthNames.Add("رَبِيع ٱلْآخِر", 4);
                    monthNames.Add("جُمَادَىٰ ٱلْأُولَىٰ", 5);
                    monthNames.Add("جُمَادَىٰ ٱلثَّانِيَة", 6);
                    monthNames.Add("جُمَادَىٰ ٱلْآخِرَة", 6);
                    monthNames.Add("رَجَب", 7);
                    monthNames.Add("شَعْبَان", 8);
                    monthNames.Add("رَمَضَان", 9);
                    monthNames.Add("شَوَّال", 10);
                    monthNames.Add("ذُو ٱلْقَعْدَة", 11);
                    monthNames.Add("ذُو ٱلْحِجَّة", 12);
                    break;

                default:
                    // Handle unsupported calendar types or consider throwing an exception
                    break;
            }
        }

        private string CleanSelectedText(string selectedText)
        {
            return selectedText
                .Trim()
                .Replace("\a", "") // Bell
                .Replace("\n", "") // New line
                .Replace("\r", "") // Carriage return
                .Replace("\v", "") // Vertical tab
                .Replace("\f", "") // Form feed
                .Replace("\t", ""); // Tab
        }

        private DateTime ConvertTargetCalendarToGregorian(DateTime date, string calendarType)
        {
            switch (calendarType)
            {
                case "Gregorian":
                    return date;
                case "Gregorian (English)":
                    return date;
                case "Gregorian (Arabic)":
                    return date;
                case "Hijri":
                    return new HijriDate().FromHijriToGregorian(date).Date;
                case "Umm al-Qura":
                    return new UmmAlQuraDate().FromUmmAlQuraToGregorian(date).Date;
                case "Hijri (English)":
                    return new HijriDate().FromHijriToGregorian(date).Date;
                case "Hijri (Arabic)":
                    return new HijriDate().FromHijriToGregorian(date).Date;
                case "Umm al-Qura (English)":
                    return new UmmAlQuraDate().FromUmmAlQuraToGregorian(date).Date;
                case "Umm al-Qura (Arabic)":
                    return new UmmAlQuraDate().FromUmmAlQuraToGregorian(date).Date;
                case "Kurdish (Central)":
                    return new KurdishDate().FromKurdishToGregorian(
                        date.Day,
                        date.Month,
                        date.Year
                    );
                case "Kurdish (Northern)":
                    return new KurdishDate().FromKurdishToGregorian(
                        date.Day,
                        date.Month,
                        date.Year
                    );
                default:
                    MessageBox.Show(
                        "Unsupported calendar type selected.",
                        "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation
                    );
                    return date;
            }
        }

        private string ConvertGregorianToTargetCalendar(
            DateTime date,
            int formatChoice,
            string calendarType,
            bool isAddSuffix
        )
        {
            switch (calendarType)
            {
                case "Gregorian (English)":
                    return new GregorianDate().FormatGregorian(
                        date,
                        formatChoice,
                        "English",
                        isAddSuffix
                    );
                case "Gregorian (Arabic)":
                    return new GregorianDate().FormatGregorian(
                        date,
                        formatChoice,
                        "Arabic",
                        isAddSuffix
                    );
                case "Gregorian (Kurdish Central)":
                    return new GregorianDate().FormatGregorian(
                        date,
                        formatChoice,
                        "Kurdish (Central)",
                        isAddSuffix
                    );
                case "Gregorian (Kurdish Northern)":
                    return new GregorianDate().FormatGregorian(
                        date,
                        formatChoice,
                        "Kurdish (Northern)",
                        isAddSuffix
                    );
                case "Hijri (English)":
                    return new HijriDate().FromGregorianToHijri(
                        date,
                        formatChoice,
                        "English",
                        isAddSuffix
                    );
                case "Hijri (Arabic)":
                    return new HijriDate().FromGregorianToHijri(
                        date,
                        formatChoice,
                        "Arabic",
                        isAddSuffix
                    );
                case "Hijri (Kurdish Central)":
                    return new HijriDate().FromGregorianToHijri(
                        date,
                        formatChoice,
                        "Kurdish (Central)",
                        isAddSuffix
                    );
                case "Hijri (Kurdish Northern)":
                    return new HijriDate().FromGregorianToHijri(
                        date,
                        formatChoice,
                        "Kurdish (Northern)",
                        isAddSuffix
                    );
                case "Umm al-Qura (English)":
                    return new UmmAlQuraDate().FromGregorianToUmmAlQura(
                        date,
                        formatChoice,
                        "English",
                        isAddSuffix
                    );
                case "Umm al-Qura (Arabic)":
                    return new UmmAlQuraDate().FromGregorianToUmmAlQura(
                        date,
                        formatChoice,
                        "Arabic",
                        isAddSuffix
                    );
                case "Umm al-Qura (Kurdish Central)":
                    return new UmmAlQuraDate().FromGregorianToUmmAlQura(
                        date,
                        formatChoice,
                        "Kurdish (Central)",
                        isAddSuffix
                    );
                case "Umm al-Qura (Kurdish Northern)":
                    return new UmmAlQuraDate().FromGregorianToUmmAlQura(
                        date,
                        formatChoice,
                        "Kurdish (Northern)",
                        isAddSuffix
                    );
                case "Kurdish (Central)":
                    return new KurdishDate().FromGregorianToKurdish(
                        date,
                        formatChoice,
                        calendarType,
                        isAddSuffix
                    );
                case "Kurdish (Northern)":
                    return new KurdishDate().FromGregorianToKurdish(
                        date,
                        formatChoice,
                        calendarType,
                        isAddSuffix
                    );
                default:
                    MessageBox.Show(
                        "Unsupported calendar type selected.",
                        "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation
                    );
                    return date.ToString(calendarType);
            }
        }
    }
}
