using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace KDCLibrary.Helpers
{
    [ComVisible(false)]
    public class CultureSetup
    {
        private void RegisterKurdishCentralCulture()
        {
            string cultureName = "ku-KRD"; // Custom culture name

            Unregister(cultureName);

            Console.WriteLine("Registering {0}", cultureName);
            try
            {
                CultureAndRegionInfoBuilder builder = new CultureAndRegionInfoBuilder(
                    cultureName,
                    CultureAndRegionModifiers.None
                );

                CultureInfo newCulture = new CultureInfo("en-US");
                RegionInfo regionInfo = new RegionInfo("US");

                DateTimeFormatInfo dtfi = GetKurdishCentralDTFI();

                newCulture.DateTimeFormat = dtfi;

                builder.LoadDataFromCultureInfo(newCulture);
                builder.LoadDataFromRegionInfo(regionInfo);

                // Set other properties as needed
                builder.CultureEnglishName = "Kurdish (Central)";
                builder.CultureNativeName = "كوردی (ناوه‌ڕاست)";
                builder.RegionEnglishName = "Kurdistan";
                builder.RegionNativeName = "کوردستان";
                builder.ThreeLetterISOLanguageName = "krd";
                builder.ThreeLetterISORegionName = "KRD";
                builder.ThreeLetterWindowsLanguageName = "KRD";
                builder.TwoLetterISOLanguageName = "ku";
                builder.TwoLetterISORegionName = "KU";
                builder.ThreeLetterWindowsRegionName = "KRD";
                builder.CurrencyEnglishName = "IQ Dinar";
                builder.CurrencyNativeName = "دیناری عێراقی";
                builder.ISOCurrencySymbol = "IQD";

                builder.IetfLanguageTag = cultureName;

                builder.NumberFormat.CurrencySymbol = "د.ع.";

                builder.Register(); // Register the new culture
            }
            catch (Exception ex)
            {
                Console.WriteLine("Registering the custom culture {0} failed", cultureName);
                Console.WriteLine(ex);
            }
            Console.WriteLine();
        }

        private void RegisterKurdishNorthernCulture()
        {
            string cultureName = "krm-KRD"; // Custom culture name
            Unregister(cultureName);

            Console.WriteLine("Registering {0}", cultureName);
            try
            {
                try
                {
                    CultureAndRegionInfoBuilder builder = new CultureAndRegionInfoBuilder(
                        cultureName,
                        CultureAndRegionModifiers.None
                    );

                    CultureInfo newCulture = new CultureInfo("en-US");
                    RegionInfo regionInfo = new RegionInfo("US");

                    DateTimeFormatInfo dtfi = GetKurdishNorthernDTFI();

                    newCulture.DateTimeFormat = dtfi;

                    builder.LoadDataFromCultureInfo(newCulture);
                    builder.LoadDataFromRegionInfo(regionInfo);

                    // Set other properties as needed
                    builder.CultureEnglishName = "Kurdish (Northern)";
                    builder.CultureNativeName = "Kurdî (Bakur)";
                    builder.RegionEnglishName = "Kurdistan";
                    builder.RegionNativeName = "Kurdistan";
                    builder.ThreeLetterISOLanguageName = "krd";
                    builder.ThreeLetterISORegionName = "KRD";
                    builder.ThreeLetterWindowsLanguageName = "KRD";
                    builder.TwoLetterISOLanguageName = "ku";
                    builder.TwoLetterISORegionName = "KU";
                    builder.ThreeLetterWindowsRegionName = "KRD";
                    builder.CurrencyEnglishName = "Lira";
                    builder.CurrencyNativeName = "Lira";
                    builder.ISOCurrencySymbol = "TRY";

                    builder.IetfLanguageTag = cultureName;

                    builder.NumberFormat.CurrencySymbol = "₺";

                    builder.Register(); // Register the new culture
                }
                catch (InvalidOperationException)
                {
                    // This is OK, means that this is already registered.
                    Console.WriteLine("The custom culture {0} was already registered", cultureName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Registering the custom culture {0} failed", cultureName);
                Console.WriteLine(ex);
            }
            Console.WriteLine();
        }

        private static void RegisterCustomCulture(
            string customCultureName,
            string parentCultureName
        )
        {
            Console.WriteLine("Registering {0}", customCultureName);
            try
            {
                CultureAndRegionInfoBuilder cib = new CultureAndRegionInfoBuilder(
                    customCultureName,
                    CultureAndRegionModifiers.None
                );
                CultureInfo ci = new CultureInfo(parentCultureName);

                cib.LoadDataFromCultureInfo(ci);

                RegionInfo ri = new RegionInfo(parentCultureName);
                cib.LoadDataFromRegionInfo(ri);
                cib.Register();
                Console.WriteLine("Success.");
            }
            catch (InvalidOperationException)
            {
                // This is OK, means that this is already registered.
                Console.WriteLine(
                    "The custom culture {0} was already registered",
                    customCultureName
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine("Registering the custom culture {0} failed", customCultureName);
                Console.WriteLine(ex);
            }
            Console.WriteLine();
        }

        public DateTimeFormatInfo GetKurdishCentralDTFI()
        {
            return new DateTimeFormatInfo
            {
                Calendar = new GregorianCalendar(),
                AbbreviatedDayNames = new string[]
                {
                    "یەکشەممە",
                    "دووشەممە",
                    "سێشەممە",
                    "چوارشەممە",
                    "پێنجشەممە",
                    "هەینی",
                    "شەممە"
                },
                ShortestDayNames = new string[] { "ی", "د", "س", "چ", "پ", "ھـ", "ش" },
                DayNames = new string[]
                {
                    "یەکشەممە",
                    "دووشەممە",
                    "سێشەممە",
                    "چوارشەممە",
                    "پێنجشەممە",
                    "هەینی",
                    "شەممە"
                },
                AbbreviatedMonthNames = new string[]
                {
                    "کانوونی یەکەم",
                    "شوبات",
                    "ئازار",
                    "نیسان",
                    "ئایار",
                    "حوزەیران",
                    "تەمموز",
                    "ئاب",
                    "ئەیلوول",
                    "تشرینی یەکەم",
                    "تشرینی دووەم",
                    "کانونی دووەم",
                    ""
                },

                MonthNames = new string[]
                {
                    "کانوونی یەکەم",
                    "شوبات",
                    "ئازار",
                    "نیسان",
                    "ئایار",
                    "حوزەیران",
                    "تەمموز",
                    "ئاب",
                    "ئەیلوول",
                    "تشرینی یەکەم",
                    "تشرینی دووەم",
                    "کانونی دووەم",
                    ""
                },

                AbbreviatedMonthGenitiveNames = new string[]
                {
                    "کانوونی یەکەم",
                    "شوبات",
                    "ئازار",
                    "نیسان",
                    "ئایار",
                    "حوزەیران",
                    "تەمموز",
                    "ئاب",
                    "ئەیلوول",
                    "تشرینی یەکەم",
                    "تشرینی دووەم",
                    "کانونی دووەم",
                    ""
                },

                MonthGenitiveNames = new string[]
                {
                    "کانوونی یەکەم",
                    "شوبات",
                    "ئازار",
                    "نیسان",
                    "ئایار",
                    "حوزەیران",
                    "تەمموز",
                    "ئاب",
                    "ئەیلوول",
                    "تشرینی یەکەم",
                    "تشرینی دووەم",
                    "کانونی دووەم",
                    ""
                },

                AMDesignator = "به‌یانی",
                PMDesignator = "ئێواره‌",
                FirstDayOfWeek = DayOfWeek.Saturday,
                CalendarWeekRule = CalendarWeekRule.FirstDay,
            };
        }

        public DateTimeFormatInfo GetKurdishNorthernDTFI()
        {
            return new DateTimeFormatInfo
            {
                Calendar = new GregorianCalendar(),
                AbbreviatedDayNames = new string[]
                {
                    "Yekşem",
                    "Duşem",
                    "Sêşem",
                    "Çarşem",
                    "Pêncşem",
                    "Înê",
                    "Şemî"
                },
                ShortestDayNames = new string[] { "Y", "D", "S", "Ç", "P", "Î", "Ş" },
                DayNames = new string[]
                {
                    "Yekşem",
                    "Duşem",
                    "Sêşem",
                    "Çarşem",
                    "Pêncşem",
                    "Înê",
                    "Şemî"
                },
                AbbreviatedMonthNames = new string[]
                {
                    "Çile",
                    "Şibat",
                    "Adar",
                    "Nîsan",
                    "Gulan",
                    "Pûşper",
                    "Tîrmeh",
                    "Tebax",
                    "Îlon",
                    "Cotmeh",
                    "Mijdar",
                    "Kanûn",
                    ""
                },

                MonthNames = new string[]
                {
                    "Çile",
                    "Şibat",
                    "Adar",
                    "Nîsan",
                    "Gulan",
                    "Pûşper",
                    "Tîrmeh",
                    "Tebax",
                    "Îlon",
                    "Cotmeh",
                    "Mijdar",
                    "Kanûn",
                    ""
                },

                AbbreviatedMonthGenitiveNames = new string[]
                {
                    "Çile",
                    "Şibat",
                    "Adar",
                    "Nîsan",
                    "Gulan",
                    "Pûşper",
                    "Tîrmeh",
                    "Tebax",
                    "Îlon",
                    "Cotmeh",
                    "Mijdar",
                    "Kanûn",
                    ""
                },

                MonthGenitiveNames = new string[]
                {
                    "Çile",
                    "Şibat",
                    "Adar",
                    "Nîsan",
                    "Gulan",
                    "Pûşper",
                    "Tîrmeh",
                    "Tebax",
                    "Îlon",
                    "Cotmeh",
                    "Mijdar",
                    "Kanûn",
                    ""
                },

                AMDesignator = "Sibê",
                PMDesignator = "Êvar",
                FirstDayOfWeek = DayOfWeek.Saturday,
                CalendarWeekRule = CalendarWeekRule.FirstDay,
            };
        }

        private void Unregister(string cultureName)
        {
            Console.WriteLine("Unregistering...");

            try
            {
                CultureAndRegionInfoBuilder.Unregister(cultureName);
                Console.WriteLine("Success");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error unregistering");
                Console.WriteLine(ex);
            }

            Console.WriteLine();
        }

        public CultureInfo CreateCultureInfoKurdishCentral()
        {
            string _cultureName = "ku-KRD";
            try
            {
                return new CultureInfo(_cultureName);
            }
            catch (ArgumentException)
            {
                try
                {
                    RegisterKurdishCentralCulture();
                    return new CultureInfo(_cultureName);
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show(
                        "Unable to register custom culture definition. You may need to run the application as an administrator."
                    );

                    return CultureInfo.InvariantCulture;
                }
            }
        }

        public CultureInfo CreateCultureInfoKurdishNorthern()
        {
            string _cultureName = "krm-KRD";
            try
            {
                return new CultureInfo(_cultureName);
            }
            catch (ArgumentException)
            {
                try
                {
                    RegisterKurdishNorthernCulture();
                    return new CultureInfo(_cultureName);
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show(
                        "Unable to register custom culture definition. You may need to run the application as an administrator."
                    );

                    return CultureInfo.InvariantCulture;
                }
            }
        }
    }
}
