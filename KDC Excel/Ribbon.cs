using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Forms;
using KDCLibrary;
using KDCLibrary.UI;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace KDC_Excel
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        #region Intializers

        readonly IKDCService kDCService = new KDCServiceImplementation();
        private IRibbonUI ribbon;
        private const string IsReverseKeyName = Constants.KeyNames.IsReverse;
        private const string SelectedDialectKeyName = Constants.KeyNames.SelectedDialect;
        private const string SelectedFormat1KeyName = Constants.KeyNames.SelectedFormat1;
        private const string SelectedFormat2KeyName = Constants.KeyNames.SelectedFormat2;
        private const string LastSelectionGroup1KeyName = Constants.KeyNames.LastSelectionGroup1;
        private const string LastSelectionGroup2KeyName = Constants.KeyNames.LastSelectionGroup2;
        private const string isAddSuffixKeyName = Constants.KeyNames.IsAddSuffix;
        private const string isAutoUpdateKeyName = Constants.KeyNames.IsAutoUpdate;
        private const string InsertFormatKeyName = Constants.KeyNames.insertFormat;
        private const string isAutoUpdateOnLoadDocKeyName = Constants
            .KeyNames
            .IsAutoUpdateOnLoadDoc;

        private readonly List<string> _dialectsList = Constants.DefaultValues.Dialects;
        private readonly List<string> _formatsList = Constants.DefaultValues.Formats;
        private readonly List<string> _calendarGroup1List = Constants.DefaultValues.CalendarGroup1;
        private readonly List<string> _calendarGroup2List = Constants.DefaultValues.CalendarGroup2;
        private readonly string AppName = Assembly.GetExecutingAssembly().GetName().Name;

        private string SelectedDialect { get; set; }
        private string SelectedCalendar { get; set; }
        private string SelectedFromFormat { get; set; }
        private string SelectedToFormat { get; set; }
        private bool IsReverse { get; set; } = false;
        private bool IsAddSuffix { get; set; } = false;
        private bool IsAutoUpdate { get; set; } = false;
        private string SelectedInsertFormat { get; set; }
        public bool IsAutoUpdateOnLoadDoc { get; set; }

        #endregion

        public Ribbon() { }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return kDCService.GetRibbonXml();
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;

            IsReverse = Convert.ToBoolean(
                kDCService.LoadSetting(IsReverseKeyName, "false", AppName)
            );
            SelectedCalendar = kDCService.LoadSetting(
                IsReverse ? LastSelectionGroup2KeyName : LastSelectionGroup1KeyName,
                IsReverse ? _calendarGroup2List[0] : _calendarGroup1List[0],
                AppName
            );
            SelectedDialect = kDCService.LoadSetting(
                SelectedDialectKeyName,
                _dialectsList[0],
                AppName
            );
            SelectedFromFormat = kDCService.LoadSetting(
                SelectedFormat1KeyName,
                _formatsList[0],
                AppName
            );
            SelectedToFormat = kDCService.LoadSetting(
                SelectedFormat2KeyName,
                _formatsList[0],
                AppName
            );
            IsAddSuffix = Convert.ToBoolean(
                kDCService.LoadSetting(isAddSuffixKeyName, "false", AppName)
            );
            IsAutoUpdate = Convert.ToBoolean(
                kDCService.LoadSetting(isAutoUpdateKeyName, "false", AppName)
            );
            IsAutoUpdateOnLoadDoc = Convert.ToBoolean(
                kDCService.LoadSetting(isAutoUpdateOnLoadDocKeyName, "false", AppName)
            );
            SelectedInsertFormat = kDCService.LoadSetting(
                InsertFormatKeyName,
                _formatsList[0],
                AppName
            );
            ribbon.InvalidateControl("toggleButton1");
            ribbon.InvalidateControl("dropDown2");
        }

        #region Callbacks for Ribbon Controls

        public Bitmap OnGetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "splitButton1__btn":
                    return KDCLibrary.Properties.Resources.calendar;
                case "splitButton3__btn":
                    return KDCLibrary.Properties.Resources.help;
                case "splitButton2__btn":
                    return KDCLibrary.Properties.Resources.convert;
                case "button3":
                    return KDCLibrary.Properties.Resources.calendar;
                case "button2":
                    return KDCLibrary.Properties.Resources.update;
                case "button4":
                    return KDCLibrary.Properties.Resources.Choose_Date;
                default:
                    return null;
            }
        }

        public bool OnGetPressed(IRibbonControl control)
        {
            Debug.WriteLine("Check pressed: " + control.Id);
            switch (control.Id)
            {
                case "toggleButton1":
                    this.IsReverse = Convert.ToBoolean(
                        kDCService.LoadSetting(IsReverseKeyName, "false", AppName)
                    );
                    return IsReverse;
                case "checkBox1":
                    this.IsAddSuffix = Convert.ToBoolean(
                        kDCService.LoadSetting(isAddSuffixKeyName, "false", AppName)
                    );
                    return IsAddSuffix;
                case "checkBox2":
                    this.IsAutoUpdateOnLoadDoc = Convert.ToBoolean(
                        kDCService.LoadSetting(isAutoUpdateOnLoadDocKeyName, "false", AppName)
                    );
                    return IsAutoUpdateOnLoadDoc;
                case "checkBox3":
                    this.IsAutoUpdate = Convert.ToBoolean(
                        kDCService.LoadSetting(isAutoUpdateKeyName, "false", AppName)
                    );
                    return IsAutoUpdate;
                default:
                    return false;
            }
        }

        public void OnCheckAction_Click(IRibbonControl control, bool isPressed)
        {
            switch (control.Id)
            {
                case "checkBox1":
                    this.IsAddSuffix = isPressed;
                    kDCService.SaveSetting(isAddSuffixKeyName, isPressed.ToString(), AppName);
                    break;
                case "checkBox2":
                    this.IsAutoUpdateOnLoadDoc = isPressed;
                    kDCService.SaveSetting(
                        isAutoUpdateOnLoadDocKeyName,
                        isPressed.ToString(),
                        AppName
                    );
                    break;
                case "checkBox3":
                    this.IsAutoUpdate = isPressed;
                    kDCService.SaveSetting(isAutoUpdateKeyName, isPressed.ToString(), AppName);
                    break;
                case "toggleButton1":
                    this.IsReverse = isPressed;
                    kDCService.SaveSetting(IsReverseKeyName, IsReverse.ToString(), AppName);

                    // Exchange label of dropdown3 with dropdown4 and vice versa
                    (SelectedToFormat, SelectedFromFormat) = (SelectedFromFormat, SelectedToFormat);

                    kDCService.SaveSetting(SelectedFormat1KeyName, SelectedFromFormat, AppName);
                    kDCService.SaveSetting(SelectedFormat2KeyName, SelectedToFormat, AppName);

                    this.ribbon.InvalidateControl("dropDown3");
                    this.ribbon.InvalidateControl("dropDown4");
                    this.ribbon.InvalidateControl("dropDown2");
                    break;
            }
        }

        public void OnDropDownAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            switch (control.Id)
            {
                case "dropDown1":
                    this.SelectedDialect = _dialectsList[selectedIndex];
                    kDCService.SaveSetting(
                        SelectedDialectKeyName,
                        _dialectsList[selectedIndex],
                        AppName
                    );
                    break;

                case "dropDown2":
                    this.SelectedCalendar = IsReverse
                        ? _calendarGroup2List[selectedIndex]
                        : _calendarGroup1List[selectedIndex];
                    var keyName = IsReverse
                        ? LastSelectionGroup2KeyName
                        : LastSelectionGroup1KeyName;
                    kDCService.SaveSetting(keyName, SelectedCalendar, AppName);
                    break;

                case "dropDown3":
                    this.SelectedFromFormat = _formatsList[selectedIndex];
                    kDCService.SaveSetting(SelectedFormat1KeyName, SelectedFromFormat, AppName);
                    break;

                case "dropDown4":
                    this.SelectedToFormat = _formatsList[selectedIndex];
                    kDCService.SaveSetting(SelectedFormat2KeyName, SelectedToFormat, AppName);
                    break;
                case "dropDown5":
                    this.SelectedInsertFormat = _formatsList[selectedIndex];
                    kDCService.SaveSetting(InsertFormatKeyName, SelectedInsertFormat, AppName);
                    break;
            }
        }

        public void OnButtonAction_Click(IRibbonControl control)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;

            switch (control.Id)
            {
                case "button1":
                    new CreditsForm().ShowDialog();
                    break;
                case "button2": // Update Dates in Excel
                    if (activeWorkbook != null)
                    {
                        UpdateDatesFromCustomXmlPartsForExcel(activeWorkbook);
                    }
                    else
                    {
                        MessageBox.Show(
                            "No active workbook found.",
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }
                    break;
                case "button3":
                    PopulateInsertDate();
                    break;
                case "button4": // Choose Date
                    Form form = new CalendarControlForm(SelectedDialect);
                    form.FormClosed += (sender, e) =>
                    {
                        if (kDCService.isClosedByDoubleClick())
                        {
                            DateTime gDate = kDCService.GetGregorianSelectedDate();

                            if (gDate == DateTime.MinValue)
                            {
                                return;
                            }

                            // Insert Kurdish Date with the determined formatChoice, dialect, and isAddSuffix
                            foreach (Excel.Range cell in selectedRange.Cells)
                            {
                                cell.Value = kDCService.ConvertDateBasedOnUserSelection(
                                    gDate.ToString("dd/MM/yyyy"),
                                    false,
                                    SelectedDialect,
                                    "dd/MM/yyyy",
                                    SelectedInsertFormat,
                                    "Gregorian",
                                    IsAddSuffix
                                );
                            }
                        }
                    };
                    form.ShowDialog();
                    break;
                case "splitButton2__btn": // Convert Dates
                    foreach (Excel.Range cell in selectedRange.Cells)
                    {
                        Object selectedText = cell.Value;

                        if (selectedText is DateTime dateTime)
                        {
                            selectedText = dateTime.ToString("d"); // Short date pattern
                        }
                        string dateString = selectedText.ToString();

                        string result = kDCService.ConvertDateBasedOnUserSelection(
                            dateString,
                            IsReverse,
                            SelectedDialect,
                            SelectedFromFormat,
                            SelectedToFormat,
                            SelectedCalendar,
                            IsAddSuffix
                        );

                        if (result != dateString && !string.IsNullOrEmpty(result))
                        {
                            cell.Interior.ColorIndex = Excel.Constants.xlNone; // Reset color if changed
                            cell.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone; // Reset border style
                            cell.Value = result; // Set new date value
                        }
                        else
                        {
                            cell.Interior.Color = ColorTranslator.ToOle(Color.Red); // Highlight the cell with red color if conversion failed
                        }
                    }
                    break;
            }
        }

        public int OnGetSelectedItemIndex(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "dropDown1":
                    this.SelectedDialect = kDCService.LoadSetting(
                        SelectedDialectKeyName,
                        _dialectsList[0],
                        AppName
                    );

                    return _dialectsList.IndexOf(SelectedDialect);
                case "dropDown2":
                    string savedCalendarKeyName = IsReverse
                        ? LastSelectionGroup2KeyName
                        : LastSelectionGroup1KeyName;
                    this.SelectedCalendar = kDCService.LoadSetting(
                        savedCalendarKeyName,
                        IsReverse ? _calendarGroup2List[0] : _calendarGroup1List[0],
                        AppName
                    );
                    List<string> selectedList = IsReverse
                        ? _calendarGroup2List
                        : _calendarGroup1List;
                    return selectedList.IndexOf(SelectedCalendar);
                case "dropDown3":
                    this.SelectedFromFormat = kDCService.LoadSetting(
                        SelectedFormat1KeyName,
                        _formatsList[0],
                        AppName
                    );
                    return _formatsList.IndexOf(SelectedFromFormat);
                case "dropDown4":
                    SelectedToFormat = kDCService.LoadSetting(
                        SelectedFormat2KeyName,
                        _formatsList[0],
                        AppName
                    );
                    return _formatsList.IndexOf(SelectedToFormat);
                case "dropDown5":
                    this.SelectedInsertFormat = kDCService.LoadSetting(
                        InsertFormatKeyName,
                        _formatsList[0],
                        AppName
                    );
                    return _formatsList.IndexOf(SelectedInsertFormat);
                default:
                    return 0;
            }
        }

        public int OnGetItemCount(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "dropDown1":
                    return _dialectsList.Count;
                case "dropDown2":
                    return IsReverse ? _calendarGroup2List.Count : _calendarGroup1List.Count;
                case "dropDown3":
                case "dropDown4":
                case "dropDown5":
                    return _formatsList.Count;
                default:
                    return 0;
            }
        }

        public string OnGetItemLabel(IRibbonControl control, int index)
        {
            switch (control.Id)
            {
                case "dropDown1":
                    return _dialectsList[index];
                case "dropDown2":
                    List<string> selectedList = IsReverse
                        ? _calendarGroup2List
                        : _calendarGroup1List;
                    return selectedList[index];
                case "dropDown3":
                case "dropDown4":
                case "dropDown5":
                    return _formatsList[index];
                default:
                    return "";
            }
        }

        #endregion

        #region Helper Methods
        private void CleanOrphanNamedRanges(Excel.Workbook workbook)
        {
            List<Excel.Name> namesToRemove = new List<Excel.Name>();

            foreach (Excel.Name name in workbook.Names)
            {
                if (name.Name.StartsWith("KDate") && IsNamedRangeOrphan(name))
                {
                    namesToRemove.Add(name);
                }
            }

            // Remove identified orphan named ranges
            foreach (Excel.Name name in namesToRemove)
            {
                name.Delete();
            }
        }

        private bool IsNamedRangeOrphan(Excel.Name namedRange)
        {
            try
            {
                Excel.Range range = namedRange.RefersToRange;
                // Assuming 'orphan' means the range is empty or not used
                if (range.Cells.Count == 1 && range.Value == null)
                {
                    return true;
                }
                return false;
            }
            catch
            {
                // If we cannot get the range, it might be referring to a non-existing range
                return true;
            }
        }

        private void PopulateInsertDate()
        {
            if (string.IsNullOrEmpty(SelectedInsertFormat))
            {
                MessageBox.Show(
                    "No format selected.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            // Determine the format choice based on the selected format.
            int formatChoice = kDCService.SelectFormatChoice(SelectedInsertFormat);
            if (formatChoice == -1) // If the format is unsupported or not found
            {
                MessageBox.Show(
                    "No valid format selected.",
                    "Warning",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation
                );
                return; // Exit the method if no valid format is selected
            }

            string kurdishDate = kDCService.Kurdish(formatChoice, SelectedDialect, IsAddSuffix);

            // Get the currently selected range in Excel
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection;

            // Insert the Kurdish date into each cell in the selected range
            foreach (Excel.Range cell in selectedRange.Cells)
            {
                cell.Value = kurdishDate;

                // If IsAutoUpdate is true, tag the cell for future updates
                if (IsAutoUpdate)
                {
                    CleanOrphanNamedRanges(Globals.ThisAddIn.Application.ActiveWorkbook);

                    string tagName = "KDate" + Guid.NewGuid().ToString().Replace("-", ""); // Generate a unique name for the cell
                    cell.Name = tagName; // Assign a unique name to the cell which can be used to identify it later for updates.

                    // Optionally store custom metadata if needed
                    AddCustomXmlPartForAutoUpdate(
                        Globals.ThisAddIn.Application.ActiveWorkbook,
                        tagName,
                        formatChoice,
                        SelectedDialect,
                        IsAddSuffix
                    );
                }
            }
        }

        private void AddCustomXmlPartForAutoUpdate(
            Excel.Workbook workbook,
            string tagName,
            int formatChoice,
            string dialect,
            bool isAddSuffix
        )
        {
            string customXml =
                $@"<KurdishDateInsertion>
                <CellInfo>
                    <TagName>{SecurityElement.Escape(tagName)}</TagName>
                    <FormatChoice>{formatChoice}</FormatChoice>
                    <Dialect>{SecurityElement.Escape(dialect)}</Dialect>
                    <IsAddSuffix>{isAddSuffix.ToString().ToLower()}</IsAddSuffix>
                </CellInfo>
           </KurdishDateInsertion>";

            workbook.CustomXMLParts.Add(customXml);
        }

        public void UpdateDatesFromCustomXmlPartsForExcel(Excel.Workbook workbook)
        {
            // Assuming a specific named range or mechanism to identify cells with dates to update
            foreach (Excel.Name name in workbook.Names)
            {
                if (name.Name.Contains("KDate"))
                {
                    Excel.Range range = workbook.Application.Range[name.RefersTo];
                    CustomXMLPart part = FindCustomXmlPartForRange(name.Name, workbook);

                    if (part != null)
                    {
                        string dialect = part.SelectSingleNode(
                            "/KurdishDateInsertion/CellInfo/Dialect"
                        )?.Text;
                        int formatChoice = int.Parse(
                            part.SelectSingleNode(
                                "/KurdishDateInsertion/CellInfo/FormatChoice"
                            )?.Text ?? "0"
                        );
                        bool isAddSuffix = bool.Parse(
                            part.SelectSingleNode(
                                "/KurdishDateInsertion/CellInfo/IsAddSuffix"
                            )?.Text ?? "false"
                        );

                        string newDate = kDCService.Kurdish(formatChoice, dialect, isAddSuffix);
                        range.Value = newDate;
                    }
                    else
                    {
                        Debug.WriteLine($"No matching XML part found for range named: {name.Name}");
                    }
                }
            }
        }

        private CustomXMLPart FindCustomXmlPartForRangeForExcel(
            string nameRef,
            Excel.Workbook workbook
        )
        {
            foreach (CustomXMLPart part in workbook.CustomXMLParts)
            {
                if (part.NamespaceURI == string.Empty)
                {
                    var nameNode = part.SelectSingleNode("/KurdishDateInsertion/CellInfo/TagName");
                    if (nameNode != null && nameNode.Text == nameRef)
                    {
                        return part;
                    }
                }
            }
            return null;
        }

        #endregion

        #endregion
    }
}
