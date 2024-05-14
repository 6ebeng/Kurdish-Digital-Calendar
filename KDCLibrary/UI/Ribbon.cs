using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Forms;
using KDCLibrary.Calendars;
using KDCLibrary.CustomControls;
using KDCLibrary.Helpers;
using KDCLibrary.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Bookmark = Microsoft.Office.Interop.Word.Bookmark;
using Document = Microsoft.Office.Interop.Word.Document;
using Excel = Microsoft.Office.Interop.Excel;
using MSProject = Microsoft.Office.Interop.MSProject;
using Outlook = Microsoft.Office.Interop.Outlook;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Visio = Microsoft.Office.Interop.Visio;
using Word = Microsoft.Office.Interop.Word;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

namespace KDCLibrary
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        #region Intializers

        public static IRibbonUI ribbon;
        private const string IsReverseKeyName = "IsReverse";
        public const string SelectedDialectKeyName = "SelectedDialect";
        private const string SelectedFormat1KeyName = "SelectedFormat1";
        private const string SelectedFormat2KeyName = "SelectedFormat2";
        private const string LastSelectionGroup1KeyName = "LastSelectionGroup1";
        private const string LastSelectionGroup2KeyName = "LastSelectionGroup2";
        public const string isAddSuffixKeyName = "IsAddSuffix";
        private const string isAutoUpdateKeyName = "IsAutoUpdate";
        private const string InsertFormatKeyName = "insertFormat";
        public const string isAutoUpdateOnLoadDocKeyName = "IsAutoUpdateOnLoadDoc";
        public const string ThemeColorKeyName = "ThemeMode";

        public static readonly List<string> _dialectsList = new List<string>
        {
            "Kurdish (Central)",
            "Kurdish (Northern)"
        };

        public static readonly List<string> _themesList = new List<string> { "Light", "Dark" };

        private readonly List<string> _formatsList = new List<string>
        {
            "dddd, dd MMMM, yyyy",
            "dddd, dd/MM/yyyy",
            "dd MMMM, yyyy",
            "MMMM dd, yyyy",
            "dd/MM/yyyy",
            "MM/dd/yyyy",
            "yyyy/MM/dd",
            "MMMM yyyy",
            "MM/yyyy",
            "MMMM",
            "yyyy"
        };
        private readonly List<string> _calendarGroup1List = new List<string>
        {
            "Gregorian",
            "Hijri",
            "Umm al-Qura"
        };
        private readonly List<string> _calendarGroup2List = new List<string>
        {
            "Gregorian (English)",
            "Gregorian (Arabic)",
            "Gregorian (Kurdish Central)",
            "Gregorian (Kurdish Northern)",
            "Hijri (English)",
            "Hijri (Arabic)",
            "Hijri (Kurdish Central)",
            "Hijri (Kurdish Northern)",
            "Umm al-Qura (English)",
            "Umm al-Qura (Arabic)",
            "Umm al-Qura (Kurdish Central)",
            "Umm al-Qura (Kurdish Northern)",
            "Kurdish (Central)",
            "Kurdish (Northern)"
        };

        public static string AppName { set; get; }

        public static string SelectedDialect { get; set; }
        private string SelectedCalendar { get; set; }
        private string SelectedFromFormat { get; set; }
        private string SelectedToFormat { get; set; }
        private bool IsReverse { get; set; } = false;
        public static bool IsAddSuffix { get; set; } = false;
        private bool IsAutoUpdate { get; set; } = false;
        private string SelectedInsertFormat { get; set; }
        public static bool IsAutoUpdateOnLoadDoc { get; set; }
        public static string SelectedTheme { get; set; }

        public Outlook.Application OutlookApp { get; set; }
        public Word.Application WordApp { get; set; }
        public Excel.Application ExcelApp { get; set; }
        public PowerPoint.Application PowerPointApp { get; set; }
        public static MSProject.Application ProjectApp { get; set; }
        public static Visio.Application VisioApp { get; set; }

        #endregion

        public Ribbon() { }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return new Helper().GetResourceText("KDCLibrary.UI.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;

            IsReverse = Convert.ToBoolean(
                new RegistryHelper().LoadSetting(IsReverseKeyName, "false", AppName)
            );
            SelectedCalendar = new RegistryHelper().LoadSetting(
                IsReverse ? LastSelectionGroup2KeyName : LastSelectionGroup1KeyName,
                IsReverse ? _calendarGroup2List[0] : _calendarGroup1List[0],
                AppName
            );
            SelectedDialect = new RegistryHelper().LoadSetting(
                SelectedDialectKeyName,
                _dialectsList[0],
                AppName
            );
            SelectedFromFormat = new RegistryHelper().LoadSetting(
                SelectedFormat1KeyName,
                _formatsList[0],
                AppName
            );
            SelectedToFormat = new RegistryHelper().LoadSetting(
                SelectedFormat2KeyName,
                _formatsList[0],
                AppName
            );
            IsAddSuffix = Convert.ToBoolean(
                new RegistryHelper().LoadSetting(isAddSuffixKeyName, "false", AppName)
            );
            IsAutoUpdate = Convert.ToBoolean(
                new RegistryHelper().LoadSetting(isAutoUpdateKeyName, "false", AppName)
            );
            IsAutoUpdateOnLoadDoc = Convert.ToBoolean(
                new RegistryHelper().LoadSetting(isAutoUpdateOnLoadDocKeyName, "false", AppName)
            );
            SelectedInsertFormat = new RegistryHelper().LoadSetting(
                InsertFormatKeyName,
                _formatsList[0],
                AppName
            );

            SelectedTheme = new RegistryHelper().LoadSetting(ThemeColorKeyName, "Dark", AppName);

            ribbon.Invalidate();
        }

        #region Callbacks for Ribbon Controls

        public bool OnGetVisible(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "button2":
                case "checkBox2":
                case "checkBox3":
                    if (ProjectApp != null || VisioApp != null)
                        return false;

                    return true;

                default:
                    return true;
            }
        }

        public string OnGetLabel(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "label1":
                    return IsReverse
                        ? "                        to Target"
                        : "                      from Source";
                case "label2":
                    return IsReverse
                        ? "                    from Source (Kurdish)"
                        : "                      to Target (Kurdish)";
                //case "dropDown3":
                //    return IsReverse ? "to" : "from";
                //case "dropDown4":
                //    return IsReverse ? "from" : "to";
                default:
                    return "";
            }
        }

        public Bitmap OnGetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "splitButton1__btn":
                    return SelectedTheme == "Dark"
                        ? Properties.Resources.Help_Black
                        : Properties.Resources.Help_White;
                case "button5":
                    return SelectedTheme == "Dark"
                        ? Properties.Resources.Convert_Black
                        : Properties.Resources.Convert_White;
                case "button3":
                    return SelectedTheme == "Dark"
                        ? Properties.Resources.Insert_Black
                        : Properties.Resources.Insert_White;
                case "button2":
                    return SelectedTheme == "Dark"
                        ? Properties.Resources.Update_Black
                        : Properties.Resources.Update_White;
                case "button4":
                    return SelectedTheme == "Dark"
                        ? Properties.Resources.Choose_Black
                        : Properties.Resources.Choose_White;
                case "checkBox4":
                    return SelectedTheme == "Dark"
                        ? Properties.Resources.Reverse_Black
                        : Properties.Resources.Reverse_White;
                case "button6":
                    return SelectedTheme == "Dark"
                        ? Properties.Resources.Settings_Black
                        : Properties.Resources.Settings_White;
                case "button1":
                    return SelectedTheme == "Dark"
                        ? Properties.Resources.Credits_Black
                        : Properties.Resources.Credits_White;
                default:
                    return null;
            }
        }

        public bool OnGetPressed(IRibbonControl control)
        {
            Debug.WriteLine("Check pressed: " + control.Id);
            switch (control.Id)
            {
                case "checkBox4":
                    this.IsReverse = Convert.ToBoolean(
                        new RegistryHelper().LoadSetting(IsReverseKeyName, "false", AppName)
                    );
                    return IsReverse;
                case "checkBox1":
                    IsAddSuffix = Convert.ToBoolean(
                        new RegistryHelper().LoadSetting(isAddSuffixKeyName, "false", AppName)
                    );
                    return IsAddSuffix;
                case "checkBox2":
                    IsAutoUpdateOnLoadDoc = Convert.ToBoolean(
                        new RegistryHelper().LoadSetting(
                            isAutoUpdateOnLoadDocKeyName,
                            "false",
                            AppName
                        )
                    );
                    return IsAutoUpdateOnLoadDoc;
                case "checkBox3":
                    this.IsAutoUpdate = Convert.ToBoolean(
                        new RegistryHelper().LoadSetting(isAutoUpdateKeyName, "false", AppName)
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
                //case "checkBox1":
                //    IsAddSuffix = isPressed;
                //    new RegistryHelper().SaveSetting(
                //        isAddSuffixKeyName,
                //        isPressed.ToString(),
                //        AppName
                //    );
                //    break;
                //case "checkBox2":
                //    IsAutoUpdateOnLoadDoc = isPressed;
                //    new RegistryHelper().SaveSetting(
                //        isAutoUpdateOnLoadDocKeyName,
                //        isPressed.ToString(),
                //        AppName
                //    );
                //    break;
                case "checkBox3":
                    this.IsAutoUpdate = isPressed;
                    new RegistryHelper().SaveSetting(
                        isAutoUpdateKeyName,
                        isPressed.ToString(),
                        AppName
                    );
                    break;
                case "checkBox4": // Reverse the conversion direction
                    this.IsReverse = isPressed;
                    new RegistryHelper().SaveSetting(
                        IsReverseKeyName,
                        IsReverse.ToString(),
                        AppName
                    );

                    // Exchange label of dropdown3 with dropdown4 and vice versa

                    //(SelectedToFormat, SelectedFromFormat) = (SelectedFromFormat, SelectedToFormat);

                    new RegistryHelper().SaveSetting(
                        SelectedFormat1KeyName,
                        SelectedFromFormat,
                        AppName
                    );
                    new RegistryHelper().SaveSetting(
                        SelectedFormat2KeyName,
                        SelectedToFormat,
                        AppName
                    );

                    //this.ribbon.InvalidateControl("dropDown3");
                    //this.ribbon.InvalidateControl("dropDown4");
                    ribbon.InvalidateControl("dropDown2");
                    ribbon.InvalidateControl("label1");
                    ribbon.InvalidateControl("label2");
                    break;
            }
        }

        public void OnDropDownAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            switch (control.Id)
            {
                //case "dropDown1":
                //    SelectedDialect = _dialectsList[selectedIndex];
                //    new RegistryHelper().SaveSetting(
                //        SelectedDialectKeyName,
                //        _dialectsList[selectedIndex],
                //        AppName
                //    );
                //    break;

                case "dropDown2":
                    this.SelectedCalendar = IsReverse
                        ? _calendarGroup2List[selectedIndex]
                        : _calendarGroup1List[selectedIndex];
                    var keyName = IsReverse
                        ? LastSelectionGroup2KeyName
                        : LastSelectionGroup1KeyName;
                    new RegistryHelper().SaveSetting(keyName, SelectedCalendar, AppName);
                    break;

                case "dropDown3":
                    this.SelectedFromFormat = _formatsList[selectedIndex];
                    new RegistryHelper().SaveSetting(
                        SelectedFormat1KeyName,
                        SelectedFromFormat,
                        AppName
                    );
                    break;

                case "dropDown4":
                    this.SelectedToFormat = _formatsList[selectedIndex];
                    new RegistryHelper().SaveSetting(
                        SelectedFormat2KeyName,
                        SelectedToFormat,
                        AppName
                    );
                    break;
                case "dropDown5":
                    this.SelectedInsertFormat = _formatsList[selectedIndex];
                    new RegistryHelper().SaveSetting(
                        InsertFormatKeyName,
                        SelectedInsertFormat,
                        AppName
                    );
                    break;
            }
        }

        public int OnGetSelectedItemIndex(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "dropDown1":
                    SelectedDialect = new RegistryHelper().LoadSetting(
                        SelectedDialectKeyName,
                        _dialectsList[0],
                        AppName
                    );

                    return _dialectsList.IndexOf(SelectedDialect);
                case "dropDown2":
                    string savedCalendarKeyName = IsReverse
                        ? LastSelectionGroup2KeyName
                        : LastSelectionGroup1KeyName;
                    this.SelectedCalendar = new RegistryHelper().LoadSetting(
                        savedCalendarKeyName,
                        IsReverse ? _calendarGroup2List[0] : _calendarGroup1List[0],
                        AppName
                    );
                    List<string> selectedList = IsReverse
                        ? _calendarGroup2List
                        : _calendarGroup1List;
                    return selectedList.IndexOf(SelectedCalendar);
                case "dropDown3":
                    this.SelectedFromFormat = new RegistryHelper().LoadSetting(
                        SelectedFormat1KeyName,
                        _formatsList[0],
                        AppName
                    );
                    return _formatsList.IndexOf(SelectedFromFormat);
                case "dropDown4":
                    SelectedToFormat = new RegistryHelper().LoadSetting(
                        SelectedFormat2KeyName,
                        _formatsList[0],
                        AppName
                    );
                    return _formatsList.IndexOf(SelectedToFormat);
                case "dropDown5":
                    this.SelectedInsertFormat = new RegistryHelper().LoadSetting(
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

        public void OnButtonAction_Click(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "button1": // Open the credits form
                    new CreditsForm().ShowDialog();
                    break;
                case "button2": // Update all dates

                    if (OutlookApp != null)
                    {
                        UpdateDatesFromCustomXmlPartsForWordOutlook(
                            GetWordEditorItem(OutlookApp.ActiveInspector().CurrentItem)
                        );
                    }

                    if (WordApp != null)
                    {
                        UpdateDatesFromCustomXmlPartsForWordOutlook(WordApp.ActiveDocument);
                    }

                    if (ExcelApp != null)
                    {
                        UpdateDatesFromCustomXmlPartsForExcel(ExcelApp.ActiveWorkbook);
                    }

                    if (PowerPointApp != null)
                    {
                        UpdateDatesFromCustomXmlPartsForPowerPoint(
                            PowerPointApp.ActivePresentation
                        );
                    }

                    // Update dates in MS Project is not supported yet
                    // Update dates in Visio is not supported yet

                    break;
                case "button3":
                    PopulateInsertDate();
                    break;
                case "button4": // Open the calendar control form
                    Form form = new DateForm();
                    form.FormClosed += (sender, e) =>
                    {
                        if (CustomXtraEditorsCalendarControl._isClosedByDoubleClick)
                        {
                            DateTime gDate =
                                CustomXtraEditorsCalendarControl._gregorianSelectedDate;

                            if (gDate == DateTime.MinValue)
                            {
                                return;
                            }

                            if (OutlookApp != null)
                            {
                                Document wordEditor = GetWordEditorItem(
                                    OutlookApp.ActiveInspector().CurrentItem
                                );
                                if (wordEditor != null && wordEditor.Application.Selection != null)
                                {
                                    wordEditor.Application.Selection.Text =
                                        new DateConversion().ConvertDateBasedOnUserSelection(
                                            gDate.ToString("dd/MM/yyyy"),
                                            false,
                                            SelectedDialect,
                                            "dd/MM/yyyy",
                                            SelectedInsertFormat,
                                            "Gregorian",
                                            IsAddSuffix
                                        );
                                    wordEditor.Application.Selection.Collapse(
                                        WdCollapseDirection.wdCollapseEnd
                                    );
                                }
                            }

                            if (WordApp != null)
                            {
                                WordApp.Selection.Text =
                                    new DateConversion().ConvertDateBasedOnUserSelection(
                                        gDate.ToString("dd/MM/yyyy"),
                                        false,
                                        SelectedDialect,
                                        "dd/MM/yyyy",
                                        SelectedInsertFormat,
                                        "Gregorian",
                                        IsAddSuffix
                                    );
                                WordApp.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                            }

                            if (ExcelApp != null)
                            {
                                // Insert Kurdish Date with the determined formatChoice, dialect, and isAddSuffix
                                foreach (Excel.Range cell in ExcelApp.Selection.Cells)
                                {
                                    cell.Value =
                                        new DateConversion().ConvertDateBasedOnUserSelection(
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

                            if (PowerPointApp != null)
                            {
                                if (PowerPointApp.ActiveWindow.Selection == null)
                                {
                                    MessageBox.Show(
                                        "No selection detected.",
                                        "Info",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information
                                    );
                                    return;
                                }

                                switch (PowerPointApp.ActiveWindow.Selection.Type)
                                {
                                    case PpSelectionType.ppSelectionShapes:
                                        foreach (
                                            PowerPoint.Shape shape in PowerPointApp
                                                .ActiveWindow
                                                .Selection
                                                .ShapeRange
                                        )
                                        {
                                            if (shape.HasTextFrame == MsoTriState.msoTrue)
                                            {
                                                shape.TextFrame.TextRange.Text =
                                                    new DateConversion().ConvertDateBasedOnUserSelection(
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
                                        break;

                                    case PpSelectionType.ppSelectionText:
                                        var textRange = PowerPointApp
                                            .ActiveWindow
                                            .Selection
                                            .TextRange;
                                        PowerPoint.Shape parentShape =
                                            textRange.Parent as PowerPoint.Shape;
                                        if (parentShape != null)
                                        {
                                            parentShape.TextFrame.TextRange.Text =
                                                new DateConversion().ConvertDateBasedOnUserSelection(
                                                    gDate.ToString("dd/MM/yyyy"),
                                                    false,
                                                    SelectedDialect,
                                                    "dd/MM/yyyy",
                                                    SelectedInsertFormat,
                                                    "Gregorian",
                                                    IsAddSuffix
                                                );
                                        }
                                        break;

                                    default:
                                        MessageBox.Show(
                                            "Unsupported selection type.",
                                            "Error",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error
                                        );
                                        break;
                                }
                            }

                            if (ProjectApp != null)
                            {
                                var selection = ProjectApp.ActiveSelection;
                                var activeCell = ProjectApp.ActiveCell;
                                string activeFieldName = activeCell.FieldName; // This is the field related to the user's current selection/interaction.

                                if (selection.Tasks == null || selection.Tasks.Count == 0)
                                {
                                    MessageBox.Show(
                                        "Please select one or more tasks.",
                                        "No Task Selected",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Warning
                                    );
                                    return;
                                }

                                // Update only the field corresponding to the active cell across all selected tasks.
                                foreach (MSProject.Task task in selection.Tasks)
                                {
                                    if (task != null)
                                    {
                                        SetTaskFieldValue(
                                            task,
                                            activeFieldName,
                                            new DateConversion().ConvertDateBasedOnUserSelection(
                                                gDate.ToString("dd/MM/yyyy"),
                                                false,
                                                SelectedDialect,
                                                "dd/MM/yyyy",
                                                SelectedInsertFormat,
                                                "Gregorian",
                                                IsAddSuffix
                                            )
                                        );
                                    }
                                }
                            }

                            if (VisioApp != null)
                            {
                                // Get the active Visio window.
                                Visio.Window activeWindow = VisioApp.ActiveWindow;

                                // Check if there are selected shapes.
                                if (activeWindow.Selection.Count > 0)
                                {
                                    foreach (Visio.Shape shape in activeWindow.Selection)
                                    {
                                        // Insert text into the selected shape(s) with the determined formatChoice, dialect, and isAddSuffix
                                        shape.Text = gDate.ToString("dd/MM/yyyy");
                                    }
                                }
                                else
                                {
                                    MessageBox.Show(
                                        "No shapes selected.",
                                        "Information",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information
                                    );
                                }
                            }
                        }
                    };
                    form.ShowDialog();
                    break;
                case "button5": // Convert selected text

                    if (OutlookApp != null)
                    {
                        Document wordEditor = GetWordEditorItem(
                            OutlookApp.ActiveInspector().CurrentItem
                        );
                        if (wordEditor != null && wordEditor.Application.Selection != null)
                        {
                            wordEditor.Application.Selection.Text =
                                new DateConversion().ConvertDateBasedOnUserSelection(
                                    wordEditor.Application.Selection.Text,
                                    IsReverse,
                                    SelectedDialect,
                                    SelectedFromFormat,
                                    SelectedToFormat,
                                    SelectedCalendar,
                                    IsAddSuffix
                                );
                        }
                    }

                    if (WordApp != null)
                    {
                        WordApp.Selection.Text =
                            new DateConversion().ConvertDateBasedOnUserSelection(
                                WordApp.Selection.Text,
                                IsReverse,
                                SelectedDialect,
                                SelectedFromFormat,
                                SelectedToFormat,
                                SelectedCalendar,
                                IsAddSuffix
                            );
                    }

                    if (ExcelApp != null)
                    {
                        foreach (Excel.Range cell in ExcelApp.Selection.Cells)
                        {
                            Object selectedText = cell.Value;

                            if (selectedText is DateTime dateTime)
                            {
                                selectedText = dateTime.ToString("d"); // Short date pattern
                            }
                            string dateString = selectedText.ToString();

                            string result = new DateConversion().ConvertDateBasedOnUserSelection(
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
                    }

                    if (PowerPointApp != null)
                    {
                        if (PowerPointApp.ActiveWindow.Selection != null)
                        {
                            var selection = PowerPointApp.ActiveWindow.Selection;
                            switch (selection.Type)
                            {
                                case PpSelectionType.ppSelectionShapes:
                                    // Handle shape selections
                                    foreach (PowerPoint.Shape shape in selection.ShapeRange)
                                    {
                                        if (
                                            shape.HasTextFrame == MsoTriState.msoTrue
                                            && shape.TextFrame.HasText == MsoTriState.msoTrue
                                        )
                                        {
                                            // Apply conversion on the shape's text
                                            shape.TextFrame.TextRange.Text =
                                                new DateConversion().ConvertDateBasedOnUserSelection(
                                                    shape.TextFrame.TextRange.Text,
                                                    IsReverse,
                                                    SelectedDialect,
                                                    SelectedFromFormat,
                                                    SelectedToFormat,
                                                    SelectedCalendar,
                                                    IsAddSuffix
                                                );
                                        }
                                    }
                                    break;

                                case PpSelectionType.ppSelectionText:
                                    // Directly handle text selections without iterating over ShapeRange
                                    var textRange = selection.TextRange;
                                    if (textRange != null && textRange.Length > 0)
                                    {
                                        // Apply conversion on the selected text
                                        textRange.Text =
                                            new DateConversion().ConvertDateBasedOnUserSelection(
                                                textRange.Text,
                                                IsReverse,
                                                SelectedDialect,
                                                SelectedFromFormat,
                                                SelectedToFormat,
                                                SelectedCalendar,
                                                IsAddSuffix
                                            );
                                    }
                                    break;

                                default:
                                    MessageBox.Show(
                                        "Please select a shape or text.",
                                        "Selection Not Supported",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Warning
                                    );
                                    break;
                            }
                        }
                        else
                        {
                            MessageBox.Show(
                                "Nothing is selected.",
                                "No Selection",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                    }

                    if (ProjectApp != null)
                    {
                        var selection = ProjectApp.ActiveSelection;

                        if (
                            selection == null
                            || selection.Tasks == null
                            || selection.Tasks.Count == 0
                        )
                        {
                            MessageBox.Show(
                                "Please select one or more tasks.",
                                "Selection Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                            return;
                        }

                        var activeCell = ProjectApp.ActiveCell;
                        string fieldName = activeCell.FieldName;

                        foreach (MSProject.Task task in selection.Tasks)
                        {
                            if (task == null)
                                continue;

                            // Extract the text based on the active field name
                            string fieldText = GetTaskFieldValue(task, fieldName);
                            string result = new DateConversion().ConvertDateBasedOnUserSelection(
                                fieldText,
                                IsReverse,
                                SelectedDialect,
                                SelectedFromFormat,
                                SelectedToFormat,
                                SelectedCalendar,
                                IsAddSuffix
                            );

                            if (!string.IsNullOrEmpty(result))
                            {
                                SetTaskFieldValue(task, fieldName, result);
                            }
                        }
                    }

                    if (VisioApp != null)
                    {
                        // Get the active Visio window and its selection
                        Visio.Selection selection = VisioApp.ActiveWindow.Selection;

                        if (selection.Count == 0)
                        {
                            MessageBox.Show(
                                "No shapes selected.",
                                "Warning",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning
                            );
                            return;
                        }

                        foreach (Visio.Shape shape in selection)
                        {
                            // Check if the shape has text
                            if (!string.IsNullOrEmpty(shape.Text))
                            {
                                string selectedText = shape.Text;

                                string result =
                                    new DateConversion().ConvertDateBasedOnUserSelection(
                                        selectedText,
                                        IsReverse,
                                        SelectedDialect,
                                        SelectedFromFormat,
                                        SelectedToFormat,
                                        SelectedCalendar,
                                        IsAddSuffix
                                    );

                                Debug.WriteLine(result);

                                shape.Text = result;
                            }
                        }
                    }

                    break;
                case "button6":
                    new Settings().ShowDialog();
                    break;
            }
        }

        #endregion

        #region Helper Methods

        private string GetTaskFieldValue(MSProject.Task task, string fieldName)
        {
            // Ideally, extend this to cover more fields as per your requirement
            switch (fieldName)
            {
                case "Name":
                    return task.Name;
                case "Notes":
                    return task.Notes;
                //case "Resource Names":         // This field is remove all spaces causes parse error
                //    return task.ResourceNames;
                case "Text1":
                    return task.Text1;
                case "Text2":
                    return task.Text2;
                case "Text3":
                    return task.Text3;
                case "Text4":
                    return task.Text4;
                case "Text5":
                    return task.Text5;
                case "Text6":
                    return task.Text6;
                case "Text7":
                    return task.Text7;
                case "Text8":
                    return task.Text8;
                case "Text9":
                    return task.Text9;
                case "Text10":
                    return task.Text10;
                case "Text11":
                    return task.Text11;
                case "Text12":
                    return task.Text12;
                default:
                    return "";
            }
        }

        private void SetTaskFieldValue(MSProject.Task task, string fieldName, string value)
        {
            // Similarly, extend this method to handle other fields
            switch (fieldName)
            {
                case "Name":
                    task.Name = value;
                    break;
                case "Notes":
                    task.Notes = value;
                    break;
                case "Resource Names":
                    task.ResourceNames = value;
                    break;
                case "Text1":
                    task.Text1 = value;
                    break;
                case "Text2":
                    task.Text2 = value;
                    break;
                case "Text3":
                    task.Text3 = value;
                    break;
                case "Text4":
                    task.Text4 = value;
                    break;
                case "Text5":
                    task.Text5 = value;
                    break;
                case "Text6":
                    task.Text6 = value;
                    break;
                case "Text7":
                    task.Text7 = value;
                    break;
                case "Text8":
                    task.Text8 = value;
                    break;
                case "Text9":
                    task.Text9 = value;
                    break;
                case "Text10":
                    task.Text10 = value;
                    break;
                case "Text11":
                    task.Text11 = value;
                    break;
                case "Text12":
                    task.Text12 = value;
                    break;
                default:
                    break;
            }
        }

        private void CleanOrphanBookmarksOutlook(object outlookItem)
        {
            if (outlookItem == null)
            {
                MessageBox.Show(
                    "No Mail Item provided.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            Document wordEditor = GetWordEditorItem(outlookItem);
            if (wordEditor == null)
            {
                MessageBox.Show(
                    $"This {outlookItem} does not support Word editing.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            CleanOrphanBookmarksWord(wordEditor);
        }

        private void CleanOrphanBookmarksWord(Document wordEditor)
        {
            List<string> bookmarksToRemove = new List<string>();

            foreach (Bookmark mark in wordEditor.Bookmarks)
            {
                if (mark.Name.StartsWith("KDate"))
                {
                    // Check if the bookmark's content is valid
                    if (IsBookmarkOrphan(mark))
                    {
                        bookmarksToRemove.Add(mark.Name);
                    }
                }
            }

            // Remove identified orphan bookmarks
            foreach (string bookmarkName in bookmarksToRemove)
            {
                wordEditor.Bookmarks[bookmarkName].Delete();
            }
        }

        private bool IsBookmarkOrphan(Bookmark bookmark)
        {
            string content = bookmark.Range.Text;
            if (string.IsNullOrWhiteSpace(content)) // Assuming dates have '/' character
            {
                return true;
            }
            return false;
        }

        private void CleanOrphanNamedExcelRanges(Workbook workbook)
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

        public void PopulateInsertDate()
        {
            if (SelectedInsertFormat == null)
            {
                MessageBox.Show(
                    "No format selected.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            int formatChoice = new Helper().SelectFormatChoice(SelectedInsertFormat);
            if (formatChoice == -1)
            {
                MessageBox.Show(
                    "No valid format selected.",
                    "Warning",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation
                );
                return;
            }

            string kurdishDate = new DateInsertion().Kurdish(
                formatChoice,
                SelectedDialect,
                IsAddSuffix
            );

            try
            {
                if (OutlookApp != null)
                {
                    var inspector = OutlookApp.ActiveInspector();
                    if (inspector == null || inspector.CurrentItem == null)
                    {
                        MessageBox.Show(
                            "No active item found.",
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                        return;
                    }

                    ProcessItemInsertOutlook(inspector.CurrentItem, kurdishDate, formatChoice);
                }

                if (WordApp != null)
                {
                    if (IsAutoUpdate)
                    {
                        CleanOrphanBookmarksWord(WordApp.ActiveDocument); // Remove orphan bookmarks before inserting a new date

                        Range currentRange = WordApp.Selection.Range;
                        currentRange.Text = kurdishDate; // This replaces the selected text with the Kurdish date

                        int start = currentRange.Start;
                        int end = start + kurdishDate.Length;

                        currentRange.SetRange(start, end);

                        // Add a bookmark for future reference
                        try
                        {
                            string bookmarkName =
                                "KDate" + Guid.NewGuid().ToString().Replace("-", ""); // Clean and valid bookmark name
                            if (!WordApp.ActiveDocument.Bookmarks.Exists(bookmarkName))
                            {
                                WordApp.ActiveDocument.Bookmarks.Add(bookmarkName, currentRange);

                                // Save parameters to Custom XML Part
                                AddCustomXmlPartForWordOutlook(
                                    bookmarkName,
                                    SelectedDialect,
                                    formatChoice,
                                    IsAddSuffix,
                                    WordApp.ActiveDocument
                                );

                                // select inserted date
                                currentRange.Select();
                            }
                        }
                        catch (COMException ex)
                        {
                            MessageBox.Show(
                                "Failed to create bookmark: " + ex.Message,
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                        }
                    }
                    else
                    {
                        WordApp.Selection.Text = kurdishDate;
                        WordApp.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }
                }

                if (ExcelApp != null)
                {
                    // Get the currently selected range in Excel
                    Excel.Range selectedRange = ExcelApp.Selection;

                    // Insert the Kurdish date into each cell in the selected range
                    foreach (Excel.Range cell in selectedRange.Cells)
                    {
                        cell.Value = kurdishDate;

                        // If IsAutoUpdate is true, tag the cell for future updates
                        if (IsAutoUpdate)
                        {
                            CleanOrphanNamedExcelRanges(ExcelApp.ActiveWorkbook);

                            string tagName = "KDate" + Guid.NewGuid().ToString().Replace("-", ""); // Generate a unique name for the cell
                            cell.Name = tagName; // Assign a unique name to the cell which can be used to identify it later for updates.

                            // Optionally store custom metadata if needed
                            AddCustomXmlPartForExcel(
                                ExcelApp.ActiveWorkbook,
                                tagName,
                                formatChoice,
                                SelectedDialect,
                                IsAddSuffix
                            );
                        }
                    }
                }

                if (PowerPointApp != null)
                {
                    if (PowerPointApp.ActiveWindow.Selection == null)
                    {
                        MessageBox.Show(
                            "No selection detected.",
                            "Info",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return;
                    }

                    switch (PowerPointApp.ActiveWindow.Selection.Type)
                    {
                        case PpSelectionType.ppSelectionShapes:
                            foreach (
                                PowerPoint.Shape shape in PowerPointApp
                                    .ActiveWindow
                                    .Selection
                                    .ShapeRange
                            )
                            {
                                if (shape.HasTextFrame == MsoTriState.msoTrue)
                                {
                                    if (IsAutoUpdate)
                                    {
                                        // Tag the shape with custom metadata
                                        AddCustomTagForPowerPointShape(
                                            shape,
                                            SelectedDialect,
                                            formatChoice,
                                            IsAddSuffix
                                        );
                                    }
                                    shape.TextFrame.TextRange.Text = kurdishDate;
                                }
                            }
                            break;

                        case PpSelectionType.ppSelectionText:
                            var textRange = PowerPointApp.ActiveWindow.Selection.TextRange;
                            PowerPoint.Shape parentShape = textRange.Parent as PowerPoint.Shape;
                            if (parentShape != null && IsAutoUpdate)
                            {
                                // Tag the parent shape of the text range
                                AddCustomTagForPowerPointShape(
                                    parentShape,
                                    SelectedDialect,
                                    formatChoice,
                                    IsAddSuffix
                                );
                            }
                            textRange.Text = kurdishDate;
                            break;

                        default:
                            MessageBox.Show(
                                "Unsupported selection type.",
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                            break;
                    }
                }

                if (ProjectApp != null)
                {
                    var selection = ProjectApp.ActiveSelection;
                    var activeCell = ProjectApp.ActiveCell;
                    string activeFieldName = activeCell.FieldName; // This is the field related to the user's current selection/interaction.

                    if (selection.Tasks == null || selection.Tasks.Count == 0)
                    {
                        MessageBox.Show(
                            "Please select one or more tasks.",
                            "No Task Selected",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                        return;
                    }

                    // Update only the field corresponding to the active cell across all selected tasks.
                    foreach (MSProject.Task task in selection.Tasks)
                    {
                        if (task != null)
                        {
                            SetTaskFieldValue(task, activeFieldName, kurdishDate);
                        }
                    }
                }

                if (VisioApp != null)
                {
                    // Get the active Visio window.
                    Visio.Window activeWindow = VisioApp.ActiveWindow;

                    // Check if there are selected shapes.
                    if (activeWindow.Selection.Count > 0)
                    {
                        foreach (Visio.Shape shape in activeWindow.Selection)
                        {
                            // Insert text into the selected shape(s) with the determined formatChoice, dialect, and isAddSuffix
                            shape.Text = kurdishDate;
                        }
                    }
                    else
                    {
                        MessageBox.Show(
                            "No shapes selected.",
                            "Information",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    "Failed to insert date: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private Document GetWordEditorItem(object item)
        {
            switch (item)
            {
                case MailItem mailItem:
                    return mailItem.GetInspector.WordEditor;
                case AppointmentItem appointmentItem:
                    return appointmentItem.GetInspector.WordEditor;
                case TaskItem taskItem:
                    return taskItem.GetInspector.WordEditor;
                case ContactItem contactItem:
                    return contactItem.GetInspector.WordEditor;
                default:
                    return null;
            }
        }

        private void ProcessItemInsertOutlook(object item, string kurdishDate, int formatChoice)
        {
            Document wordEditor = GetWordEditorItem(item);

            if (wordEditor == null || wordEditor.Application.Selection == null)
            {
                MessageBox.Show(
                    "This item does not support Word editing.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            if (IsAutoUpdate)
            {
                Range currentRange = wordEditor.Application.Selection.Range;
                currentRange.Text = kurdishDate;
                currentRange.SetRange(currentRange.Start, currentRange.Start + kurdishDate.Length);

                string bookmarkName = "KDate" + Guid.NewGuid().ToString().Replace("-", "");
                if (!wordEditor.Bookmarks.Exists(bookmarkName))
                {
                    wordEditor.Bookmarks.Add(bookmarkName, currentRange);
                    AddCustomXmlPartForWordOutlook(
                        bookmarkName,
                        SelectedDialect,
                        formatChoice,
                        IsAddSuffix,
                        wordEditor
                    );
                }
                currentRange.Select();
            }
            else
            {
                wordEditor.Application.Selection.TypeText(kurdishDate);
                wordEditor.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
            }
        }

        private void AddCustomTagForPowerPointShape(
            PowerPoint.Shape shape,
            string dialect,
            int formatChoice,
            bool isAddSuffix
        )
        {
            // Set custom tags on the shape to store the date format and other relevant data
            shape.Tags.Add("kDateShape", "true");
            shape.Tags.Add("Dialect", dialect);
            shape.Tags.Add("FormatChoice", formatChoice.ToString());
            shape.Tags.Add("IsAddSuffix", isAddSuffix.ToString());
        }

        private void AddCustomXmlPartForWordOutlook(
            string bookmarkName,
            string dialect,
            int formatChoice,
            bool isAddSuffix,
            Document WordEditor
        )
        {
            string customXml =
                $@"<KurdishDateInsertion>
                     <DateInfo>
                        <Dialect>{dialect}</Dialect>
                        <FormatChoice>{formatChoice}</FormatChoice>
                        <IsAddSuffix>{isAddSuffix}</IsAddSuffix>
                        <BookmarkName>{bookmarkName}</BookmarkName>
                     </DateInfo>
                   </KurdishDateInsertion>";
            WordEditor.CustomXMLParts.Add(customXml);
        }

        private void AddCustomXmlPartForExcel(
            Workbook workbook,
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

        public void UpdateDatesFromCustomXmlPartsForWordOutlook(Document wordEditor)
        {
            List<Bookmark> bookmarksToUpdate = new List<Bookmark>();

            // Collect all relevant bookmarks to update
            foreach (Bookmark bookmark in wordEditor.Bookmarks)
            {
                if (bookmark.Name.StartsWith("KDate"))
                {
                    bookmarksToUpdate.Add(bookmark);
                }
            }

            // Process each bookmark to update its content
            foreach (Bookmark bookmark in bookmarksToUpdate)
            {
                Range bookmarkRange = bookmark.Range;
                string bookmarkName = bookmark.Name;
                CustomXMLPart part = FindCustomXmlPartForBookmarkForWordOutlook(
                    bookmarkName,
                    wordEditor
                );

                if (part != null)
                {
                    var dialect = part.SelectSingleNode(
                        "/KurdishDateInsertion/DateInfo/Dialect"
                    )?.Text;
                    var formatChoice = int.Parse(
                        part.SelectSingleNode("/KurdishDateInsertion/DateInfo/FormatChoice")?.Text
                            ?? "0"
                    );
                    var isAddSuffix = bool.Parse(
                        part.SelectSingleNode("/KurdishDateInsertion/DateInfo/IsAddSuffix")?.Text
                            ?? "false"
                    );

                    string newDate = new DateInsertion().Kurdish(
                        formatChoice,
                        dialect,
                        isAddSuffix
                    );

                    // Replace old content and reset the bookmark
                    bookmarkRange.Text = newDate;

                    // Create a new range for the new text
                    int newStart = bookmarkRange.Start;
                    int newEnd = newStart + newDate.Length;

                    // update bookmark range
                    bookmarkRange.SetRange(newStart, newEnd);
                    wordEditor.Bookmarks.Add(bookmarkName, bookmarkRange);
                }
                else
                {
                    Debug.WriteLine($"No matching XML part found for bookmark: {bookmarkName}");
                }
            }
        }

        public void UpdateDatesFromCustomXmlPartsForExcel(Workbook workbook)
        {
            // Assuming a specific named range or mechanism to identify cells with dates to update
            foreach (Excel.Name name in workbook.Names)
            {
                if (name.Name.Contains("KDate"))
                {
                    Excel.Range range = workbook.Application.Range[name.RefersTo];
                    CustomXMLPart part = FindCustomXmlPartForRangeForExcel(name.Name, workbook);

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

                        string newDate = new DateInsertion().Kurdish(
                            formatChoice,
                            dialect,
                            isAddSuffix
                        );
                        range.Value = newDate;
                    }
                    else
                    {
                        Debug.WriteLine($"No matching XML part found for range named: {name.Name}");
                    }
                }
            }
        }

        public void UpdateDatesFromCustomXmlPartsForPowerPoint(Presentation presentation)
        {
            try
            {
                foreach (Slide slide in presentation.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (
                            shape.HasTextFrame == MsoTriState.msoTrue
                            && shape.TextFrame.HasText == MsoTriState.msoTrue
                        )
                        {
                            // Check if this shape has a tag indicating it contains a date needing updates
                            if (shape.Tags["kDateShape"] == "true")
                            {
                                string dialect = shape.Tags["Dialect"];
                                int formatChoice = int.Parse(shape.Tags["FormatChoice"]);
                                bool isAddSuffix = bool.Parse(shape.Tags["IsAddSuffix"]);

                                string newDate = new DateInsertion().Kurdish(
                                    formatChoice,
                                    dialect,
                                    isAddSuffix
                                );

                                // Update the text in the shape
                                shape.TextFrame.TextRange.Text = newDate;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    "Failed to update dates: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private CustomXMLPart FindCustomXmlPartForBookmarkForWordOutlook(
            string bookmarkName,
            Document doc
        )
        {
            foreach (CustomXMLPart part in doc.CustomXMLParts)
            {
                if (part.NamespaceURI == string.Empty)
                {
                    var nameNode = part.SelectSingleNode(
                        "/KurdishDateInsertion/DateInfo/BookmarkName"
                    );
                    if (nameNode != null && nameNode.Text == bookmarkName)
                    {
                        return part;
                    }
                }
            }
            return null;
        }

        private CustomXMLPart FindCustomXmlPartForRangeForExcel(string nameRef, Workbook workbook)
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
