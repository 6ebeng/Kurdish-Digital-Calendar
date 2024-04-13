using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KDCLibrary;
using KDCLibrary.Calendars;
using KDCLibrary.Helpers;
using Microsoft.Office.Core;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Project = Microsoft.Office.Interop.MSProject;
using Visio = Microsoft.Office.Interop.Visio;
using Word = Microsoft.Office.Interop.Word;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace KDCLibrary
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private const string IsReverseKeyName = KDCConstants.KeyNames.IsReverse;
        private const string SelectedDialectKeyName = KDCConstants.KeyNames.SelectedDialect;
        private const string SelectedFormat1KeyName = KDCConstants.KeyNames.SelectedFormat1;
        private const string SelectedFormat2KeyName = KDCConstants.KeyNames.SelectedFormat2;
        private const string LastSelectionGroup1KeyName = KDCConstants.KeyNames.LastSelectionGroup1;
        private const string LastSelectionGroup2KeyName = KDCConstants.KeyNames.LastSelectionGroup2;
        private const string CheckBoxStatesKeyName = KDCConstants.KeyNames.CheckBoxStates;
        private const string isAddSuffixKeyName = KDCConstants.KeyNames.IsAddSuffix;

        private readonly List<string> _dialectsList = KDCConstants.DefaultValues.Dialects;
        private readonly List<string> _formatsList = KDCConstants.DefaultValues.Formats;
        private readonly List<string> _calendarGroup1List = KDCConstants
            .DefaultValues
            .CalendarGroup1;
        private readonly List<string> _calendarGroup2List = KDCConstants
            .DefaultValues
            .CalendarGroup2;

        private string _selectedDialect { get; set; }
        private string _selectedCalendar { get; set; }
        private string _selectedFromFormat { get; set; }
        private string _selectedToFormat { get; set; }
        private bool _isReverse { get; set; }
        private bool _isAddSuffix { get; set; }
        private string _selectedInsertFormat { get; set; }

        private readonly Outlook.Application _outlookApp = null;
        private readonly Word.Application _wordApp = null;
        private readonly Excel.Application _excelApp = null;
        private readonly PowerPoint.Application _powerPointApp = null;
        private readonly Visio.Application _visioApp = null;
        private readonly Project.Application _projectApp = null;
        private readonly IRibbonControl _control = null;

        public Ribbon(
            Outlook.Application outlookApp,
            Word.Application wordApp,
            Excel.Application excelApp,
            PowerPoint.Application powerPointApp,
            Visio.Application visioApp,
            Project.Application projectApp
        )
        {
            _outlookApp = outlookApp;
            _wordApp = wordApp;
            _excelApp = excelApp;
            _powerPointApp = powerPointApp;
            _visioApp = visioApp;
            _projectApp = projectApp;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return new Helper().GetResourceText("KDCLibrary.UI.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

            // Restore the _isReverse state from the registry when the ribbon loads
            this._isReverse = Convert.ToBoolean(
                RegistryHelper.LoadSetting(IsReverseKeyName, "false")
            );
            string savedCalendarKeyName = _isReverse
                ? LastSelectionGroup2KeyName
                : LastSelectionGroup1KeyName;
            this._selectedCalendar = RegistryHelper.LoadSetting(
                savedCalendarKeyName,
                _isReverse ? _calendarGroup2List[0] : _calendarGroup1List[0]
            );
            this._selectedDialect = RegistryHelper.LoadSetting(
                SelectedDialectKeyName,
                _dialectsList[0]
            );
            this._selectedFromFormat = RegistryHelper.LoadSetting(
                SelectedFormat1KeyName,
                _formatsList[0]
            );
            this._selectedToFormat = RegistryHelper.LoadSetting(
                SelectedFormat2KeyName,
                _formatsList[0]
            );
            this._isAddSuffix = Convert.ToBoolean(
                RegistryHelper.LoadSetting(isAddSuffixKeyName, "false")
            );
            this._selectedInsertFormat = getSelectedCheckBox();

            // Invalidate the controls that depend on the _isReverse state, if necessary
            // This ensures that UI elements reflect the correct state from the beginning
            ribbon.InvalidateControl("toggleButton1");
            ribbon.InvalidateControl("dropDown2");
        }

        public bool getCheckBoxPressed(Office.IRibbonControl control)
        {
            // Load the saved states
            var states = LoadCheckBoxStates();

            // Return the state for the specified control
            return states.TryGetValue(control.Id, out bool isPressed) && isPressed;
        }

        private void SaveCheckBoxStates(Dictionary<string, bool> states)
        {
            RegistryHelper.SaveSetting(CheckBoxStatesKeyName, JsonConvert.SerializeObject(states));
        }

        private Dictionary<string, bool> LoadCheckBoxStates()
        {
            var serializedState = RegistryHelper.LoadSetting(CheckBoxStatesKeyName, "{}");
            Debug.WriteLine($"Loaded serialized state: {serializedState}");

            try
            {
                var states = JsonConvert.DeserializeObject<Dictionary<string, bool>>(
                    serializedState
                );
                // If the dictionary is empty, default to first checkbox checked
                if (states == null || !states.Any())
                {
                    return new Dictionary<string, bool> { { "checkBox1", true } };
                }
                return states;
            }
            catch (Newtonsoft.Json.JsonReaderException ex)
            {
                Debug.WriteLine($"JSON parsing error: {ex.Message}");
                // If parsing fails, default to first checkbox checked
                return new Dictionary<string, bool> { { "checkBox1", true } };
            }
        }

        public void onToggleButtonAction(Office.IRibbonControl control, bool isPressed)
        {
            if (control.Id == "toggleButton1")
            {
                RegistryHelper.SaveSetting(IsReverseKeyName, isPressed.ToString());
                // Invalidate dropdown2 to refresh its items based on the new IsReverse state
                ribbon.InvalidateControl("dropDown2");
            }
        }

        public void onDropDownAction(
            Office.IRibbonControl control,
            string selectedId,
            int selectedIndex
        )
        {
            Debug.WriteLine("Selected DropDown Index: " + selectedIndex);

            switch (control.Id)
            {
                case "dropDown1":
                    this._selectedDialect = _dialectsList[selectedIndex];
                    RegistryHelper.SaveSetting(
                        SelectedDialectKeyName,
                        _dialectsList[selectedIndex]
                    );
                    break;

                case "dropDown2":
                    this._selectedCalendar = _isReverse
                        ? _calendarGroup2List[selectedIndex]
                        : _calendarGroup1List[selectedIndex];
                    var keyName = _isReverse
                        ? LastSelectionGroup2KeyName
                        : LastSelectionGroup1KeyName;
                    RegistryHelper.SaveSetting(keyName, _selectedCalendar);
                    break;

                case "dropDown3":
                    this._selectedFromFormat = _formatsList[selectedIndex];
                    RegistryHelper.SaveSetting(SelectedFormat1KeyName, _selectedFromFormat);
                    break;

                case "dropDown4":
                    this._selectedToFormat = _formatsList[selectedIndex];
                    RegistryHelper.SaveSetting(SelectedFormat2KeyName, _selectedToFormat);
                    break;
            }
        }

        public void onCheckBoxAction(Office.IRibbonControl control, bool isPressed)
        {
            // Load the current states of checkboxes
            var states = LoadCheckBoxStates();

            // Count how many checkboxes are currently checked
            int checkedCount = states.Count(kvp => kvp.Value);

            if (isPressed)
            {
                // If the checkbox is being checked, ensure all others are unchecked
                foreach (var key in states.Keys.ToList())
                {
                    states[key] = false;
                }
                // Check the current checkbox
                states[control.Id] = true;
            }
            else if (!isPressed && checkedCount <= 1)
            {
                // If the checkbox is being unchecked but it's the only one checked, prevent this action
                // Essentially, do nothing to keep the current checkbox checked
                // This block can be empty or display a message if desired
            }
            else
            {
                // If unchecking and other checkboxes are checked, allow unchecking
                states[control.Id] = false;
            }

            // Save the updated states
            SaveCheckBoxStates(states);

            this._selectedInsertFormat = getSelectedCheckBox();
            // Invalidate all checkboxes to update their states in the UI
            ribbon.Invalidate(); // This will refresh the whole ribbon, alternatively, you could invalidate each control individually
            populateInsertDate();
        }

        public string getSelectedCheckBox()
        {
            // Load the current states of checkboxes
            var checkBoxStates = LoadCheckBoxStates();

            if (checkBoxStates.Count == 0 || !checkBoxStates.Values.Any(v => v))
            {
                // If no checkboxes are checked or if the states dictionary is empty, default to the first checkbox being selected
                // Optionally, you could ensure that the first checkbox's state is set to true in checkBoxStates here
                return _formatsList[0];
            }

            foreach (var checkBoxState in checkBoxStates)
            {
                if (checkBoxState.Value) // If the checkbox is selected
                {
                    // Extract the numeric part of the checkBox ID and use it to index into _formats
                    if (
                        int.TryParse(checkBoxState.Key.Replace("checkBox", ""), out int index)
                        && index <= _formatsList.Count
                    )
                    {
                        // Adjust for zero-based indexing if necessary
                        int adjustedIndex = index - 1; // Assuming checkBox1 corresponds to _formats[0]

                        if (adjustedIndex >= 0 && adjustedIndex < _formatsList.Count)
                        {
                            return _formatsList[adjustedIndex];
                        }
                    }
                }
            }

            // Return null or an empty string if no checkbox is selected
            // This line should never be reached with the new logic added above, but is kept for safety
            return null;
        }

        public bool restoreisAddSuffixState(Office.IRibbonControl control)
        {
            this._isAddSuffix = Convert.ToBoolean(
                RegistryHelper.LoadSetting(isAddSuffixKeyName, "false")
            );
            return _isAddSuffix;
        }

        public bool restoreIsReverseState(Office.IRibbonControl control)
        {
            this._isReverse = Convert.ToBoolean(
                RegistryHelper.LoadSetting(IsReverseKeyName, "false")
            );
            return _isReverse;
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "splitButton1__btn":
                    return KDCLibrary.Properties.Resources.calendar;
                case "splitButton3__btn":
                    return KDCLibrary.Properties.Resources.help;
                case "splitButton2__btn":
                    return KDCLibrary.Properties.Resources.convert;
                default:
                    return null;
            }
        }

        public string GetCheckBoxLabelById(string CheckBoxId)
        {
            Debug.WriteLine(CheckBoxId);
            switch (CheckBoxId)
            {
                case "checkBox1":
                    return _formatsList[0];
                case "checkBox2":
                    return _formatsList[1];
                case "checkBox3":
                    return _formatsList[2];
                case "checkBox4":
                    return _formatsList[3];
                case "checkBox15":
                    return _formatsList[4];
                case "checkBox6":
                    return _formatsList[5];
                default:
                    return "Unknown";
            }
        }

        private void populateInsertDate()
        {
            if (_selectedInsertFormat == null)
            {
                MessageBox.Show(
                    "No format selected.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            // Determine the format choice based on the label of the checked checkbox.
            int formatChoice = DetermineFormatChoiceFromCheckbox(_selectedInsertFormat);

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

            if (_outlookApp != null)
            {
                var inspector = _outlookApp.Application.ActiveInspector();
                if (
                    inspector != null
                    && inspector.CurrentItem is Microsoft.Office.Interop.Outlook.MailItem mailItem
                )
                {
                    var wordEditor = mailItem.GetInspector.WordEditor as Word.Document;
                    if (wordEditor != null && wordEditor.Application.Selection != null)
                    {
                        // Insert Kurdish Date with the determined formatChoice, dialect, and isAddSuffix
                        wordEditor.Application.Selection.TypeText(
                            new DateInsertion().Kurdish(
                                formatChoice,
                                _selectedDialect,
                                _isAddSuffix
                            )
                        );
                    }
                }
                else
                {
                    MessageBox.Show(
                        "No checkbox selected.",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }

            if (_wordApp != null)
            {
                _wordApp.Application.Selection.TypeText(
                    new DateInsertion().Kurdish(formatChoice, _selectedDialect, _isAddSuffix)
                );
            }

            if (_powerPointApp != null)
            {
                // Get the active application

                if (_powerPointApp.ActiveWindow.Selection == null)
                {
                    MessageBox.Show(
                        "No selection detected.",
                        "Info",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                // Handle different selection types
                switch (_powerPointApp.ActiveWindow.Selection.Type)
                {
                    case PowerPoint.PpSelectionType.ppSelectionShapes:
                        // For shape selection, insert at the end of the text in the first shape
                        foreach (
                            PowerPoint.Shape shape in _powerPointApp
                                .ActiveWindow
                                .Selection
                                .ShapeRange
                        )
                        {
                            if (
                                shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                                && shape.TextFrame.HasText
                                    == Microsoft.Office.Core.MsoTriState.msoTrue
                            )
                            {
                                shape.TextFrame.TextRange.InsertAfter(
                                    new DateInsertion().Kurdish(
                                        formatChoice,
                                        _selectedDialect,
                                        _isAddSuffix
                                    )
                                );
                            }
                            else if (
                                shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                            )
                            {
                                shape.TextFrame.TextRange.Text = new DateInsertion().Kurdish(
                                    formatChoice,
                                    _selectedDialect,
                                    _isAddSuffix
                                );
                            }
                        }
                        break;

                    case PowerPoint.PpSelectionType.ppSelectionText:
                        // For text selection, replace the selected text
                        var selectedTextRange = _powerPointApp.ActiveWindow.Selection.TextRange;
                        selectedTextRange.Text = new DateInsertion().Kurdish(
                            formatChoice,
                            _selectedDialect,
                            _isAddSuffix
                        );
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

            if (_excelApp != null)
            {
                // Get the selected range
                Excel.Range selectedRange = _excelApp.Application.Selection;
                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    // Insert Kurdish Date with the determined formatChoice, dialect, and isAddSuffix
                    cell.Value = new DateInsertion().Kurdish(
                        formatChoice,
                        _selectedDialect,
                        _isAddSuffix
                    );
                }
            }
        }

        private int DetermineFormatChoiceFromCheckbox(string checkboxLabel)
        {
            switch (checkboxLabel)
            {
                case "dd/MM/yyyy":
                    return 4;
                case "MM/dd/yyyy":
                    return 10;
                case "yyyy/MM/dd":
                    return 11;
                case "dddd, dd MMMM, yyyy":
                    return 1;
                case "dddd, dd/MM/yyyy":
                    return 2;
                case "dd MMMM, yyyy":
                    return 3;
                default:
                    MessageBox.Show(
                        "Unsupported target format selected.",
                        "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation
                    );
                    return -1; // Indicates an unsupported format
            }
        }

        #endregion

        #region Callbacks for Ribbon Controls



        public void checkBox7_Click(Office.IRibbonControl control, bool isPressed)
        {
            this._isAddSuffix = isPressed;
            RegistryHelper.SaveSetting(isAddSuffixKeyName, isPressed.ToString());
        }

        public void toggleButton1_Click(Office.IRibbonControl control, bool isPressed)
        {
            this._isReverse = isPressed;
            RegistryHelper.SaveSetting(IsReverseKeyName, _isReverse.ToString());
            this.ribbon.InvalidateControl("dropDown2");
        }

        public void splitButton1_Click(Office.IRibbonControl control)
        {
            populateInsertDate();
        }

        public void splitButton2_Click(Office.IRibbonControl control)
        {
            if (_outlookApp != null)
            {
                var inspector = _outlookApp.Application.ActiveInspector();
                if (
                    inspector != null
                    && inspector.CurrentItem is Microsoft.Office.Interop.Outlook.MailItem mailItem
                )
                {
                    var wordEditor = mailItem.GetInspector.WordEditor as Word.Document;
                    if (wordEditor != null && wordEditor.Application.Selection != null)
                    {
                        // Assuming ConvertDateBasedOnUserSelection returns the converted date as a string
                        string convertedDate = new DateConversion().ConvertDateBasedOnUserSelection(
                            wordEditor.Application.Selection.Text,
                            _isReverse,
                            RegistryHelper.LoadSetting(SelectedDialectKeyName, ""),
                            RegistryHelper.LoadSetting(SelectedFormat1KeyName, ""),
                            RegistryHelper.LoadSetting(SelectedFormat2KeyName, ""),
                            _selectedCalendar,
                            _isAddSuffix
                        );

                        wordEditor.Application.Selection.Text = convertedDate;
                    }
                }
            }

            if (_powerPointApp != null)
            {
                if (_powerPointApp.ActiveWindow.Selection != null)
                {
                    var selection = _powerPointApp.ActiveWindow.Selection;
                    switch (selection.Type)
                    {
                        case PowerPoint.PpSelectionType.ppSelectionShapes:
                            // Handle shape selections
                            foreach (PowerPoint.Shape shape in selection.ShapeRange)
                            {
                                if (
                                    shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                                    && shape.TextFrame.HasText
                                        == Microsoft.Office.Core.MsoTriState.msoTrue
                                )
                                {
                                    Debug.WriteLine(
                                        "PowerPoint.PpSelectionType.ppSelectionShapes",
                                        shape.TextFrame.TextRange.Text
                                    );
                                    // Apply conversion on the shape's text
                                    shape.TextFrame.TextRange.Text =
                                        new DateConversion().ConvertDateBasedOnUserSelection(
                                            shape.TextFrame.TextRange.Text,
                                            _isReverse,
                                            _selectedDialect,
                                            _selectedFromFormat,
                                            _selectedToFormat,
                                            _selectedCalendar,
                                            _isAddSuffix
                                        );
                                }
                            }
                            break;

                        case PowerPoint.PpSelectionType.ppSelectionText:
                            // Directly handle text selections without iterating over ShapeRange
                            var textRange = selection.TextRange;
                            if (textRange != null && textRange.Length > 0)
                            {
                                Debug.WriteLine(
                                    "PowerPoint.PpSelectionType.ppSelectionText",
                                    textRange.Text
                                );
                                // Apply conversion on the selected text
                                textRange.Text =
                                    new DateConversion().ConvertDateBasedOnUserSelection(
                                        textRange.Text,
                                        _isReverse,
                                        _selectedDialect,
                                        _selectedFromFormat,
                                        _selectedToFormat,
                                        _selectedCalendar,
                                        _isAddSuffix
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

            if (_wordApp != null)
            {
                _wordApp.Application.Selection.Text =
                    new DateConversion().ConvertDateBasedOnUserSelection(
                        _wordApp.Application.Selection.Text,
                        _isReverse,
                        _selectedDialect,
                        _selectedFromFormat,
                        _selectedToFormat,
                        _selectedCalendar,
                        _isAddSuffix
                    );
            }

            if (_excelApp != null)
            {
                // Get the selected range
                Excel.Range selectedRange = _excelApp.Application.Selection;
                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    Object selectedText = cell.Value;

                    // if the object is DateTime then make toString("d") to get the short date pattern
                    if (selectedText is DateTime)
                    {
                        selectedText = ((DateTime)selectedText).ToString("d");
                    }
                    string dateString = selectedText.ToString(); // 'd' format string returns the short date pattern

                    string result = new DateConversion().ConvertDateBasedOnUserSelection(
                        dateString,
                        _isReverse,
                        _selectedDialect,
                        _selectedFromFormat,
                        _selectedToFormat,
                        _selectedCalendar,
                        _isAddSuffix
                    );

                    Debug.WriteLine(result);
                    // if result is empty break the loop
                    if (result == dateString || string.IsNullOrEmpty(result))
                    {
                        //if(selectedRange.Cells.Count > 1) MessageBox.Show("Some cells were not converted.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        // Highlight the cell with red color
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(
                            System.Drawing.Color.Red
                        );
                    }
                    else
                    {
                        // delete highlight color set it to no Fill and set to no border
                        cell.Interior.ColorIndex = Excel.Constants.xlNone;
                        // set no border
                        cell.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                        // set the value of the cell to the result
                        cell.Value = result;
                    }
                }
            }
        }

        #endregion


        #region Load current List of Dialects, Formats, and Calendar Groups

        public string getDialectLabel(Office.IRibbonControl control, int index)
        {
            // Return the label of the dialect at the specified index
            return _dialectsList[index];
        }

        public int getDialectCount(Office.IRibbonControl control)
        {
            // Return the number of dialects available
            return _dialectsList.Count;
        }

        public int getSelectedDialectIndex(Office.IRibbonControl control)
        {
            // Load the saved dialect name from your settings
            this._selectedDialect = RegistryHelper.LoadSetting(
                SelectedDialectKeyName,
                _dialectsList[0]
            );

            int index = _dialectsList.IndexOf(_selectedDialect);
            // Return the index of the saved dialect or default to the first item if not found
            return index >= 0 ? index : 0;
        }

        public string getSelectedDialectLabel(Office.IRibbonControl control)
        {
            this._selectedDialect = _dialectsList[getSelectedDialectIndex(control)];
            return _selectedDialect;
        }

        public string getCalendarLabel(Office.IRibbonControl control, int index)
        {
            // Check the _isReverse state to decide which list to use
            List<string> selectedList = _isReverse ? _calendarGroup2List : _calendarGroup1List;
            // Return the label of the calendar at the specified index from the appropriate list
            return selectedList[index];
        }

        public int getCalendarCount(Office.IRibbonControl control)
        {
            // Return the number of calendars available based on the _isReverse state
            return _isReverse ? _calendarGroup2List.Count : _calendarGroup1List.Count;
        }

        public int getSelectedCalendarIndex(Office.IRibbonControl control)
        {
            // Determine the correct key name based on the _isReverse state
            string savedCalendarKeyName = _isReverse
                ? LastSelectionGroup2KeyName
                : LastSelectionGroup1KeyName;
            this._selectedCalendar = RegistryHelper.LoadSetting(
                savedCalendarKeyName,
                _isReverse ? _calendarGroup2List[0] : _calendarGroup1List[0]
            );
            List<string> selectedList = _isReverse ? _calendarGroup2List : _calendarGroup1List;
            int index = selectedList.IndexOf(_selectedCalendar);
            // Default to the first item if the saved calendar is not found
            return index >= 0 ? index : 0;
        }

        public string getSelectedCalendarLabel(Office.IRibbonControl control)
        {
            this._selectedCalendar = _isReverse
                ? _calendarGroup2List[getSelectedCalendarIndex(control)]
                : _calendarGroup1List[getSelectedCalendarIndex(control)];
            return _selectedCalendar;
        }

        public string getFromFormatLabel(Office.IRibbonControl control, int index)
        {
            // Return the label of the format at the specified index
            return _formatsList[index];
        }

        public int getFromFormatCount(Office.IRibbonControl control)
        {
            // Return the number of formats available
            return _formatsList.Count;
        }

        public int getSelectedFromFormatIndex(Office.IRibbonControl control)
        {
            // Load the saved format name from your settings
            this._selectedFromFormat = RegistryHelper.LoadSetting(
                SelectedFormat1KeyName,
                _formatsList[0]
            );
            int index = _formatsList.IndexOf(_selectedFromFormat);
            // Return the index of the saved format or default to the first item if not found
            return index >= 0 ? index : 0; // Ensure we return an integer
        }

        public string getSelectedFromFormatLabel(Office.IRibbonControl control)
        {
            this._selectedFromFormat = _formatsList[getSelectedFromFormatIndex(control)];
            return _selectedFromFormat;
        }

        public string getToFormatLabel(Office.IRibbonControl control, int index)
        {
            // Return the label of the format at the specified index
            return _formatsList[index];
        }

        public int getToFormatCount(Office.IRibbonControl control)
        {
            // Return the number of formats available
            return _formatsList.Count;
        }

        public int getSelectedToFormatIndex(Office.IRibbonControl control)
        {
            // Load the saved format name from your settings
            _selectedToFormat = RegistryHelper.LoadSetting(SelectedFormat2KeyName, _formatsList[0]);
            int index = _formatsList.IndexOf(_selectedToFormat);
            return index >= 0 ? index : 0; // Ensure we return an integer
        }

        public string getSelectedToFormatLabel(Office.IRibbonControl control)
        {
            this._selectedToFormat = _formatsList[getSelectedToFormatIndex(control)];
            return _selectedToFormat;
        }

        public void button1_Click(Office.IRibbonControl control)
        {
            new CreditsForm();
        }

        #endregion
    }
}
