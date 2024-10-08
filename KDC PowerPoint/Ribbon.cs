﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KDCLibrary;
using Microsoft.Office.Core;
using Newtonsoft.Json;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace KDC_PowerPoint
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        #region Intializers

        IKDCService kDCService = new KDCServiceImplementation();
        private Office.IRibbonUI ribbon;
        private const string IsReverseKeyName = Constants.KeyNames.IsReverse;
        private const string SelectedDialectKeyName = Constants.KeyNames.SelectedDialect;
        private const string SelectedFormat1KeyName = Constants.KeyNames.SelectedFormat1;
        private const string SelectedFormat2KeyName = Constants.KeyNames.SelectedFormat2;
        private const string LastSelectionGroup1KeyName = Constants.KeyNames.LastSelectionGroup1;
        private const string LastSelectionGroup2KeyName = Constants.KeyNames.LastSelectionGroup2;
        private const string CheckBoxStatesKeyName = Constants.KeyNames.CheckBoxStates;
        private const string isAddSuffixKeyName = Constants.KeyNames.IsAddSuffix;

        private readonly List<string> _dialectsList = Constants.DefaultValues.Dialects;
        private readonly List<string> _formatsList = Constants.DefaultValues.Formats;
        private readonly List<string> _calendarGroup1List = Constants.DefaultValues.CalendarGroup1;
        private readonly List<string> _calendarGroup2List = Constants.DefaultValues.CalendarGroup2;
        private readonly string AppName = Assembly.GetExecutingAssembly().GetName().Name;

        private string _selectedDialect { get; set; }
        private string _selectedCalendar { get; set; }
        private string _selectedFromFormat { get; set; }
        private string _selectedToFormat { get; set; }
        private bool _isReverse { get; set; }
        private bool _isAddSuffix { get; set; }
        private string _selectedInsertFormat { get; set; }

        #endregion

        public Ribbon() { }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return kDCService.GetRibbonXml();
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

            // Restore the _isReverse state from the registry when the ribbon loads
            this._isReverse = Convert.ToBoolean(
                kDCService.LoadSetting(IsReverseKeyName, "false", AppName)
            );
            string savedCalendarKeyName = _isReverse
                ? LastSelectionGroup2KeyName
                : LastSelectionGroup1KeyName;
            this._selectedCalendar = kDCService.LoadSetting(
                savedCalendarKeyName,
                _isReverse ? _calendarGroup2List[0] : _calendarGroup1List[0],
                AppName
            );
            this._selectedDialect = kDCService.LoadSetting(
                SelectedDialectKeyName,
                _dialectsList[0],
                AppName
            );
            this._selectedFromFormat = kDCService.LoadSetting(
                SelectedFormat1KeyName,
                _formatsList[0],
                AppName
            );
            this._selectedToFormat = kDCService.LoadSetting(
                SelectedFormat2KeyName,
                _formatsList[0],
                AppName
            );
            this._isAddSuffix = Convert.ToBoolean(
                kDCService.LoadSetting(isAddSuffixKeyName, "false", AppName)
            );
            this._selectedInsertFormat = getSelectedCheckBox();

            // Invalidate the controls that depend on the _isReverse state, if necessary
            // This ensures that UI elements reflect the correct state from the beginning
            ribbon.InvalidateControl("toggleButton1");
            ribbon.InvalidateControl("dropDown2");
        }

        public bool getCheckBoxPressed(IRibbonControl control)
        {
            // Load the saved states
            var states = LoadCheckBoxStates();

            // Return the state for the specified control
            return states.TryGetValue(control.Id, out bool isPressed) && isPressed;
        }

        private void SaveCheckBoxStates(Dictionary<string, bool> states)
        {
            kDCService.SaveSetting(
                CheckBoxStatesKeyName,
                JsonConvert.SerializeObject(states),
                AppName
            );
        }

        private Dictionary<string, bool> LoadCheckBoxStates()
        {
            var serializedState = kDCService.LoadSetting(
                CheckBoxStatesKeyName,
                "{ 'checkBox1', true }",
                AppName
            );
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

        public void onToggleButtonAction(IRibbonControl control, bool isPressed)
        {
            if (control.Id == "toggleButton1")
            {
                kDCService.SaveSetting(IsReverseKeyName, isPressed.ToString(), AppName);
                // Invalidate dropdown2 to refresh its items based on the new IsReverse state
                ribbon.InvalidateControl("dropDown2");
            }
        }

        public void onDropDownAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
            Debug.WriteLine("Selected DropDown Index: " + selectedIndex);

            switch (control.Id)
            {
                case "dropDown1":
                    this._selectedDialect = _dialectsList[selectedIndex];
                    kDCService.SaveSetting(
                        SelectedDialectKeyName,
                        _dialectsList[selectedIndex],
                        AppName
                    );
                    break;

                case "dropDown2":
                    this._selectedCalendar = _isReverse
                        ? _calendarGroup2List[selectedIndex]
                        : _calendarGroup1List[selectedIndex];
                    var keyName = _isReverse
                        ? LastSelectionGroup2KeyName
                        : LastSelectionGroup1KeyName;
                    kDCService.SaveSetting(keyName, _selectedCalendar, AppName);
                    break;

                case "dropDown3":
                    this._selectedFromFormat = _formatsList[selectedIndex];
                    kDCService.SaveSetting(SelectedFormat1KeyName, _selectedFromFormat, AppName);
                    break;

                case "dropDown4":
                    this._selectedToFormat = _formatsList[selectedIndex];
                    kDCService.SaveSetting(SelectedFormat2KeyName, _selectedToFormat, AppName);
                    break;
            }
        }

        public void onCheckBoxAction(IRibbonControl control, bool isPressed)
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

        public bool restoreisAddSuffixState(IRibbonControl control)
        {
            this._isAddSuffix = Convert.ToBoolean(
                kDCService.LoadSetting(isAddSuffixKeyName, "false", AppName)
            );
            return _isAddSuffix;
        }

        public bool restoreIsReverseState(IRibbonControl control)
        {
            this._isReverse = Convert.ToBoolean(
                kDCService.LoadSetting(IsReverseKeyName, "false", AppName)
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
            int formatChoice = kDCService.DetermineFormatChoiceFromCheckbox(_selectedInsertFormat);

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

            // Get the active application

            if (Globals.ThisAddIn.Application.ActiveWindow.Selection == null)
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
            switch (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type)
            {
                case PowerPoint.PpSelectionType.ppSelectionShapes:
                    // For shape selection, insert at the end of the text in the first shape
                    foreach (
                        PowerPoint.Shape shape in Globals
                            .ThisAddIn
                            .Application
                            .ActiveWindow
                            .Selection
                            .ShapeRange
                    )
                    {
                        if (
                            shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                            && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue
                        )
                        {
                            shape.TextFrame.TextRange.InsertAfter(
                                kDCService.Kurdish(formatChoice, _selectedDialect, _isAddSuffix)
                            );
                        }
                        else if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            shape.TextFrame.TextRange.Text = kDCService.Kurdish(
                                formatChoice,
                                _selectedDialect,
                                _isAddSuffix
                            );
                        }
                    }
                    break;

                case PowerPoint.PpSelectionType.ppSelectionText:
                    // For text selection, replace the selected text
                    var selectedTextRange = Globals
                        .ThisAddIn
                        .Application
                        .ActiveWindow
                        .Selection
                        .TextRange;
                    selectedTextRange.Text = kDCService.Kurdish(
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

        #region Callbacks for Ribbon Controls

        public void checkBox7_Click(IRibbonControl control, bool isPressed)
        {
            this._isAddSuffix = isPressed;
            kDCService.SaveSetting(isAddSuffixKeyName, isPressed.ToString(), AppName);
        }

        public void toggleButton1_Click(IRibbonControl control, bool isPressed)
        {
            this._isReverse = isPressed;
            kDCService.SaveSetting(IsReverseKeyName, _isReverse.ToString(), AppName);
            this.ribbon.InvalidateControl("dropDown2");
        }

        public void splitButton1_Click(IRibbonControl control)
        {
            populateInsertDate();
        }

        public void splitButton2_Click(IRibbonControl control)
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection != null)
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
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
                                    kDCService.ConvertDateBasedOnUserSelection(
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
                            textRange.Text = kDCService.ConvertDateBasedOnUserSelection(
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

        #endregion


        #region Load current List of Dialects, Formats, and Calendar Groups

        public string getDialectLabel(IRibbonControl control, int index)
        {
            // Return the label of the dialect at the specified index
            return _dialectsList[index];
        }

        public int getDialectCount(IRibbonControl control)
        {
            // Return the number of dialects available
            return _dialectsList.Count;
        }

        public int getSelectedDialectIndex(IRibbonControl control)
        {
            // Load the saved dialect name from your settings
            this._selectedDialect = kDCService.LoadSetting(
                SelectedDialectKeyName,
                _dialectsList[0],
                AppName
            );

            int index = _dialectsList.IndexOf(_selectedDialect);
            // Return the index of the saved dialect or default to the first item if not found
            return index >= 0 ? index : 0;
        }

        public string getSelectedDialectLabel(IRibbonControl control)
        {
            this._selectedDialect = _dialectsList[getSelectedDialectIndex(control)];
            return _selectedDialect;
        }

        public string getCalendarLabel(IRibbonControl control, int index)
        {
            // Check the _isReverse state to decide which list to use
            List<string> selectedList = _isReverse ? _calendarGroup2List : _calendarGroup1List;
            // Return the label of the calendar at the specified index from the appropriate list
            return selectedList[index];
        }

        public int getCalendarCount(IRibbonControl control)
        {
            // Return the number of calendars available based on the _isReverse state
            return _isReverse ? _calendarGroup2List.Count : _calendarGroup1List.Count;
        }

        public int getSelectedCalendarIndex(IRibbonControl control)
        {
            // Determine the correct key name based on the _isReverse state
            string savedCalendarKeyName = _isReverse
                ? LastSelectionGroup2KeyName
                : LastSelectionGroup1KeyName;
            this._selectedCalendar = kDCService.LoadSetting(
                savedCalendarKeyName,
                _isReverse ? _calendarGroup2List[0] : _calendarGroup1List[0],
                AppName
            );
            List<string> selectedList = _isReverse ? _calendarGroup2List : _calendarGroup1List;
            int index = selectedList.IndexOf(_selectedCalendar);
            // Default to the first item if the saved calendar is not found
            return index >= 0 ? index : 0;
        }

        public string getSelectedCalendarLabel(IRibbonControl control)
        {
            this._selectedCalendar = _isReverse
                ? _calendarGroup2List[getSelectedCalendarIndex(control)]
                : _calendarGroup1List[getSelectedCalendarIndex(control)];
            return _selectedCalendar;
        }

        public string getFromFormatLabel(IRibbonControl control, int index)
        {
            // Return the label of the format at the specified index
            return _formatsList[index];
        }

        public int getFromFormatCount(IRibbonControl control)
        {
            // Return the number of formats available
            return _formatsList.Count;
        }

        public int getSelectedFromFormatIndex(IRibbonControl control)
        {
            // Load the saved format name from your settings
            this._selectedFromFormat = kDCService.LoadSetting(
                SelectedFormat1KeyName,
                _formatsList[0],
                AppName
            );
            int index = _formatsList.IndexOf(_selectedFromFormat);
            // Return the index of the saved format or default to the first item if not found
            return index >= 0 ? index : 0; // Ensure we return an integer
        }

        public string getSelectedFromFormatLabel(IRibbonControl control)
        {
            this._selectedFromFormat = _formatsList[getSelectedFromFormatIndex(control)];
            return _selectedFromFormat;
        }

        public string getToFormatLabel(IRibbonControl control, int index)
        {
            // Return the label of the format at the specified index
            return _formatsList[index];
        }

        public int getToFormatCount(IRibbonControl control)
        {
            // Return the number of formats available
            return _formatsList.Count;
        }

        public int getSelectedToFormatIndex(IRibbonControl control)
        {
            // Load the saved format name from your settings
            _selectedToFormat = kDCService.LoadSetting(
                SelectedFormat2KeyName,
                _formatsList[0],
                AppName
            );
            int index = _formatsList.IndexOf(_selectedToFormat);
            return index >= 0 ? index : 0; // Ensure we return an integer
        }

        public string getSelectedToFormatLabel(IRibbonControl control)
        {
            this._selectedToFormat = _formatsList[getSelectedToFormatIndex(control)];
            return _selectedToFormat;
        }

        public void button1_Click(IRibbonControl control)
        {
            new CreditsForm().Show();
        }

        #endregion


        #endregion
    }
}
