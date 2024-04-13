using System;
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
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;

namespace KDC_Project
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        #region Intializers

        IKDCService kDCService = new KDCServiceImplementation();
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

            int formatChoice = kDCService.DetermineFormatChoiceFromCheckbox(_selectedInsertFormat);
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

            try
            {
                var application = Globals.ThisAddIn.Application;
                var selection = application.ActiveSelection;
                var activeCell = application.ActiveCell;
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

                string newText = kDCService.Kurdish(formatChoice, _selectedDialect, _isAddSuffix);

                // Update only the field corresponding to the active cell across all selected tasks.
                foreach (MSProject.Task task in selection.Tasks)
                {
                    if (task != null)
                    {
                        SetTaskFieldValue(task, activeFieldName, newText);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"An error occurred: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
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

        public void SplitButton2_Click(IRibbonControl control)
        {
            try
            {
                var application = Globals.ThisAddIn.Application;
                var selection = application.ActiveSelection;

                if (selection == null || selection.Tasks == null || selection.Tasks.Count == 0)
                {
                    MessageBox.Show(
                        "Please select one or more tasks.",
                        "Selection Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    return;
                }

                var activeCell = application.ActiveCell;
                string fieldName = activeCell.FieldName;

                foreach (MSProject.Task task in selection.Tasks)
                {
                    if (task == null)
                        continue;

                    // Extract the text based on the active field name
                    string fieldText = GetTaskFieldValue(task, fieldName);
                    string result = kDCService.ConvertDateBasedOnUserSelection(
                        fieldText,
                        _isReverse,
                        _selectedDialect,
                        _selectedFromFormat,
                        _selectedToFormat,
                        _selectedCalendar,
                        _isAddSuffix
                    );

                    if (!string.IsNullOrEmpty(result))
                    {
                        SetTaskFieldValue(task, fieldName, result);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"An error occurred: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private string GetTaskFieldValue(MSProject.Task task, string fieldName)
        {
            // Ideally, extend this to cover more fields as per your requirement
            switch (fieldName)
            {
                case "Name":
                    return task.Name;
                case "Notes":
                    return task.Notes;
                case "Resource Names":
                    return task.ResourceNames;
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
