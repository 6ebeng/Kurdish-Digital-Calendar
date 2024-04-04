using Kurdish_Digital_Calendar.DateConversionLibrary;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Kurdish_Digital_Calendar
{
    public partial class Ribbon1
    {
        private const string RegistryPath = @"SOFTWARE\6ebeng\KurdishDigitalCalendar";
        private const string IsReverseKeyName = "IsReverse";
        private const string SelectedDialectKeyName = "SelectedDialect";
        private const string SelectedFormat1KeyName = "SelectedFormat1";
        private const string SelectedFormat2KeyName = "SelectedFormat2";
        private const string LastSelectionGroup1KeyName = "LastSelectionGroup1";
        private const string LastSelectionGroup2KeyName = "LastSelectionGroup2";
        private const string CheckBoxStatesKeyName = "CheckBoxStates";
        private const string isAddSuffixKeyName = "IsAddSuffix";

        private readonly List<string> _dialects = new List<string> { "Kurdish (Central)", "Kurdish (Northern)" };
        private readonly List<string> _formats = new List<string> { "dddd, dd MMMM, yyyy", "dddd, dd/MM/yyyy", "dd MMMM, yyyy", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy/MM/dd" };
        private readonly List<string> _calendarGroup1 = new List<string> { "Gregorian", "Hijri", "Umm al-Qura" };
        private readonly List<string> _calendarGroup2 = new List<string> { "Gregorian (English)", "Gregorian (Arabic)", "Gregorian (Kurdish Central)", "Gregorian (Kurdish Northern)", "Hijri (English)", "Hijri (Arabic)", "Hijri (Kurdish Central)", "Hijri (Kurdish Northern)", "Umm al-Qura (English)", "Umm al-Qura (Arabic)", "Umm al-Qura (Kurdish Central)", "Umm al-Qura (Kurdish Northern)", "Kurdish (Central)", "Kurdish (Northern)" };

        private void SaveState()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath))
            {
                if (key != null)
                {
                    key.SetValue(SelectedDialectKeyName, dropDown1.SelectedItem.Label);
                    SaveDropDownSelectionGroup(key);
                    key.SetValue(SelectedFormat1KeyName, dropDown3.SelectedItem.Label);
                    key.SetValue(SelectedFormat2KeyName, dropDown4.SelectedItem.Label);
                    key.SetValue(IsReverseKeyName, toggleButton1.Checked.ToString());
                    key.SetValue(CheckBoxStatesKeyName, GetCheckBoxStates());
                    key.SetValue(isAddSuffixKeyName, checkBox7.Checked.ToString());
                }
            }
        }

        private string GetCheckBoxStates()
        {
            var checkBoxes = new List<RibbonCheckBox> { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 };
            var checkedStates = checkBoxes.Where(cb => cb.Checked).Select(cb => cb.Name).ToArray();
            return string.Join(",", checkedStates);
        }

        private void SaveDropDownSelectionGroup(RegistryKey key)
        {
            string groupName = _calendarGroup1.Any(item => item == dropDown2.SelectedItem.Label) ? LastSelectionGroup1KeyName :
                               _calendarGroup2.Any(item => item == dropDown2.SelectedItem.Label) ? LastSelectionGroup2KeyName : null;
            if (groupName != null)
            {
                key.SetValue(groupName, dropDown2.SelectedItem.Label);
            }
        }

        // Consolidate checkbox click event handling
        private void CheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            if (sender is RibbonCheckBox clickedCheckBox)
            {
                foreach (var checkBox in new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 })
                {
                    checkBox.Checked = checkBox == clickedCheckBox;
                }
                SaveState();
                populateInsertDate();
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
                    MessageBox.Show("Unsupported target format selected.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return -1; // Indicates an unsupported format
            }
        }

        private void LoadState()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
            {
                if (key == null)
                {
                    SetDefaultControlStates();
                    return;
                }

                toggleButton1.Checked = bool.TryParse(key.GetValue(IsReverseKeyName, "false").ToString(), out var isReverse) && isReverse;
                checkBox7.Checked = bool.TryParse(key.GetValue(isAddSuffixKeyName, "false").ToString(), out var isAddSuffix) && isAddSuffix;

                PopulateDropDownBasedOnToggleButton();

                RestoreDropDownSelection(dropDown1, key.GetValue(SelectedDialectKeyName) as string);
                RestoreDropDownSelection(dropDown3, key.GetValue(SelectedFormat1KeyName) as string);
                RestoreDropDownSelection(dropDown4, key.GetValue(SelectedFormat2KeyName) as string);
                RestoreLastSelectionGroupDropDownSelection();

                ApplyCheckBoxStates(key.GetValue(CheckBoxStatesKeyName) as string);
            }
        }

        private void SetDefaultControlStates()
        {
            // Assuming default control state settings are encapsulated here
            // This method sets the default state for controls when no registry state exists
            toggleButton1.Checked = false;
            PopulateDropDownBasedOnToggleButton();
            checkBox1.Checked = true;
            UncheckOtherCheckBoxes(checkBox1);
            InitializeDropDown(dropDown1, _dialects);
            InitializeDropDown(dropDown3, _formats);
            InitializeDropDown(dropDown4, _formats);
        }

        private void ApplyCheckBoxStates(string savedState)
        {
            if (string.IsNullOrEmpty(savedState))
            {
                checkBox1.Checked = true;
                UncheckOtherCheckBoxes(checkBox1);
                return;
            }

            var states = savedState.Split(',');
            foreach (var checkBox in new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 })
            {
                checkBox.Checked = states.Contains(checkBox.Name);
            }
        }

        private void RestoreDropDownSelection(RibbonDropDown dropDown, string savedValue)
        {
            if (!string.IsNullOrEmpty(savedValue))
            {
                var itemToSelect = dropDown.Items.FirstOrDefault(item => item.Label == savedValue);
                if (itemToSelect != null)
                {
                    dropDown.SelectedItem = itemToSelect;
                }
            }
        }

        private void RestoreLastSelectionGroupDropDownSelection()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
            {
                if (key != null)
                {
                    var groupName = toggleButton1.Checked ? LastSelectionGroup2KeyName : LastSelectionGroup1KeyName;
                    var savedValue = key.GetValue(groupName) as string;
                    RestoreDropDownSelection(dropDown2, savedValue);
                }
            }
        }

        private void PopulateDropDownBasedOnToggleButton()
        {
            dropDown2.Items.Clear();
            var groupToUse = toggleButton1.Checked ? _calendarGroup2 : _calendarGroup1;
            foreach (var calendar in groupToUse)
            {
                var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = calendar;
                dropDown2.Items.Add(item);
            }
        }

        private void populateInsertDate()
        {
            // Find the checked checkbox among the ones you have.
            RibbonCheckBox checkedCheckBox = new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 }
                                              .FirstOrDefault(cb => cb.Checked);

            if (checkedCheckBox != null)
            {
                // Determine the format choice based on the label of the checked checkbox.
                int formatChoice = DetermineFormatChoiceFromCheckbox(checkedCheckBox.Label);
                if (formatChoice == -1) // If the format is unsupported or not found
                {
                    MessageBox.Show("No valid format selected.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return; // Exit the method if no valid format is selected
                }

                // Use the selection from dropDown1 for the dialect
                string dialect = dropDown1.SelectedItem.Label;
                bool isAddSuffix = checkBox7.Checked;


                // Call Insert Kurdish Date with the determined formatChoice and dialect and isAddSuffix
                Globals.ThisAddIn.Application.Selection.TypeText(InsertDate.Kurdish(formatChoice, dialect, isAddSuffix));

            }
            else
            {
                MessageBox.Show("No checkbox selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void InitializeDropDown(RibbonDropDown dropDown, List<string> items)
        {
            dropDown.Items.Clear();
            foreach (var item in items)
            {
                var ribbonItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                ribbonItem.Label = item;
                dropDown.Items.Add(ribbonItem);
            }
            if (dropDown.Items.Count > 0)
            {
                dropDown.SelectedItem = dropDown.Items[0]; // Default to the first item
            }
        }

        private void UncheckOtherCheckBoxes(RibbonCheckBox exceptThis)
        {
            foreach (var checkBox in new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 })
            {
                if (checkBox != exceptThis)
                {
                    checkBox.Checked = false;
                }
            }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Set up event handlers
            checkBox1.Click += CheckBox_Click;
            checkBox2.Click += CheckBox_Click;
            checkBox3.Click += CheckBox_Click;
            checkBox4.Click += CheckBox_Click;
            checkBox5.Click += CheckBox_Click;
            checkBox6.Click += CheckBox_Click;


            // Initialize DropDowns with default values
            InitializeDropDown(dropDown1, _dialects);
            InitializeDropDown(dropDown2, _calendarGroup1);
            InitializeDropDown(dropDown3, _formats);
            InitializeDropDown(dropDown4, _formats);

            // Load the saved state, which will also populate dropDown2 based on the checkbox
            LoadState();

        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            SaveState();
        }

        private void dropDown2_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            SaveState();
        }

        private void dropDown3_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            SaveState();
        }

        private void dropDown4_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            SaveState();
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            PopulateDropDownBasedOnToggleButton();
            RestoreLastSelectionGroupDropDownSelection();
            SaveState();

        }


        private void checkBox7_Click(object sender, RibbonControlEventArgs e)
        {
            SaveState();
        }

        private void splitButton2_Click(object sender, RibbonControlEventArgs e)
        {

            Globals.ThisAddIn.Application.Selection.Text = DateConversion.ConvertDateBasedOnUserSelection(Globals.ThisAddIn.Application.Selection.Text, toggleButton1.Checked, dropDown1.SelectedItem.Label, dropDown3.SelectedItem.Label, dropDown4.SelectedItem.Label, dropDown2.SelectedItem.Label, checkBox7.Checked);
        }

        private void splitButton1_Click_1(object sender, RibbonControlEventArgs e)
        {
            populateInsertDate();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            new Form1().Show();
        }
    }
}
