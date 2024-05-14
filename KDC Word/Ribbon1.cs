using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using KDCLibrary;
using KDCLibrary;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;

namespace KDC_Word
{
    public partial class Ribbon1
    {
        IKDCService kdcService = new KDCServiceImplementation();

        private const string IsReverseKeyName = KDCConstants.KeyNames.IsReverse;
        private const string SelectedDialectKeyName = KDCConstants.KeyNames.SelectedDialect;
        private const string SelectedFormat1KeyName = KDCConstants.KeyNames.SelectedFormat1;
        private const string SelectedFormat2KeyName = KDCConstants.KeyNames.SelectedFormat2;
        private const string LastSelectionGroup1KeyName = KDCConstants.KeyNames.LastSelectionGroup1;
        private const string LastSelectionGroup2KeyName = KDCConstants.KeyNames.LastSelectionGroup2;
        private const string CheckBoxStatesKeyName = KDCConstants.KeyNames.CheckBoxStates;
        private const string isAddSuffixKeyName = KDCConstants.KeyNames.IsAddSuffix;

        private readonly List<string> _dialects = KDCConstants.DefaultValues.Dialects;
        private readonly List<string> _formats = KDCConstants.DefaultValues.Formats;
        private readonly List<string> _calendarGroup1 = KDCConstants.DefaultValues.CalendarGroup1;
        private readonly List<string> _calendarGroup2 = KDCConstants.DefaultValues.CalendarGroup2;

        private void SaveState()
        {
            kdcService.SaveSetting(SelectedDialectKeyName, dropDown1.SelectedItem.Label);
            kdcService.SaveSetting(SelectedFormat1KeyName, dropDown3.SelectedItem.Label);
            kdcService.SaveSetting(SelectedFormat2KeyName, dropDown4.SelectedItem.Label);
            kdcService.SaveSetting(IsReverseKeyName, toggleButton1.Checked.ToString());
            kdcService.SaveSetting(
                CheckBoxStatesKeyName,
                kdcService.GetCheckBoxStates(
                    new RibbonCheckBox[]
                    {
                        checkBox1,
                        checkBox2,
                        checkBox3,
                        checkBox4,
                        checkBox5,
                        checkBox6
                    }
                )
            );
            kdcService.SaveSetting(isAddSuffixKeyName, checkBox7.Checked.ToString());
            SaveDropDownSelectionGroup();
        }

        private void SaveDropDownSelectionGroup()
        {
            string groupName = _calendarGroup1.Any(item => item == dropDown2.SelectedItem.Label)
                ? LastSelectionGroup1KeyName
                : _calendarGroup2.Any(item => item == dropDown2.SelectedItem.Label)
                    ? LastSelectionGroup2KeyName
                    : null;
            if (groupName != null)
            {
                kdcService.SaveSetting(groupName, dropDown2.SelectedItem.Label);
            }
        }

        private void CheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            if (sender is RibbonCheckBox clickedCheckBox)
            {
                foreach (
                    var checkBox in new[]
                    {
                        checkBox1,
                        checkBox2,
                        checkBox3,
                        checkBox4,
                        checkBox5,
                        checkBox6
                    }
                )
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
                    MessageBox.Show(
                        "Unsupported target format selected.",
                        "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation
                    );
                    return -1; // Indicates an unsupported format
            }
        }

        private void LoadState()
        {
            toggleButton1.Checked =
                bool.TryParse(kdcService.LoadSetting(IsReverseKeyName, "false"), out bool isReverse)
                && isReverse;
            checkBox7.Checked =
                bool.TryParse(
                    kdcService.LoadSetting(isAddSuffixKeyName, "false"),
                    out bool isAddSuffix
                ) && isAddSuffix;

            PopulateDropDownBasedOnToggleButton();

            kdcService.RestoreDropDownSelection(
                dropDown1,
                kdcService.LoadSetting(SelectedDialectKeyName, "")
            );
            kdcService.RestoreDropDownSelection(
                dropDown3,
                kdcService.LoadSetting(SelectedFormat1KeyName, "")
            );
            kdcService.RestoreDropDownSelection(
                dropDown4,
                kdcService.LoadSetting(SelectedFormat2KeyName, "")
            );
            RestoreLastSelectionGroupDropDownSelection();

            kdcService.ApplyCheckBoxStates(
                new RibbonCheckBox[]
                {
                    checkBox1,
                    checkBox2,
                    checkBox3,
                    checkBox4,
                    checkBox5,
                    checkBox6
                },
                kdcService.LoadSetting(CheckBoxStatesKeyName, "")
            );
        }

        private void RestoreLastSelectionGroupDropDownSelection()
        {
            var groupName = toggleButton1.Checked
                ? LastSelectionGroup2KeyName
                : LastSelectionGroup1KeyName;
            var savedValue = kdcService.LoadSetting(groupName, "");
            kdcService.RestoreDropDownSelection(dropDown2, savedValue);
        }

        private void PopulateDropDownBasedOnToggleButton()
        {
            dropDown2.Items.Clear();
            var groupToUse = toggleButton1.Checked ? _calendarGroup2 : _calendarGroup1;
            kdcService.InitializeDropDown(
                dropDown2,
                groupToUse,
                Globals.Factory.GetRibbonFactory()
            );
        }

        private void populateInsertDate()
        {
            // Find the checked checkbox among the ones you have.
            RibbonCheckBox checkedCheckBox = new[]
            {
                checkBox1,
                checkBox2,
                checkBox3,
                checkBox4,
                checkBox5,
                checkBox6
            }.FirstOrDefault(cb => cb.Checked);

            if (checkedCheckBox != null)
            {
                // Determine the format choice based on the label of the checked checkbox.
                int formatChoice = DetermineFormatChoiceFromCheckbox(checkedCheckBox.Label);
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

                // Use the selection from dropDown1 for the dialect
                string dialect = dropDown1.SelectedItem.Label;
                bool isAddSuffix = checkBox7.Checked;

                // Call Insert Kurdish Date with the determined formatChoice and dialect and isAddSuffix
                Globals.ThisAddIn.Application.Selection.TypeText(
                    kdcService.Kurdish(formatChoice, dialect, isAddSuffix)
                );
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

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Register checkbox click event handlers to a single handler for simplification
            checkBox1.Click += CheckBox_Click;
            checkBox2.Click += CheckBox_Click;
            checkBox3.Click += CheckBox_Click;
            checkBox4.Click += CheckBox_Click;
            checkBox5.Click += CheckBox_Click;
            checkBox6.Click += CheckBox_Click;
            checkBox7.Click += checkBox7_Click;

            // Initialize DropDowns with default values using the UIHelper
            kdcService.InitializeDropDown(dropDown1, _dialects, Globals.Factory.GetRibbonFactory());
            kdcService.InitializeDropDown(dropDown3, _formats, Globals.Factory.GetRibbonFactory());
            kdcService.InitializeDropDown(dropDown4, _formats, Globals.Factory.GetRibbonFactory());

            // Populate and initialize dropDown2 based on the toggleButton1's state
            PopulateDropDownBasedOnToggleButton();

            // Load the saved state, which will also restore selections in dropdowns based on the saved registry values
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
            Globals.ThisAddIn.Application.Selection.Text =
                kdcService.ConvertDateBasedOnUserSelection(
                    Globals.ThisAddIn.Application.Selection.Text,
                    toggleButton1.Checked,
                    dropDown1.SelectedItem.Label,
                    dropDown3.SelectedItem.Label,
                    dropDown4.SelectedItem.Label,
                    dropDown2.SelectedItem.Label,
                    checkBox7.Checked
                );
        }

        private void splitButton1_Click(object sender, RibbonControlEventArgs e)
        {
            populateInsertDate();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            kdcService.Credits();
        }

        private void splitButton3_Click(object sender, RibbonControlEventArgs e) { }

        private void splitButton4_Click(object sender, RibbonControlEventArgs e) { }
    }
}
