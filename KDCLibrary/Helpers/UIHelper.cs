using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;

namespace KDCLibrary.Helpers
{
    internal class UIHelper
    {
        public void InitializeDropDown(
            RibbonDropDown dropDown,
            List<string> items,
            Microsoft.Office.Tools.Ribbon.RibbonFactory ribbonFactory
        )
        {
            dropDown.Items.Clear();
            foreach (var item in items)
            {
                var ribbonItem = ribbonFactory.CreateRibbonDropDownItem();
                ribbonItem.Label = item;
                dropDown.Items.Add(ribbonItem);
            }
            if (dropDown.Items.Count > 0)
            {
                dropDown.SelectedItem = dropDown.Items[0]; // Default to the first item
            }
        }

        public void ApplyCheckBoxStates(IEnumerable<RibbonCheckBox> checkBoxes, string savedState)
        {
            var states = savedState.Split(',');
            foreach (var checkBox in checkBoxes)
            {
                checkBox.Checked = states.Contains(checkBox.Name);
            }
        }

        public string GetCheckBoxStates(IEnumerable<RibbonCheckBox> checkBoxes)
        {
            var checkedStates = checkBoxes.Where(cb => cb.Checked).Select(cb => cb.Name).ToArray();
            return string.Join(",", checkedStates);
        }

        public void RestoreDropDownSelection(RibbonDropDown dropDown, string savedValue)
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

        public void UncheckOtherCheckBoxes(
            IEnumerable<RibbonCheckBox> checkBoxes,
            RibbonCheckBox exceptThis
        )
        {
            foreach (var checkBox in checkBoxes)
            {
                if (checkBox != exceptThis)
                {
                    checkBox.Checked = false;
                }
            }
        }
    }
}
