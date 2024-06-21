using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using DevExpress.XtraBars;
using KDCLibrary.Helpers;

namespace KDCLibrary.UI
{
    [ComVisible(false)]
    public partial class Settings : DevExpress.XtraEditors.XtraForm
    {
        private BarManager barManager1;
        private BarManager barManager2;

        public Settings()
        {
            InitializeComponent();
            InitializeDialectSelector();
            InitializeThemeSelector();
            SetupCheckBoxes();

            FormClosing += (sender, e) =>
            {
                barManager1?.Dispose();
                barManager2?.Dispose();
            };
        }

        private void InitializeDialectSelector()
        {
            barManager1 = new BarManager();
            PopupMenu popupMenu = new PopupMenu(barManager1) { MinWidth = dropDownButton1.Width };
            var dialectItems = Ribbon
                ._dialectsList.Select(dialect => new BarButtonItem(barManager1, dialect))
                .ToArray();
            popupMenu.AddItems(dialectItems);
            dropDownButton1.DropDownControl = popupMenu;
            dropDownButton1.Text = Ribbon.SelectedDialect ?? "Select Dialect";

            barManager1.ItemClick += (sender, e) =>
            {
                dropDownButton1.Text = e.Item.Caption;
                Ribbon.SelectedDialect = e.Item.Caption;
                new RegistryHelper().SaveSetting(
                    Ribbon.SelectedDialectKeyName,
                    e.Item.Caption,
                    Ribbon.AppName
                );
            };
        }

        private void InitializeThemeSelector()
        {
            barManager2 = new BarManager();
            PopupMenu popupMenu2 = new PopupMenu(barManager2) { MinWidth = dropDownButton2.Width };
            var themeItems = Ribbon
                ._themesList.Select(theme => new BarButtonItem(barManager2, theme))
                .ToArray();
            popupMenu2.AddItems(themeItems);
            dropDownButton2.DropDownControl = popupMenu2;
            dropDownButton2.Text = Ribbon.SelectedTheme ?? "Select Theme";

            barManager2.ItemClick += (sender, e) =>
            {
                dropDownButton2.Text = e.Item.Caption;
                Ribbon.SelectedTheme = e.Item.Caption;
                new RegistryHelper().SaveSetting(
                    Ribbon.ThemeColorKeyName,
                    e.Item.Caption,
                    Ribbon.AppName
                );
                Ribbon.ribbon.Invalidate();
            };
        }

        private void SetupCheckBoxes()
        {
            checkBox1.Checked = Ribbon.IsAddSuffix;
            checkBox2.Checked = Ribbon.IsAutoUpdateOnLoadDoc;
            checkBox1.CheckedChanged += (sender, e) =>
            {
                Ribbon.IsAddSuffix = checkBox1.Checked;
                new RegistryHelper().SaveSetting(
                    Ribbon.isAddSuffixKeyName,
                    checkBox1.Checked.ToString(),
                    Ribbon.AppName
                );
            };
            checkBox2.CheckedChanged += (sender, e) =>
            {
                Ribbon.IsAutoUpdateOnLoadDoc = checkBox2.Checked;
                new RegistryHelper().SaveSetting(
                    Ribbon.isAutoUpdateOnLoadDocKeyName,
                    checkBox2.Checked.ToString(),
                    Ribbon.AppName
                );
            };
            if (Ribbon.VisioApp != null || Ribbon.ProjectApp != null)
            {
                checkBox2.Enabled = false;
            }
        }
    }
}
