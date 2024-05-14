using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KDCLibrary;
using KDCLibrary.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using Bookmark = Microsoft.Office.Interop.Word.Bookmark;
using Document = Microsoft.Office.Interop.Word.Document;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace KDC_Outlook
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
            Outlook.Application app = Globals.ThisAddIn.Application;
            Inspector currentItem = app.ActiveInspector().CurrentItem;

            Document wordEditor = GetWordEditorItem(currentItem);

            switch (control.Id)
            {
                case "button1":
                    new CreditsForm().ShowDialog();
                    break;
                case "button2":
                    UpdateDatesFromCustomXmlParts(GetWordEditorItem(currentItem));
                    break;
                case "button3":
                    PopulateInsertDate();
                    break;
                case "button4": // Open the calendar control form
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

                            if (wordEditor != null && wordEditor.Application.Selection != null)
                            {
                                wordEditor.Application.Selection.Text =
                                    kDCService.ConvertDateBasedOnUserSelection(
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
                    };
                    form.ShowDialog();
                    break;
                case "splitButton2__btn": // Convert selected text
                    if (wordEditor != null && wordEditor.Application.Selection != null)
                    {
                        wordEditor.Application.Selection.Text =
                            kDCService.ConvertDateBasedOnUserSelection(
                                wordEditor.Application.Selection.Text,
                                IsReverse,
                                SelectedDialect,
                                SelectedFromFormat,
                                SelectedToFormat,
                                SelectedCalendar,
                                IsAddSuffix
                            );
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
        private void CleanOrphanBookmarks(object outlookItem)
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

            List<string> bookmarksToRemove = new List<string>();

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

            int formatChoice = kDCService.SelectFormatChoice(SelectedInsertFormat);
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

            string kurdishDate = kDCService.Kurdish(formatChoice, SelectedDialect, IsAddSuffix);
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
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

                ProcessItemInsert(inspector.CurrentItem, kurdishDate, formatChoice);
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

        private void ProcessItemInsert(object item, string kurdishDate, int formatChoice)
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
                    AddCustomXmlPart(
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

        private void AddCustomXmlPart(
            string bookmarkName,
            string dialect,
            int formatChoice,
            bool isAddSuffix,
            Document doc
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
            doc.CustomXMLParts.Add(customXml);
        }

        public void UpdateDatesFromCustomXmlParts(Document Doc)
        {
            List<Bookmark> bookmarksToUpdate = new List<Bookmark>();

            // Collect all relevant bookmarks to update
            foreach (Bookmark bookmark in Doc.Bookmarks)
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
                CustomXMLPart part = FindCustomXmlPartForBookmark(bookmarkName, Doc);

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

                    string newDate = kDCService.Kurdish(formatChoice, dialect, isAddSuffix);

                    // Replace old content and reset the bookmark
                    bookmarkRange.Text = newDate;

                    // Create a new range for the new text
                    Range newRange = Doc.Range(
                        bookmarkRange.Start,
                        bookmarkRange.Start + newDate.Length
                    );

                    Doc.Bookmarks.Add(bookmarkName, newRange);
                }
                else
                {
                    Debug.WriteLine($"No matching XML part found for bookmark: {bookmarkName}");
                }
            }
        }

        private CustomXMLPart FindCustomXmlPartForBookmark(string bookmarkName, Document doc)
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
        #endregion

        #endregion
    }
}
