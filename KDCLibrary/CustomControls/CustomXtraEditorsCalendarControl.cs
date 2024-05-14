using System;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Utils;
using DevExpress.Utils.Drawing;
using DevExpress.Utils.Extensions;
using DevExpress.XtraEditors.Calendar;
using DevExpress.XtraEditors.Controls;
using KDCLibrary.Calendars;
using KDCLibrary.Helpers;

namespace KDCLibrary.CustomControls
{
    internal class CustomXtraEditorsCalendarControl : CalendarControl
    {
        public static DateTime _gregorianSelectedDate { get; set; }
        public static bool _isClosedByDoubleClick { get; private set; } = false;

        public CustomXtraEditorsCalendarControl()
        {
            this.DateFormat =
                Ribbon.SelectedDialect == "Kurdish (Central)"
                    ? new CultureSetup().GetKurdishCentralDTFI()
                    : new CultureSetup().GetKurdishNorthernDTFI();

            this.MinValue = new DateTime(700, 1, 1); // Start of Kurdish year
            this.MaxValue = new DateTime(9999, 12, 31); // End of Kurdish year

            //this.InactiveDaysVisibility = CalendarInactiveDaysVisibility.Hidden;

            this.CustomWeekDayAbbreviation += KurdishCalendarControl_CustomWeekDayAbbreviation;

            this.MouseDoubleClick += KurdishCalendarControl_MouseDoubleClick;

            this.CustomDrawDayNumberCell += KurdishCalendarControl_CustomDrawDayNumberCell;
        }

        private void KurdishCalendarControl_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            CalendarHitInfo hitInfo = this.GetHitInfo(e.Location);
            if (hitInfo.HitTest == CalendarHitInfoType.MonthNumber && hitInfo.Cell.Selected)
            {
                if (this.ToolTipController != null)
                {
                    this.ToolTipController.Dispose();
                }
                _isClosedByDoubleClick = true;
                this.FindForm()?.Close();
            }
        }

        // Ensure to reset the flag when the control is being shown again or reused
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (this.Visible)
            {
                _isClosedByDoubleClick = false; // Reset flag when control becomes visible again
            }
        }

        private void KurdishCalendarControl_CustomWeekDayAbbreviation(
            object sender,
            CustomWeekDayAbbreviationEventArgs e
        )
        {
            switch (e.Day)
            {
                case "Sunday":
                case "یەکشەممە":
                case "Yekşem":
                    e.Value =
                        Ribbon.SelectedDialect == "Kurdish (Central)" ? "یه‌كشه‌ممه‌‌" : "Yekşem";
                    break;
                case "Monday":
                case "دووشەممە":
                case "Duşem":
                    e.Value =
                        Ribbon.SelectedDialect == "Kurdish (Central)" ? "دووشه‌ممه‌‌" : "Duşem";
                    break;
                case "Tuesday":
                case "سێشەممە":
                case "Sêşem":
                    e.Value =
                        Ribbon.SelectedDialect == "Kurdish (Central)" ? "سێشه‌ممه‌‌" : "Sêşem";
                    break;
                case "Wednesday":
                case "چوارشەممە":
                case "Çarşem":
                    e.Value =
                        Ribbon.SelectedDialect == "Kurdish (Central)" ? "چوارشه‌ممه‌" : "Çarşem";
                    break;
                case "Thursday":
                case "پێنجشەممە":
                case "Pêncşem":
                    e.Value =
                        Ribbon.SelectedDialect == "Kurdish (Central)" ? "پێنجشه‌ممه‌" : "Pêncşem";
                    break;
                case "Friday":
                case "هەینی":
                case "Înê":
                    e.Value = Ribbon.SelectedDialect == "Kurdish (Central)" ? "هه‌ینی" : "Înê";
                    break;
                case "Saturday":
                case "شەممە":
                case "Şemî":
                    e.Value = Ribbon.SelectedDialect == "Kurdish (Central)" ? "شه‌ممه‌" : "Şemî";
                    break;
            }
        }

        private void KurdishCalendarControl_CustomDrawDayNumberCell(
            object sender,
            CustomDrawDayNumberCellEventArgs e
        )
        {
            StringFormat stringFormatKurdish = new StringFormat();
            StringFormat stringFormatGregorian = new StringFormat();

            stringFormatGregorian.Alignment = StringAlignment.Center;
            stringFormatGregorian.LineAlignment = StringAlignment.Center;

            stringFormatKurdish.Alignment = StringAlignment.Far;
            stringFormatKurdish.LineAlignment = StringAlignment.Far;

            Brush textColor = new SolidBrush(Color.Black);
            Brush textColorGregorian = new SolidBrush(Color.Gray);
            Brush backgroundColor = new SolidBrush(e.Style.BackColor);

            if (CalendarDateEditing) { }

            if (IsPopupCalendar) { }

            if (IsMultiMonthView) { }

            if (e.Inactive)
            {
                textColor = new SolidBrush(Color.Gray);
            }

            if (e.Holiday) { }

            if (e.Disabled) { }

            if (e.Handled) { }

            if (e.Bounds.IsEmpty) { }

            if (e.Selected) // active cell
            {
                backgroundColor = new SolidBrush(Color.LightSteelBlue);
                textColor = new SolidBrush(Color.Black);
                textColorGregorian = new SolidBrush(Color.Black);
            }

            if (e.Today)
            {
                textColor = new SolidBrush(Color.White);
                backgroundColor = new SolidBrush(Color.FromArgb(180, 0, 24, 168));
                textColorGregorian = new SolidBrush(Color.White);
            }

            if (e.Highlighted) // hover cell
            {
                backgroundColor = new SolidBrush(Color.LightSteelBlue);
                textColor = new SolidBrush(Color.Black);
                textColorGregorian = new SolidBrush(Color.Black);
            }

            if (e.IsPressed) { }

            // fill aqua background cell if selected
            e.Cache.FillRectangle(backgroundColor, e.Bounds);

            FontFamily fontFamily = new FontFamily("Calibri");
            // Gregorian small font size
            Font kFont = new Font(e.Style.Font.FontFamily, 7, FontStyle.Regular);
            Font gFont = !e.Inactive
                ? new Font(e.Style.Font.FontFamily, 8, FontStyle.Bold)
                : new Font(e.Style.Font.FontFamily, 8, FontStyle.Regular);

            switch (e.View)
            {
                case DateEditCalendarViewType.MonthInfo:

                    if (e.Selected)
                        _gregorianSelectedDate = e.Date;

                    string[] kDateParts = (
                        new KurdishDate().FromGregorianToKurdish(
                            e.Date,
                            13,
                            Ribbon.SelectedDialect,
                            false
                        )
                    ).Split('/');

                    e.Graphics.DrawString(
                        e.Date.Day.ToString(),
                        gFont,
                        textColor,
                        e.Bounds,
                        stringFormatGregorian
                    );

                    e.Graphics.DrawString(
                        kDateParts[0],
                        kFont,
                        textColorGregorian,
                        e.Bounds.ApplyPadding(new Padding(0, 0, 4, 0)),
                        stringFormatKurdish
                    );

                    if (e.State == ObjectState.Hot && e.Selected)
                    {
                        this.ToolTipController = new ToolTipController();

                        ToolTipControlInfo toolTipInfo = new ToolTipControlInfo();
                        toolTipInfo.Object = "";
                        toolTipInfo.Text = new KurdishDate().FromGregorianToKurdish(
                            e.Date,
                            14,
                            Ribbon.SelectedDialect,
                            true
                        );
                        toolTipInfo.Title = kDateParts[0];
                        // toolTipInfo.HideHintOnMouseMove = true;


                        toolTipInfo.ToolTipType = ToolTipType.Standard;
                        toolTipInfo.ToolTipLocation = ToolTipLocation.BottomCenter;
                        toolTipInfo.ToolTipAnchor = ToolTipAnchor.Object;

                        toolTipInfo.SuperTip = new SuperToolTip();
                        toolTipInfo.SuperTip.Items.Add(toolTipInfo.Title);
                        toolTipInfo.SuperTip.Items.Add(toolTipInfo.Text);

                        this.ToolTipController.KeepWhileHovered = false;
                        //this.ToolTipController.InitialDelay = 1000;
                        //this.ToolTipController.ReshowDelay = 3000;
                        //this.ToolTipController.AutoPopDelay = 1000;

                        this.ToolTipController.ShowHint(toolTipInfo);
                    }

                    e.Handled = true; // Indicate that the day number drawing is handled
                    break;
                case DateEditCalendarViewType.QuarterInfo:
                    break;
                case DateEditCalendarViewType.YearInfo:
                    e.Graphics.DrawString(
                        e.Date.Month.ToString(),
                        new Font(fontFamily, 8, FontStyle.Regular),
                        new SolidBrush(Color.Gray),
                        e.Bounds.ApplyPadding(new Padding(0, 0, 80, 27)),
                        stringFormatGregorian
                    );

                    e.Graphics.DrawString(
                        Ribbon.SelectedDialect == "Kurdish (Central)"
                            ? new GregorianDate().GregorianMonthNameKurdishCentral(
                                int.Parse(e.Date.Month.ToString())
                            )
                            : new GregorianDate().GregorianMonthNameKurdishNorthern(
                                int.Parse(e.Date.Month.ToString())
                            ),
                        new Font(fontFamily, 10, FontStyle.Bold),
                        textColor,
                        e.Bounds,
                        stringFormatGregorian
                    );

                    e.Handled = true; // Indicate that the day number drawing is handled
                    break;
                case DateEditCalendarViewType.YearsInfo:
                    break;
                case DateEditCalendarViewType.YearsGroupInfo:
                    break;

                default:

                    break;
            }
        }
    }
}
