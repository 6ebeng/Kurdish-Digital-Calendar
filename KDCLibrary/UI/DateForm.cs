using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using DevExpress.XtraEditors.Controls;
using KDCLibrary.CustomControls;

namespace KDCLibrary.UI
{
    [ComVisible(false)]
    public partial class DateForm : DevExpress.XtraEditors.XtraForm
    {
        private CustomXtraEditorsCalendarControl kurdishCalendarControl;

        public DateForm()
        {
            InitializeComponent();
            InitializeKurdishCalendar();
        }

        private void InitializeKurdishCalendar()
        {
            kurdishCalendarControl = new CustomXtraEditorsCalendarControl
            {
                //kurdishCalendar.RowCount = 3;
                //kurdishCalendar.ColumnCount = 4;
                //kurdishCalendar.ShowTodayButton = false;
                //kurdishCalendar.VistaCalendarViewStyle = DevExpress
                //    .XtraEditors
                //    .VistaCalendarViewStyle
                //    .MonthView;

                Dock = DockStyle.Fill,
                ShowToolTips = true,
                ShowTodayButton = true,
                BorderStyle = BorderStyles.NoBorder,
                RightToLeftLayout =
                    Ribbon.SelectedDialect == "Kurdish (Central)"
                        ? DevExpress.Utils.DefaultBoolean.True
                        : DevExpress.Utils.DefaultBoolean.False,
                RightToLeft =
                    Ribbon.SelectedDialect == "Kurdish (Central)"
                        ? RightToLeft.Yes
                        : RightToLeft.No,
                ShowWeekNumbers = true,
                AutoSizeInLayoutControl = true
            };

            //kurdishCalendarControl.Resize += (sender, e) =>
            //{
            //    this.Width = kurdishCalendarControl.Width;
            //    this.Height = kurdishCalendarControl.Height;
            //};

            Width = kurdishCalendarControl.Width;
            Height = kurdishCalendarControl.Height;
            //kurdishCalendar.RightToLeft = RightToLeft.Yes;
            this.Controls.Add(kurdishCalendarControl);
        }
    }
}
