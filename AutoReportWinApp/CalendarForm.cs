using System.Windows.Forms;

namespace AutoReportWinApp
{
    public partial class CalendarForm : Form
    {
        private InputDailyReportForm inputDailyReportForm;
        public CalendarForm(InputDailyReportForm inputDailyReportForm)
        {
            this.inputDailyReportForm = inputDailyReportForm;
            InitializeComponent();
        }

        private void mousePointer_DataSelected(object sender, DateRangeEventArgs e)
        {
            inputDailyReportForm.Controls[SetValue.AppConstants.InputCalenderDataPassFormItemNameStr].Text = this.monthCalendar1.SelectionRange.Start.Date.ToString(SpecialStr.AppConstants.DateFormat);
            this.Close();
        }
    }
}
