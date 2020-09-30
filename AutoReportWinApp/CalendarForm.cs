using System.Windows.Forms;

namespace AutoReportWinApp
{
    /// <summary>
    /// カレンダーフォームクラス
    /// </summary>
    /// <remarks>カレンダーのためのフォーム</remarks>
    public partial class CalendarForm : Form
    {
        private InputDailyReportForm inputDailyReportForm;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <remarks>カレンダーフォームクラス</remarks>
        /// <param name="inputDailyReportForm">入力日報フォームオブジェクト</param>
        public CalendarForm(InputDailyReportForm inputDailyReportForm)
        {
            this.inputDailyReportForm = inputDailyReportForm;
            InitializeComponent();
        }

        /// <summary>
        /// カレンダーの日付が選択されたときのイベント
        /// </summary>
        /// <remarks>入力日報フォームの「日付」項目に選択した日付文字列を出力</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">カレンダーイベントに関わる引数</param>
        private void mousePointer_DataSelected(object sender, DateRangeEventArgs e)
        {
            inputDailyReportForm.Controls[SetValue.AppConstants.InputCalenderDataPassFormItemNameStr].Text = this.monthCalendar1.SelectionRange.Start.Date.ToString(SpecialStr.AppConstants.DateFormat);
            this.Close();
        }
    }
}
