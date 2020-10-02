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

        private static readonly string InputCalenderDataFormItemName = "textBox1";
        private static readonly string DateFormat = "yyyy/MM/dd";

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <remarks>カレンダーフォームクラス</remarks>
        /// <param name="inputDailyReportForm">入力日報フォームオブジェクト</param>
        public CalendarForm(InputDailyReportForm inputDailyReportForm)
        {
            InitializeComponent();
            this.inputDailyReportForm = inputDailyReportForm;
        }

        /// <summary>
        /// カレンダーの日付が選択されたときのイベント
        /// </summary>
        /// <remarks>入力日報フォームの「日付」項目に選択した日付文字列を出力</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">カレンダーイベントに関わる引数</param>
        private void MousePointer_DataSelected(object sender, DateRangeEventArgs e)
        {
            // カレンダーフォームで取得した日付文字列を入力日報フォームの「日付」項目に入れる
            this.inputDailyReportForm.Controls[InputCalenderDataFormItemName].Text = monthCalendar1.SelectionRange.Start.Date.ToString(DateFormat);
            this.Close();
        }
    }
}
