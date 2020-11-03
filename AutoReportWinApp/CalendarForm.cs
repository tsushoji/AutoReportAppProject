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

        private static readonly string inputCalenderDataFormItemName = "textBox1";
        private static readonly string dateFormat = "yyyy/MM/dd";

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
        /// カレンダーフォームで日付が選択されたときのイベント
        /// </summary>
        /// <remarks>カレンダーフォームで選択された日付（文字列）を入力日報フォームの「日付」項目にセット</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">カレンダーイベントに関わる引数</param>
        private void MousePointer_DataSelected(object sender, DateRangeEventArgs e)
        {
            this.inputDailyReportForm.Controls[inputCalenderDataFormItemName].Text = monthCalendar1.SelectionRange.Start.Date.ToString(dateFormat);
            this.Close();
        }
    }
}
