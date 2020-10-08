using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoReportWinApp
{

    /// <summary>
    /// データ作成モード
    /// </summary>
    /// <remarks>APPEND：追加、UPDATE：更新</remarks>
    public enum CreateDataMode
    {
        APPEND,
        UPDATE
    }

    /// <summary>
    /// 入力日報フォームクラス
    /// </summary>
    /// <remarks>日報を入力し作成するためのフォーム</remarks>
    public partial class InputDailyReportForm : Form
    {
        private DailyReportDataListForm dailyReportDataListForm;
        private CreateDataMode createDataMode;
        private int createDataControlNum;

        private static readonly string RegExpDate = @"^\d{4}/\d{2}/\d{2}$";
        private static readonly char slashChar = '/';
        private static readonly string readingPointStr = "、";
        private static readonly string replaceErrMsgFirstArgStr = "{FIRSTARG}";

        public DailyReportDataListForm DailyReportDataListForm { get => dailyReportDataListForm; set => dailyReportDataListForm = value; }
        public CreateDataMode CreateDataMode { get => createDataMode; set => createDataMode = value; }
        public int CreateDataControlNum { get => createDataControlNum; set => createDataControlNum = value; }
        public string TextBox1Text { get => this.textBox1.Text; set => this.textBox1.Text = value; }
        public string TextBox2Text { get => this.textBox2.Text; set => this.textBox2.Text = value; }
        public string TextBox3Text { get => this.textBox3.Text; set => this.textBox3.Text = value; }
        public string TextBox4Text { get => this.textBox4.Text; set => this.textBox4.Text = value; }
        public static string ReadingPointStr { get => readingPointStr; }
        public static string ReplaceErrMsgFirstArgStr { get => replaceErrMsgFirstArgStr; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <remarks>入力日報フォームクラス</remarks>
        public InputDailyReportForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 「カレンダー」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonCalendar_Click(object sender, EventArgs e)
        {
            using (var calendarForm = new CalendarForm(this))
            {
                calendarForm.ShowDialog();
            }
        }

        /// <summary>
        /// 「データ作成」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonCreateData_Click(object sender, EventArgs e)
        {
            if (this.Inputcheck(TextBox1Text, TextBox2Text, TextBox3Text, TextBox4Text))
            {
                switch (CreateDataMode)
                {
                    //新規追加
                    case CreateDataMode.APPEND:
                        this.AppendDailyReportData();
                        break;

                    //更新
                    case CreateDataMode.UPDATE:
                        this.UpdateDailyReportData();
                        break;

                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// 「データリストへ」ボタン押下時、イベント
        /// </summary>
        /// <remarks>「日報データリストフォーム」へ遷移</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonToDailyReportDataListForm_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 日報データ新規作成
        /// </summary>
        private void AppendDailyReportData()
        {
            if (File.Exists(DailyReportDataListForm.CsvDailyReportDataPath))
            {
                File.Delete(DailyReportDataListForm.CsvDailyReportDataPath);
                using (var writeFileStream = new FileStream(DailyReportDataListForm.CsvDailyReportDataPath, FileMode.Create, FileAccess.Write))
                using (var writer = new StreamWriter(writeFileStream, Encoding.GetEncoding(DailyReportDataListForm.WinCharCode)))
                using (var csv = new CsvWriter(writer, CultureInfo.CurrentCulture))
                {
                    var dailyReport = new DailyReportEntity();
                    //データグリッドビューに日報データ新規追加
                    DailyReportDataListForm.DataGridView1.Rows.Add(CreateDataControlNum.ToString(), TextBox1Text, TextBox2Text, TextBox3Text, TextBox4Text);
                    dailyReport.ControlNum = CreateDataControlNum.ToString();
                    dailyReport.DateStr = TextBox1Text;
                    dailyReport.ImplementationContent = DailyReportEntity.ReplaceToUserNewLineStr(TextBox2Text);
                    dailyReport.TomorrowPlan = DailyReportEntity.ReplaceToUserNewLineStr(TextBox3Text);
                    dailyReport.Task = DailyReportEntity.ReplaceToUserNewLineStr(TextBox4Text);
                    DailyReportDataListForm.CsvDailyReportDataMap.Add(DailyReportDataListForm.CsvDailyReportDataMap.Count, dailyReport);
                    //日報データファイルに書き込み
                    foreach (KeyValuePair<int, DailyReportEntity> keyValuePair in DailyReportDataListForm.CsvDailyReportDataMap)
                    {
                        csv.WriteRecord(keyValuePair.Value);
                        csv.NextRecord();
                    }
                    //日報データ作成後、管理番号を最新に更新
                    CreateDataControlNum++;
                }
            }
            MessageBox.Show(Properties.Resources.I0001);
            this.Close();
        }

        /// <summary>
        /// 日報データ更新
        /// </summary>
        private void UpdateDailyReportData()
        {
            if (File.Exists(DailyReportDataListForm.CsvDailyReportDataPath))
            {
                File.Delete(DailyReportDataListForm.CsvDailyReportDataPath);

                using (var writeFileStream = new FileStream(DailyReportDataListForm.CsvDailyReportDataPath, FileMode.Create, FileAccess.Write))
                using (var writer = new StreamWriter(writeFileStream, Encoding.GetEncoding(DailyReportDataListForm.WinCharCode)))
                using (var csv = new CsvWriter(writer, CultureInfo.CurrentCulture))
                {
                    int controlNum;
                    foreach (KeyValuePair<int, DailyReportEntity> keyValuePair in DailyReportDataListForm.CsvDailyReportDataMap)
                    {
                        controlNum = Int32.Parse(keyValuePair.Value.ControlNum);
                        //更新する管理番号は日報データリストフォームで「CreateDataControlNum」プロパティーにセット
                        if (controlNum == CreateDataControlNum)
                        {
                            string[] writingData = { CreateDataControlNum.ToString(), TextBox1Text, TextBox2Text, TextBox3Text, TextBox4Text };

                            for (var i = 0; i < writingData.Length; i++)
                            {
                                DailyReportDataListForm.DataGridView1.Rows[keyValuePair.Key].Cells[i].Value = writingData[i];
                            }

                            DailyReportDataListForm.CsvDailyReportDataMap[keyValuePair.Key].ControlNum = CreateDataControlNum.ToString();
                            DailyReportDataListForm.CsvDailyReportDataMap[keyValuePair.Key].DateStr = DailyReportEntity.ReplaceToUserNewLineStr(TextBox1Text);
                            DailyReportDataListForm.CsvDailyReportDataMap[keyValuePair.Key].ImplementationContent = DailyReportEntity.ReplaceToUserNewLineStr(TextBox2Text);
                            DailyReportDataListForm.CsvDailyReportDataMap[keyValuePair.Key].TomorrowPlan = DailyReportEntity.ReplaceToUserNewLineStr(TextBox3Text);
                            DailyReportDataListForm.CsvDailyReportDataMap[keyValuePair.Key].Task = DailyReportEntity.ReplaceToUserNewLineStr(TextBox4Text);
                        }
                        //日報データファイルに書き込み
                        csv.WriteRecord(keyValuePair.Value);
                        csv.NextRecord();
                    }
                }

                MessageBox.Show(Properties.Resources.I0002);
                this.Close();
            }
        }

        /// <summary>
        /// 日報データ作成時、入力チェック
        /// </summary>
        /// <remarks>メッセージボックスにエラーメッセージ出力</remarks>
        /// <param name="inputDateStr">「日付」項目入力値</param>
        /// <param name="inputImplementationContent">「実施内容」項目入力値</param>
        /// <param name="inputTomorrowPlan">「翌日予定」項目入力値</param>
        /// <param name="inputTask">「課題」項目入力値</param>
        /// <returns>判定結果</returns>
        private Boolean Inputcheck(string inputDateStr, string inputImplementationContent, string inputTomorrowPlan, string inputTask)
        {
            //日報作成時、入力項目取得
            var inputDateStrErrMsgEle = this.label1.Text.Substring(0, 2);
            var inputImplementationContentErrMsgEle = this.label2.Text.Substring(0, 4);
            var inputTomorrowPlanErrMsgEle = this.label3.Text.Substring(0, 4);
            var inputTaskErrMsgEle = this.label4.Text.Substring(0, 2);

            var pattern = RegExpDate;
            var errMsgEleList = new List<string>();
            var errMsg = new StringBuilder();
            Boolean rtnFlag = true;

            if (string.IsNullOrEmpty(inputDateStr))
            {
                errMsgEleList.Add(inputDateStrErrMsgEle);
            }

            if (string.IsNullOrEmpty(inputImplementationContent))
            {
                errMsgEleList.Add(inputImplementationContentErrMsgEle);
            }

            if (string.IsNullOrEmpty(inputTomorrowPlan))
            {
                errMsgEleList.Add(inputTomorrowPlanErrMsgEle);
            }

            if (string.IsNullOrEmpty(inputTask))
            {
                errMsgEleList.Add(inputTaskErrMsgEle);
            }

            if (errMsgEleList.Count > 0)
            {
                string partialErrMsg = errMsgEleList.Aggregate((i, j) => i + readingPointStr + j);
                errMsg.Append(Properties.Resources.E0001.Replace(replaceErrMsgFirstArgStr, partialErrMsg));
            }

            if (!DuplicateCheck(inputDateStr, DailyReportDataListForm.CsvDailyReportDataMap))
            {
                if (errMsg.Length > 0)
                {
                    errMsg.Append(DailyReportEntity.NewLineStr);
                }
                errMsg.Append(Properties.Resources.E0003);
            }

            if (errMsg.Length > 0)
            {
                MessageBox.Show(errMsg.ToString());
                rtnFlag = false;
            }

            return rtnFlag;
        }

        /// <summary>
        /// 文字列が日付であるかチェック
        /// </summary>
        /// <remarks>日付年月日の数値をチェック</remarks>
        /// <param name="dateStr">日付文字列</param>
        /// <returns>判定結果</returns>
        private Boolean IsDate(string dateStr)
        {
            string[] dateEleArray = dateStr.Split(slashChar);
            int dateYear = Int32.Parse(dateEleArray[0]);
            int dateMonth = Int32.Parse(dateEleArray[1]);
            if (DateTime.MinValue.Year > dateYear || DateTime.MaxValue.Year < dateYear)
            {
                return false;
            }

            if (DateTime.MinValue.Month > dateMonth || DateTime.MaxValue.Month < dateMonth)
            {
                return false;
            }

            int dateLastDayNum = DateTime.DaysInMonth(dateYear, dateMonth);
            if (DateTime.MinValue.Day > Int32.Parse(dateEleArray[2]) || dateLastDayNum < Int32.Parse(dateEleArray[2]))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 日報作成データ重複チェック
        /// </summary>
        /// <remarks>「日付」で日報作成データ重複をチェック</remarks>
        /// <param name="dateStr">日付文字列</param>
        /// <param name="dailyReportData">作成済み日報データ</param>
        /// <returns>dataの平均値(出力)</returns>
        private Boolean DuplicateCheck(string dateStr, Dictionary<int, DailyReportEntity> dailyReportData)
        {
            Boolean rtnFlag = true;

            if (dailyReportData.Count > 0)
            {
                foreach (KeyValuePair<int, DailyReportEntity> keyValuePair in dailyReportData)
                {
                    if (Int32.Parse(keyValuePair.Value.ControlNum) == CreateDataControlNum)
                    {
                        continue;
                    }
                    if (dateStr.Equals(keyValuePair.Value.DateStr))
                    {
                        rtnFlag = false;
                        break;
                    }
                }
            }
            return rtnFlag;
        }

        /// <summary>
        /// 入力日報フォームを閉じたときのイベント
        /// </summary>
        /// <remarks>日報データリストフォーム遷移後のデータグリッドビューフォーカスをクリア</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void InputDailyReportForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //データグリッドビューフォーカスをクリア
            DailyReportDataListForm.DataGridView1.CurrentCell = null;
        }
    }
}
