﻿using CsvHelper;
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
    /// <remarks>入力日報フォームを入力し日報データを作成するためのフォーム</remarks>
    public partial class InputDailyReportForm : Form
    {
        private DailyReportDataListForm dailyReportDataListForm;
        private CreateDataMode createDataMode;
        private int createDataControlNum;

        public DailyReportDataListForm DailyReportDataListForm { get => dailyReportDataListForm; set => dailyReportDataListForm = value; }
        public CreateDataMode CreateDataMode { get => createDataMode; set => createDataMode = value; }
        public int CreateDataControlNum { get => createDataControlNum; set => createDataControlNum = value; }
        public string TextBox1Text { get => this.textBox1.Text; set => this.textBox1.Text = value; }
        public string TextBox2Text { get => this.textBox2.Text; set => this.textBox2.Text = value; }
        public string TextBox3Text { get => this.textBox3.Text; set => this.textBox3.Text = value; }
        public string TextBox4Text { get => this.textBox4.Text; set => this.textBox4.Text = value; }

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
        private void ButtonCreate_Click(object sender, EventArgs e)
        {
            if (!Inputcheck(TextBox1Text, TextBox2Text, TextBox3Text, TextBox4Text, DailyReportDataListForm.dailyReportDataList))
            {
                switch (CreateDataMode)
                {
                    case CreateDataMode.APPEND:
                        AppendDailyReportData();
                        break;

                    case CreateDataMode.UPDATE:
                        UpdateDailyReportData();
                        break;

                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// 「データリストへ」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonForDataList_Click(object sender, EventArgs e)
        {
            Close();
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
                    dailyReport.ControlNum = CreateDataControlNum.ToString();
                    dailyReport.DateStr = TextBox1Text;
                    dailyReport.ImplementationContent = DailyReportEntity.ReplaceToUserNewLineStr(TextBox2Text);
                    dailyReport.TomorrowPlan = DailyReportEntity.ReplaceToUserNewLineStr(TextBox3Text);
                    dailyReport.Task = DailyReportEntity.ReplaceToUserNewLineStr(TextBox4Text);
                    DailyReportDataListForm.dailyReportDataList.Add(dailyReport);
                    var dailyReportDataListOrderByControlNum = DailyReportDataListForm.dailyReportDataList.OrderBy(value => Int32.Parse(value.ControlNum)).ToList();
                    csv.Configuration.HasHeaderRecord = false;
                    csv.WriteRecords(dailyReportDataListOrderByControlNum);
                    DailyReportDataListForm.SetPageCountProperty(DailyReportDataListForm.dailyReportDataList);
                    if (DailyReportDataListForm.PageCount > 1)
                    {
                        DailyReportDataListForm.CurrentDailyReportDataIndex = DailyReportDataListForm.PageCount - 1;
                    }
                    DailyReportDataListForm.SetPagingDailyReportDataToDataGridView(DailyReportDataListForm.dailyReportDataList);
                    //日報データ作成後、管理番号を最新に更新
                    CreateDataControlNum++;
                }
            }
            MessageBox.Show(Properties.Resources.I0001);
            Close();
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
                    foreach (var dailyReportData in DailyReportDataListForm.dailyReportDataList)
                    {
                        var controlNum = Int32.Parse(dailyReportData.ControlNum);
                        //更新する管理番号は日報データリストフォームでセットした「CreateDataControlNum」プロパティーとする
                        if (controlNum == CreateDataControlNum)
                        {
                            dailyReportData.ControlNum = CreateDataControlNum.ToString();
                            dailyReportData.DateStr = DailyReportEntity.ReplaceToUserNewLineStr(TextBox1Text);
                            dailyReportData.ImplementationContent = DailyReportEntity.ReplaceToUserNewLineStr(TextBox2Text);
                            dailyReportData.TomorrowPlan = DailyReportEntity.ReplaceToUserNewLineStr(TextBox3Text);
                            dailyReportData.Task = DailyReportEntity.ReplaceToUserNewLineStr(TextBox4Text);
                        }
                    }
                    var dailyReportDataListOrderByControlNum = DailyReportDataListForm.dailyReportDataList.OrderBy(value => Int32.Parse(value.ControlNum)).ToList();
                    csv.Configuration.HasHeaderRecord = false;
                    csv.WriteRecords(dailyReportDataListOrderByControlNum);
                    DailyReportDataListForm.SetPagingDailyReportDataToDataGridView(DailyReportDataListForm.dailyReportDataList);
                }
                MessageBox.Show(Properties.Resources.I0002);
                Close();
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
        /// <param name="dailyReportDataList">日報データリスト</param>
        /// <returns>判定結果</returns>
        private Boolean Inputcheck(string inputDateStr, string inputImplementationContent, string inputTomorrowPlan, string inputTask, List<DailyReportEntity> dailyReportDataList)
        {
            string inputDateStrErrMsgEle = this.label1.Text.Substring(0, 2);
            string inputImplementationContentErrMsgEle = this.label2.Text.Substring(0, 4);
            string inputTomorrowPlanErrMsgEle = this.label3.Text.Substring(0, 4);
            string inputTaskErrMsgEle = this.label4.Text.Substring(0, 2);
            var errMsgEleList = new List<string>();
            var errMsg = new StringBuilder();
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
                string partialErrMsg = errMsgEleList.Aggregate((i, j) => i + DailyReportDataListForm.ReadingPointStr + j);
                errMsg.Append(Properties.Resources.E0001.Replace(DailyReportDataListForm.ReplaceErrMsgFirstArgStr, partialErrMsg));
            }
            if (!DuplicateCheck(inputDateStr, dailyReportDataList))
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
            }
            return errMsg.Length > 0;
        }

        /// <summary>
        /// 文字列が日付であるかチェック
        /// </summary>
        /// <remarks>日付年月日の数値をチェック</remarks>
        /// <param name="dateStr">日付文字列</param>
        /// <returns>判定結果</returns>
        private Boolean IsDate(string dateStr)
        {
            string[] dateEleArray = dateStr.Split(DailyReportDataListForm.SlashChar);
            var dateYear = Int32.Parse(dateEleArray[0]);
            var dateMonth = Int32.Parse(dateEleArray[1]);
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
        /// 日報作成データ「日付」重複チェック
        /// </summary>
        /// <param name="dateStr">日付文字列</param>
        /// <param name="dailyReportDataList">日報データリスト</param>
        /// <returns>判定結果</returns>
        private Boolean DuplicateCheck(string dateStr, List<DailyReportEntity> dailyReportDataList)
        {
            var rtnFlag = true;
            foreach (DailyReportEntity dailyReportData in dailyReportDataList)
            {
                if (Int32.Parse(dailyReportData.ControlNum) == CreateDataControlNum)
                {
                    continue;
                }
                if (dateStr.Equals(dailyReportData.DateStr))
                {
                    rtnFlag = false;
                    break;
                }
            }
            return rtnFlag;
        }

        /// <summary>
        /// 入力日報フォームを閉じたときのイベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void InputDailyReportForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            DailyReportDataListForm.DataGridView1.CurrentCell = null;
        }
    }
}
