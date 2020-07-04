using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoReportWinApp
{
    public enum CreateDataMode
    {
        APPEND,
        UPDATE
    }
    public partial class InputDailyReportForm : Form
    {
        private DailyReportDataListForm _dailyReportDataListForm;
        private CreateDataMode _createDataMode;
        private Dictionary<int, DailyReport> _csvDailyReportDataMap;
        private int _createDataUpdateColNum;
        public InputDailyReportForm()
        {
            InitializeComponent();
            this._csvDailyReportDataMap = new Dictionary<int, DailyReport>();
        }
        public DailyReportDataListForm DailyReportDataListForm { get => _dailyReportDataListForm; set => _dailyReportDataListForm = value; }
        public CreateDataMode CreateDataMode { get => _createDataMode; set => _createDataMode = value; }
        public int CreateDataUpdateColNum { get => _createDataUpdateColNum; set => _createDataUpdateColNum = value; }
        public string TextBox1Text { get => this.textBox1.Text; set => this.textBox1.Text = value; }
        public string TextBox2Text { get => this.textBox2.Text; set => this.textBox2.Text = value; }
        public string TextBox3Text { get => this.textBox3.Text; set => this.textBox3.Text = value; }
        public string TextBox4Text { get => this.textBox4.Text; set => this.textBox4.Text = value; }

        private void buttonCalendar_Click(object sender, EventArgs e)
        {
            using (var calendarForm = new CalendarForm(this))
            {
                calendarForm.ShowDialog();
            }
        }

        private void buttonCreateData_Click(object sender, EventArgs e)
        {
            this._csvDailyReportDataMap.Clear();

            if (File.Exists(StartMenuForm.CreateDataFilePath))
            {
                using (var readFileStream = new FileStream(StartMenuForm.CreateDataFilePath, FileMode.Open, FileAccess.Read))
                using (var streamReader = new StreamReader(readFileStream, Encoding.Default))
                {
                    var rowNumIndex = 0;
                    while (streamReader.Peek() >= 0)
                    {
                        var dailyReport = new DailyReport();
                        string readingLine = streamReader.ReadLine();
                        string[] rowArray = readingLine.Split(AppConstants.DailyReportDataLineSeparate);
                        int rowConNum = Int32.Parse(rowArray[0]);
                        dailyReport.ControlNum = rowConNum;
                        dailyReport.CsvDailyReportLine = readingLine;
                        this._csvDailyReportDataMap.Add(rowNumIndex, dailyReport);
                        rowNumIndex++;
                    }
                }
            }

            if (this.inputcheck(this.textBox1.Text, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text))
            {
                switch (this.CreateDataMode)
                {
                    case CreateDataMode.APPEND:
                        this.createDataAppend();
                        break;

                    case CreateDataMode.UPDATE:
                        this.createDataUpdate();
                        break;
                }
            }
        }

        private void buttonForDataList_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void createDataAppend()
        {
            using (var fileStream = new FileStream(StartMenuForm.CreateDataFilePath, FileMode.Append, FileAccess.Write))
            using (var streamWriter = new StreamWriter(fileStream, Encoding.Default))
            {
                var writingList = new List<string>();
                string[] writingData = { this.DailyReportDataListForm.CreateDataColNum.ToString(), this.textBox1.Text, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text };
                foreach (var writingEle in writingData)
                {
                    if (writingEle.Contains(AppConstants.NewLineStr))
                    {
                        writingList.Add(writingEle.Replace(AppConstants.NewLineStr, AppConstants.UserNewLineStr));
                    }
                    else
                    {
                        writingList.Add(writingEle);
                    }
                }
                streamWriter.WriteLine(writingList.Aggregate((i, j) => i + AppConstants.CreateDataAppendSeparate + j));
                this.DailyReportDataListForm.CreateDataColNum++;
            }
            MessageBox.Show(AppConstants.CreateDataAppendCmpMsg);
            this.Close();
        }

        private void createDataUpdate()
        {
            if (File.Exists(StartMenuForm.CreateDataFilePath))
            {
                File.Delete(StartMenuForm.CreateDataFilePath);

                using (var writeFileStream = new FileStream(StartMenuForm.CreateDataFilePath, FileMode.Create, FileAccess.Write))
                using (var streamWriter = new StreamWriter(writeFileStream, Encoding.Default))
                {
                    foreach (KeyValuePair<int, DailyReport> keyValuePair in this._csvDailyReportDataMap)
                    {
                        if (keyValuePair.Value.ControlNum == this.CreateDataUpdateColNum)
                        {
                            string[] writingData = { keyValuePair.Value.ControlNum.ToString(), this.textBox1.Text, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text };
                            keyValuePair.Value.CsvDailyReportLine = string.Join(",", writingData);
                        }

                        if (keyValuePair.Value.CsvDailyReportLine.Contains("\n"))
                        {
                            keyValuePair.Value.CsvDailyReportLine = keyValuePair.Value.CsvDailyReportLine.Replace(AppConstants.NewLineStr, AppConstants.UserNewLineStr);
                        }
                        streamWriter.WriteLine(keyValuePair.Value.CsvDailyReportLine);
                    }
                }

                MessageBox.Show(AppConstants.CreateDataUpdateCmpMsg);
                this.Close();
            }
        }

        private Boolean inputcheck(string inputDate, string inputImpContent, string inputScheContent, string inputTask)
        {
            var pattern = AppConstants.DateRegExp;
            var errorMsgEleList = new List<string>();
            var errorMsg = new StringBuilder();
            Boolean rtnFlag = true;

            if ((inputDate == null) || (inputDate.Length == 0))
            {
                errorMsgEleList.Add(AppConstants.NotInputDateMsgEle);
            }

            if ((inputImpContent == null) || (inputImpContent.Length == 0))
            {
                errorMsgEleList.Add(AppConstants.NotInputImpContentMsgEle);
            }

            if ((inputScheContent == null) || (inputScheContent.Length == 0))
            {
                errorMsgEleList.Add(AppConstants.NotInputSchContentMsgEle);
            }

            if ((inputTask == null) || (inputTask.Length == 0))
            {
                errorMsgEleList.Add(AppConstants.NotInputTaskMsgEle);
            }

            if (errorMsgEleList.Count > 0)
            {
                string errorPartialMsg = errorMsgEleList.Aggregate((i, j) => i + AppConstants.NotInputCheckItemMsgSeparate + j);
                errorMsg.Append(errorPartialMsg).Append(AppConstants.NotInputCheckItemMsgEnd);
            }

            if (!Regex.IsMatch(inputDate, pattern) || !isDate(inputDate))
            {
                if (errorMsg.Length > 0)
                {
                    errorMsg.Append(AppConstants.NewLineStr);
                }
                errorMsg.Append(AppConstants.NotInputDateFormatMsg);
            }

            if (!DuplicateCheck(inputDate, this._csvDailyReportDataMap))
            {
                if (errorMsg.Length > 0)
                {
                    errorMsg.Append(AppConstants.NewLineStr);
                }
                errorMsg.Append(AppConstants.DuplicateDailyReportDataMsg);
            }

            if (errorMsg.Length > 0)
            {
                MessageBox.Show(errorMsg.ToString());
                rtnFlag = false;
            }

            return rtnFlag;
        }

        private Boolean isDate(string dateStr)
        {
            var pattern = AppConstants.DateFormatCheckSeparate;
            string[] dateEleArray = dateStr.Split(pattern);
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
        private Boolean DuplicateCheck(string dateStr, Dictionary<int, DailyReport> dailyReportData)
        {
            Boolean rtnFlag = true;

            if (dailyReportData.Count > 0)
            {
                foreach (KeyValuePair<int, DailyReport> keyValuePair in dailyReportData)
                {
                    string[] dailyReportDataEle = keyValuePair.Value.CsvDailyReportLine.Split(AppConstants.DailyReportDataLineSeparate);
                    if (dateStr.Equals(dailyReportDataEle[1]))
                    {
                        rtnFlag = false;
                        break;
                    }
                }
            }
            return rtnFlag;
        }
        private void InputDailyReportForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.DailyReportDataListForm.initDailyReportDataReader(StartMenuForm.CreateDataFilePath);
            this.DailyReportDataListForm.DataGridView1.CurrentCell = null;
        }
    }
}
