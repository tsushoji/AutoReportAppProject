using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        private int _createDataColNum;
        public InputDailyReportForm()
        {
            InitializeComponent();
        }
        public DailyReportDataListForm DailyReportDataListForm { get => _dailyReportDataListForm; set => _dailyReportDataListForm = value; }
        public CreateDataMode CreateDataMode { get => _createDataMode; set => _createDataMode = value; }
        public int CreateDataColNum { get => _createDataColNum; set => _createDataColNum = value; }
        public string TextBox1Text { get => this.textBox1.Text; set => this.textBox1.Text = value; }
        public string TextBox2Text { get => this.textBox2.Text; set => this.textBox2.Text = value; }
        public string TextBox3Text { get => this.textBox3.Text; set => this.textBox3.Text = value; }
        public string TextBox4Text { get => this.textBox4.Text; set => this.textBox4.Text = value; }
        public string Label1Text { get => this.label1.Text; }
        public string Label2Text { get => this.label2.Text; }
        public string Label3Text { get => this.label3.Text; }
        public string Label4Text { get => this.label4.Text; }

        private void buttonCalendar_Click(object sender, EventArgs e)
        {
            using (var calendarForm = new CalendarForm(this))
            {
                calendarForm.ShowDialog();
            }
        }

        private void buttonCreateData_Click(object sender, EventArgs e)
        {
            if (this.inputcheck(TextBox1Text, TextBox2Text, TextBox3Text, TextBox4Text))
            {
                switch (CreateDataMode)
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
            if (File.Exists(DailyReportDataListForm.CsvDailyReportDataPath))
            {
                File.Delete(DailyReportDataListForm.CsvDailyReportDataPath);
                using (var writeFileStream = new FileStream(DailyReportDataListForm.CsvDailyReportDataPath, FileMode.Create, FileAccess.Write))
                using (var writer = new StreamWriter(writeFileStream, Encoding.GetEncoding(SetValue.AppConstants.WinEncoding)))
                using (var csv = new CsvWriter(writer, CultureInfo.CurrentCulture))
                {
                    var dailyReport = new DailyReportEntity();
                    DailyReportDataListForm.DataGridView1.Rows.Add(CreateDataColNum.ToString(), TextBox1Text, TextBox2Text, TextBox3Text, TextBox4Text);
                    dailyReport.controlNum = CreateDataColNum.ToString();
                    dailyReport.date = TextBox1Text;
                    dailyReport.impContent = DailyReportEntity.replaceToStrWithUserNewLine(TextBox2Text);
                    dailyReport.schContent = DailyReportEntity.replaceToStrWithUserNewLine(TextBox3Text);
                    dailyReport.task = DailyReportEntity.replaceToStrWithUserNewLine(TextBox4Text);
                    DailyReportDataListForm._csvDailyReportDataMap.Add(DailyReportDataListForm._csvDailyReportDataMap.Count, dailyReport);
                    foreach (KeyValuePair<int, DailyReportEntity> keyValuePairt in DailyReportDataListForm._csvDailyReportDataMap)
                    {
                        csv.WriteRecord(keyValuePairt.Value);
                        csv.NextRecord();
                    }
                    CreateDataColNum++;
                }
            }
            MessageBox.Show(Properties.Resources.I0001);
            this.Close();
        }
        private void createDataUpdate()
        {
            if (File.Exists(DailyReportDataListForm.CsvDailyReportDataPath))
            {
                File.Delete(DailyReportDataListForm.CsvDailyReportDataPath);

                using (var writeFileStream = new FileStream(DailyReportDataListForm.CsvDailyReportDataPath, FileMode.Create, FileAccess.Write))
                using (var writer = new StreamWriter(writeFileStream, Encoding.GetEncoding(SetValue.AppConstants.WinEncoding)))
                using (var csv = new CsvWriter(writer, CultureInfo.CurrentCulture))
                {
                    int controlNum;
                    foreach (KeyValuePair<int, DailyReportEntity> keyValuePair in DailyReportDataListForm._csvDailyReportDataMap)
                    {
                        controlNum = Int32.Parse(keyValuePair.Value.controlNum);
                        if (controlNum == CreateDataColNum)
                        {
                            string[] writingData = { CreateDataColNum.ToString(), TextBox1Text, TextBox2Text, TextBox3Text, TextBox4Text };

                            for (var i = 0; i < writingData.Length; i++)
                            {
                                DailyReportDataListForm.DataGridView1.Rows[keyValuePair.Key].Cells[i].Value = writingData[i];
                            }

                            DailyReportDataListForm._csvDailyReportDataMap[keyValuePair.Key].controlNum = CreateDataColNum.ToString();
                            DailyReportDataListForm._csvDailyReportDataMap[keyValuePair.Key].date = DailyReportEntity.replaceToStrWithUserNewLine(TextBox1Text);
                            DailyReportDataListForm._csvDailyReportDataMap[keyValuePair.Key].impContent = DailyReportEntity.replaceToStrWithUserNewLine(TextBox2Text);
                            DailyReportDataListForm._csvDailyReportDataMap[keyValuePair.Key].schContent = DailyReportEntity.replaceToStrWithUserNewLine(TextBox3Text);
                            DailyReportDataListForm._csvDailyReportDataMap[keyValuePair.Key].task = DailyReportEntity.replaceToStrWithUserNewLine(TextBox4Text);
                        }
                        csv.WriteRecord(keyValuePair.Value);
                        csv.NextRecord();
                    }
                }

                MessageBox.Show(Properties.Resources.I0002);
                this.Close();
            }
        }

        private Boolean inputcheck(string inputDate, string inputImpContent, string inputScheContent, string inputTask)
        {
            var NotInputDateMsgEle = Label1Text.Substring(0, 2);
            var NotInputImpContentMsgEle = Label2Text.Substring(0, 4);
            var NotInputSchContentMsgEle = Label3Text.Substring(0, 4);
            var NotInputTaskMsgEle = Label4Text.Substring(0, 2);
            var pattern = RegExp.AppConstants.DateRegExp;
            var errorMsgEleList = new List<string>();
            var errorMsg = new StringBuilder();
            Boolean rtnFlag = true;

            if (string.IsNullOrEmpty(inputDate))
            {
                errorMsgEleList.Add(NotInputDateMsgEle);
            }

            if (string.IsNullOrEmpty(inputImpContent))
            {
                errorMsgEleList.Add(NotInputImpContentMsgEle);
            }

            if (string.IsNullOrEmpty(inputScheContent))
            {
                errorMsgEleList.Add(NotInputSchContentMsgEle);
            }

            if (string.IsNullOrEmpty(inputTask))
            {
                errorMsgEleList.Add(NotInputTaskMsgEle);
            }

            if (errorMsgEleList.Count > 0)
            {
                string errorPartialMsg = errorMsgEleList.Aggregate((i, j) => i + SpecialStr.AppConstants.ReadingPointStr + j);
                errorMsg.Append(Properties.Resources.E0003.Replace("{ARGFIRST}", errorPartialMsg));
            }

            if (!Regex.IsMatch(inputDate, pattern) || !isDate(inputDate))
            {
                if (errorMsg.Length > 0)
                {
                    errorMsg.Append(SpecialStr.AppConstants.NewLineStr);
                }
                errorMsg.Append(Properties.Resources.E0002);
            }

            if (!DuplicateCheck(inputDate, this.DailyReportDataListForm._csvDailyReportDataMap))
            {
                if (errorMsg.Length > 0)
                {
                    errorMsg.Append(SpecialStr.AppConstants.NewLineStr);
                }
                errorMsg.Append(Properties.Resources.E0004);
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
            var pattern = SpecialStr.AppConstants.SlashChar;
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
        private Boolean DuplicateCheck(string dateStr, Dictionary<int, DailyReportEntity> dailyReportData)
        {
            Boolean rtnFlag = true;

            if (dailyReportData.Count > 0)
            {
                foreach (KeyValuePair<int, DailyReportEntity> keyValuePair in dailyReportData)
                {
                    if (Int32.Parse(keyValuePair.Value.controlNum) == CreateDataColNum)
                    {
                        continue;
                    }
                    if (dateStr.Equals(keyValuePair.Value.date))
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
            DailyReportDataListForm.DataGridView1.CurrentCell = null;
        }
    }
}
