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
    public partial class InputDailyReportForm : Form
    {
        public static int controlNum = 0;
        private static readonly string _createDataFilePath = "C:/Users/Shota Tsuji/Desktop/DailyReportData.csv";
        public InputDailyReportForm()
        {
            InitializeComponent();
        }
        public static string CreateDataFilePath { get => _createDataFilePath; }

        private void InputDailyReportForm_Load(object sender, EventArgs e)
        {

        }

        private void buttonCalendar_Click(object sender, EventArgs e)
        {
            using(var calendarForm = new CalendarForm(this))
            {
                calendarForm.ShowDialog();
            }
        }

        private void buttonCreateData_Click(object sender, EventArgs e)
        {
            if (this.inputcheck(this.textBox1.Text, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text)) 
            {

                using (var fileStream = new FileStream(InputDailyReportForm.CreateDataFilePath, FileMode.Append, FileAccess.Write))
                using (var streamWriter = new StreamWriter(fileStream, Encoding.Default))
                {
                    var writingList = new List<string>();
                    controlNum++;
                    string[] writingData = { InputDailyReportForm.controlNum.ToString(), this.textBox1.Text, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text };
                    foreach (var writingEle in writingData) 
                    {
                        if(writingEle.Contains("\n")){
                            writingList.Add(writingEle.Replace("\r\n", "@NewLine"));
                        }else {
                            writingList.Add(writingEle);
                        }
                    }
                    streamWriter.WriteLine(writingList.Aggregate((i, j)=> i + "," + j));
                }
            }
        }

        private Boolean inputcheck(string inputDate, string inputImpContent, string inputScheContent, string inputTask) 
        {
            var pattern = @"\d{4}/\d{2}/\d{2}";

            if ((inputDate == null) || (inputDate.Length == 0))
            {
                MessageBox.Show("日付が入力されていません。");
                return false;
            }

            if (!Regex.IsMatch(inputDate, pattern) || isDate(inputDate)) 
            {
                MessageBox.Show("正しい形式で日付を入力してください。");
                return false;
            }

            if ((inputImpContent == null) || (inputImpContent.Length == 0)) 
            {
                MessageBox.Show("実施内容が入力されていません。");
                return false;
            }

            if ((inputScheContent == null) || (inputScheContent.Length == 0))
            {
                MessageBox.Show("翌日予定が入力されていません。");
                return false;
            }

            if ((inputTask == null) || (inputTask.Length == 0))
            {
                MessageBox.Show("課題が入力されていません。");
                return false;
            }

            return true;
        }

        private Boolean isDate(string dateStr) 
        {
            var pattern = '/';
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

    }
}
