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
            if (inputcheck(this.textBox1.Text, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text)) 
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
            var pattern = "";
            var returnFlg = true;

            if ((inputDate == null) || (inputDate.Length == 0))
            {
                MessageBox.Show("日付が入力されていません。");
                returnFlg = false;
            }

            if (!Regex.IsMatch(inputDate, pattern)) 
            {
                MessageBox.Show("正しい形式で日付を入力してください。");
                returnFlg = false;
            }

            if ((inputImpContent == null) || (inputImpContent.Length == 0)) 
            {
                MessageBox.Show("実施内容が入力されていません。");
                returnFlg = false;
            }

            if ((inputScheContent == null) || (inputScheContent.Length == 0))
            {
                MessageBox.Show("翌日予定が入力されていません。");
                returnFlg = false;
            }

            if ((inputTask == null) || (inputTask.Length == 0))
            {
                MessageBox.Show("課題が入力されていません。");
                returnFlg = false;
            }

            return returnFlg;
        }

    }
}
