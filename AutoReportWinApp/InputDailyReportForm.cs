using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
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
            using (var fileStream = new FileStream(InputDailyReportForm.CreateDataFilePath, FileMode.Append, FileAccess.Write))
            using (var streamWriter = new StreamWriter(fileStream, Encoding.Default))
            {
                controlNum++;
                string[] writingData = { InputDailyReportForm.controlNum.ToString(), this.textBox1.Text, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text };
                var writingLine = string.Join(",", writingData);
                streamWriter.WriteLine(writingLine);
            }
        }
    }
}
