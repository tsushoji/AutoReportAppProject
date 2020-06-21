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
        public InputDailyReportForm()
        {
            InitializeComponent();
        }

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
            using (var dailyReport = new DailyReport(DateTime.Parse(this.textBox1.Text), this.textBox2.Text, this.textBox3.Text, this.textBox4.Text))
            using (var csvWriter = new CsvWriter())
            {
                string[] writingData = { DailyReport.ControlNum.ToString(), dailyReport.CreateDate.ToString(), dailyReport.ImplementationContent, dailyReport.ScheduledForNextDay, dailyReport.Task };
                csvWriter.WriteRow(writingData);
            }
        }
    }
}
