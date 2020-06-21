﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoReportWinApp
{
    public partial class CalendarForm : Form
    {
        private InputDailyReportForm inputDailyReportForm;
        public CalendarForm(InputDailyReportForm inputDailyReportForm)
        {
            this.inputDailyReportForm = inputDailyReportForm;
            InitializeComponent();
        }

        private void mousePointer_DataSelected(object sender, DateRangeEventArgs e)
        {
            TextBox inputDailyReportForm_textBox1 = (TextBox)inputDailyReportForm.Controls["textBox1"];
            inputDailyReportForm_textBox1.Text = this.monthCalendar1.SelectionRange.Start.Date.ToString("yyyy/MM/dd");
        }
    }
}
