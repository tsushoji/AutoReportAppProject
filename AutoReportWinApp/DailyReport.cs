using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReportWinApp
{
    class DailyReport
    {
        private int _controlNum;
        private string _csvDailyReportLine;
        public int ControlNum { get => _controlNum; set => _controlNum = value; }
        public string CsvDailyReportLine { get => _csvDailyReportLine; set => _csvDailyReportLine = value; }
    }
}
