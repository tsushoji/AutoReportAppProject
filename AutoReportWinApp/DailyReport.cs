using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReportWinApp
{
    class DailyReport:IDisposable
    {
        public static int ControlNum;
        private DateTime _createDate;
        private string _implementationContent;
        private string _scheduledForNextDay;
        private string _task;

        public DateTime CreateDate { get => _createDate; set => _createDate = value; }

        public string ImplementationContent { get => _implementationContent; set => _implementationContent = value; }

        public string ScheduledForNextDay { get => _scheduledForNextDay; set => _scheduledForNextDay = value; }

        public string Task { get => _task; set => _task = value; }

        public DailyReport(DateTime createDate, string implementationContent, string scheduledForNextDay, string task) 
        {
            ControlNum++;
            this.CreateDate = createDate;
            this.ImplementationContent = implementationContent;
            this.ScheduledForNextDay = scheduledForNextDay;
            this.Task = task;
        }

        public void Dispose()
        {
            
        }
    }
}
