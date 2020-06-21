using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReportWinApp
{
    class CsvWriter:IDisposable
    {
        private static readonly string _createDataFilePath = "C:/Users/Shota Tsuji/Desktop/DailyReportData.csv";
        private StreamWriter streamWriter = null;

        public CsvWriter() : this(Encoding.Default)
        {      
        }

        public CsvWriter(Encoding encoding) 
        {
            var fileStream = new FileStream(CsvWriter.CreateDataFilePath, FileMode.Append, FileAccess.Write);
            streamWriter = new StreamWriter(fileStream, encoding);
        }
        public static string CreateDataFilePath { get => _createDataFilePath; }
        public void WriteRow (string[] rowArray) 
        {
            var writingLine = string.Join(",", rowArray);
            streamWriter.WriteLine(writingLine);
        }

        public void Dispose()
        {
            if (this.streamWriter == null) 
            {
                return;
            }

            this.streamWriter.Close();
            this.streamWriter.Dispose();
            this.streamWriter = null;
        }
    }
}
