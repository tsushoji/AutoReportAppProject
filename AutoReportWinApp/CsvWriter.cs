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
        private StreamWriter streamWriter = null;

        public CsvWriter(string path) : this(path, Encoding.Default)
        {      
        }

        public CsvWriter(string path, Encoding encoding) 
        {
            var fileStream = new FileStream(path, FileMode.Append, FileAccess.Write);
            streamWriter = new StreamWriter(fileStream, encoding);
        }
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
