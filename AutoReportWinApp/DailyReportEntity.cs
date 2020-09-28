using CsvHelper.Configuration.Attributes;

namespace AutoReportWinApp
{
    class DailyReportEntity
    {
        [Index(0)]
        public string controlNum { get; set; }
        [Index(1)]
        public string date { get; set; }
        [Index(2)]
        public string impContent { get; set; }
        [Index(3)]
        public string schContent { get; set; }
        [Index(4)]
        public string task { get; set; }
        public string getLineAtConma()
        {
            string[] properties = { controlNum, date, impContent, schContent, task };
            string rtnLine = string.Join(SpecialStr.AppConstants.CommaStr, properties);
            return rtnLine;
        }
        public static string replaceToStrWithNewLine(string strWithUserNewLine)
        {
            string rtnStr = strWithUserNewLine;
            if (strWithUserNewLine.Contains(SetValue.AppConstants.UserNewLineStr))
            {
                rtnStr = strWithUserNewLine.Replace(SetValue.AppConstants.UserNewLineStr, SpecialStr.AppConstants.NewLineStr);
            }
            return rtnStr;
        }
        public static string replaceToStrWithUserNewLine(string strWithNewLine)
        {
            string rtnStr = strWithNewLine;
            if (strWithNewLine.Contains(SpecialStr.AppConstants.NewLineStr))
            {
                rtnStr = strWithNewLine.Replace(SpecialStr.AppConstants.NewLineStr, SetValue.AppConstants.UserNewLineStr);
            }
            return rtnStr;
        }
    }
}
