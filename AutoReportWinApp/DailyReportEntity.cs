using CsvHelper.Configuration.Attributes;

namespace AutoReportWinApp
{
    /// <summary>
    /// 日報データエンティティクラス
    /// </summary>
    /// <remarks>日報データ用クラス</remarks>
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

        /// <summary>
        /// 「@NewLine」を改行に置換
        /// </summary>
        /// <param name="strWithUserNewLine">置換前文字列</param>
        /// <returns>置換後文字列</returns>
        public static string replaceToStrWithNewLine(string strWithUserNewLine)
        {
            string rtnStr = strWithUserNewLine;
            if (strWithUserNewLine.Contains(SetValue.AppConstants.UserNewLineStr))
            {
                rtnStr = strWithUserNewLine.Replace(SetValue.AppConstants.UserNewLineStr, SpecialStr.AppConstants.NewLineStr);
            }
            return rtnStr;
        }

        /// <summary>
        /// 改行を「@NewLine」に置換
        /// </summary>
        /// <param name="strWithNewLine">置換前文字列</param>
        /// <returns>置換後文字列</returns>
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
