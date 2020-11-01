using CsvHelper.Configuration.Attributes;
using System;

namespace AutoReportWinApp
{

    /// <summary>
    /// 日報データエンティティクラス
    /// </summary>
    /// <remarks>日報データクラス</remarks>
    class DailyReportEntity
    {
        private static readonly string userNewLineStr = "@NewLine";
        private static readonly string newLineStr = "\r\n";

        [Index(0)]
        public string ControlNum { get; set; }
        [Index(1)]
        public string DateStr { get; set; }
        [Index(2)]
        public string ImplementationContent { get; set; }
        [Index(3)]
        public string TomorrowPlan { get; set; }
        [Index(4)]
        public string Task { get; set; }

        public static string NewLineStr { get => newLineStr; }

        /// <summary>
        /// 「@NewLine」を改行に置換
        /// </summary>
        /// <param name="strWithUserNewLine">置換前文字列</param>
        /// <returns>置換後文字列</returns>
        public static string ReplaceToNewLineStr(string strWithUserNewLine)
        {
            var rtnStr = strWithUserNewLine;
            if (rtnStr.Contains(userNewLineStr))
            {
                rtnStr = rtnStr.Replace(userNewLineStr, NewLineStr);
            }
            return rtnStr;
        }

        /// <summary>
        /// 改行を「@NewLine」に置換
        /// </summary>
        /// <param name="strWithNewLine">置換前文字列</param>
        /// <returns>置換後文字列</returns>
        public static string ReplaceToUserNewLineStr(string strWithNewLine)
        {
            var rtnStr = strWithNewLine;
            if (rtnStr.Contains(NewLineStr))
            {
                rtnStr = rtnStr.Replace(NewLineStr, userNewLineStr);
            }
            return rtnStr;
        }
    }
}
