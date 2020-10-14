using CsvHelper.Configuration.Attributes;
using System;

namespace AutoReportWinApp
{

    /// <summary>
    /// 日報データエンティティクラス
    /// </summary>
    /// <remarks>日報データ用クラス</remarks>
    class DailyReportEntity
    {
        //csvファイルを読み込み、日報データを表示している
        //以下csvファイル用の改行文字列とする
        private static readonly string UserNewLineStr = "@NewLine";
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
        public static string ReplaceToNewLineStr(string userNewLineStr)
        {
            string rtnStr = userNewLineStr;
            if (userNewLineStr.Contains(UserNewLineStr))
            {
                rtnStr = userNewLineStr.Replace(UserNewLineStr, NewLineStr);
            }
            return rtnStr;
        }

        /// <summary>
        /// 改行を「@NewLine」に置換
        /// </summary>
        /// <param name="strWithNewLine">置換前文字列</param>
        /// <returns>置換後文字列</returns>
        public static string ReplaceToUserNewLineStr(string newLineStr)
        {
            string rtnStr = newLineStr;
            if (newLineStr.Contains(NewLineStr))
            {
                rtnStr = newLineStr.Replace(NewLineStr, UserNewLineStr);
            }
            return rtnStr;
        }
    }
}
