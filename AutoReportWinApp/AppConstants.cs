using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReportWinApp
{
    class AppConstants
    {
        public static readonly char DateFormatCheckSeparate = '/';
        public static readonly int CreateDataFirstColNum = 1;
        public static readonly string DateFormat = "yyyy/MM/dd";
        public static readonly string DateRegExp = @"\d{4}/\d{2}/\d{2}";
        public static readonly string NewLineStr = "\r\n";
        public static readonly string UserNewLineStr = "@NewLine";
        public static readonly string NotInputFilePathMsg = "csvデータファイルパスが入力されていません。";
        public static readonly string NotInputDateMsg = "日付が入力されていません。";
        public static readonly string NotInputDateFormatMsg = "正しい形式で日付を入力してください。";
        public static readonly string NotInputImpContentMsg = "実施内容が入力されていません。";
        public static readonly string NotInputSchContentMsg = "翌日予定が入力されていません。";
        public static readonly string NotInputTaskMsg = "課題が入力されていません。";
        public static readonly string CreateDataAppendCmpMsg = "日報データを追加しました。";
        public static readonly string CreateDataUpdateCmpMsg = "日報データを更新しました。";
    }
}
