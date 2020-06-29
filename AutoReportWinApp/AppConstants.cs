using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReportWinApp
{
    class AppConstants
    {
        public static readonly char DailyReportDataLineSeparate = ',';
        public static readonly char DateFormatCheckSeparate = '/';
        public static readonly int CreateDataFirstColNum = 1;
        public static readonly string DateFormat = "yyyy/MM/dd";
        public static readonly string DateRegExp = @"^\d{4}/\d{2}/\d{2}$";
        public static readonly string NewLineStr = "\r\n";
        public static readonly string UserNewLineStr = "@NewLine";
        public static readonly string NotInputFilePathMsg = "csvデータファイルパスが入力されていません。";
        public static readonly string NotInputCheckItemMsgSeparate = "、";
        public static readonly string NotInputDateMsgEle = "日付";
        public static readonly string NotInputDateFormatMsg = "正しい形式で日付を入力してください。";
        public static readonly string NotInputImpContentMsgEle = "実施内容";
        public static readonly string NotInputSchContentMsgEle = "翌日予定";     
        public static readonly string NotInputTaskMsgEle = "課題";
        public static readonly string NotInputCheckItemMsgEnd = "が入力されていません";
        public static readonly string DuplicateDailyReportDataMsg = "日報データが重複しています。";
        public static readonly string CreateDataAppendSeparate = ",";
        public static readonly string CreateDataAppendCmpMsg = "日報データを追加しました。";
        public static readonly string CreateDataUpdateCmpMsg = "日報データを更新しました。";
    }
}
