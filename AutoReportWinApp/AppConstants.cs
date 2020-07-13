using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReportWinApp
{
    class AppConstants
    {
        public static readonly char DailyReportDataLineSeparateChar = ',';
        public static readonly char DateFormatCheckSeparate = '/';
        public static readonly int CreateReportDataColNumFirst = 1;
        public static readonly string CsvDailyReportDataPathEnd = @"\data\daily_report_data.csv";
        public static readonly string DailyReportDataLineSeparateStr = ",";
        public static readonly string DateFormat = "yyyy/MM/dd";
        public static readonly string DateRegExp = @"^\d{4}/\d{2}/\d{2}$";
        public static readonly string NewLineStr = "\r\n";
        public static readonly string UserNewLineStr = "@NewLine";
        public static readonly string NotInputFilePathMsg = "csvデータファイルパスが入力されていません。";
        public static readonly string NotInputCheckItemMsgSeparate = "、";
        public static readonly string NotInputDateFormatMsg = "正しい形式で日付を入力してください。";
        public static readonly string NotInputCheckItemMsgEnd = "が入力されていません";
        public static readonly string DuplicateDailyReportDataMsg = "日報データが重複しています。";
        public static readonly string CreateDataAppendSeparate = ",";
        public static readonly string CreateDataAppendCmpMsg = "日報データを追加しました。";
        public static readonly string CreateDataUpdateCmpMsg = "日報データを更新しました。";
        public static readonly string InputCalenderDataPassFormItemNameStr = "textBox1";
        public static readonly string FolderDialogBoxInitSelected = @"c:\temp\";
        public static readonly string AppendOutputFileDialogBoxMsg = "同じファイルパスが存在します。上書きしてもよろしいでしょうか。";
        public static readonly string AppendOutputFileDialogBoxTitle = "確認";
        public static readonly string OutputFilePathEnd = @"\daily_report_data.csv";
        public static readonly string FolderDialogBoxExplation = "フォルダを選択してください。";
        public static readonly string OutputDailyReportDataCompMsg = "日報データを出力いたしました。";
        public static readonly string OutputDailyReportDataAppendCompMsg = "日報データを上書きいたしました。";
        public static readonly string NotOutputFolderPathMsg = "入力したフォルダパスは存在しません。";
        public static readonly string CancelMsg = "キャンセルしました。";
    }
}
