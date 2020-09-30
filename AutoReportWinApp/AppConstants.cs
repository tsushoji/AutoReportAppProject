using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReportWinApp
{
    namespace SetValue
    {
        /// <summary>
        /// 定数ユーティリティークラス
        /// </summary>
        /// <remarks>アプリケーション用</remarks>
        class AppConstants
        {
            public static readonly int CreateReportDataColNumFirst = 1;
            public static readonly string CsvDailyReportDataPathEnd = @"\data\daily_report_data.csv";
            public static readonly string UserNewLineStr = "@NewLine";
            public static readonly string InputCalenderDataPassFormItemNameStr = "textBox1";
            public static readonly string FolderDialogBoxInitSelected = @"c:\temp\";
            public static readonly string AppendOutputFileDialogBoxTitle = "確認";
            public static readonly string OutputFilePathEnd = @"\daily_report_data.csv";
            public static readonly string AscResName = ".Resource";
            public static readonly string WinEncoding = "Shift_JIS";
        }
    }
    namespace RegExp
    {
        /// <summary>
        /// 定数ユーティリティークラス
        /// </summary>
        /// <remarks>正規表現用</remarks>
        class AppConstants
        {
            public static readonly string DateRegExp = @"^\d{4}/\d{2}/\d{2}$";
        }
    }
    namespace SpecialStr
    {
        /// <summary>
        /// 定数ユーティリティークラス
        /// </summary>
        /// <remarks>特殊文字用</remarks>
        class AppConstants
        {
            public static readonly char CommaChar = ',';
            public static readonly char SlashChar = '/';
            public static readonly string CommaStr = ",";
            public static readonly string ReadingPointStr = "、";
            public static readonly string NewLineStr = "\r\n";
            public static readonly string DateFormat = "yyyy/MM/dd";
        }
    }
}
