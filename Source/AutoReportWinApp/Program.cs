using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoReportWinApp
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメインエントリポイントクラス
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //日報データリストフォーム初期表示
            Application.Run(new DailyReportDataListForm());
        }
    }
}
