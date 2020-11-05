using CsvHelper;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Data;

namespace AutoReportWinApp
{
    /// <summary>
    /// 出力タイプ
    /// </summary>
    /// <remarks>DAILY_REPORT_DATA：日報データ、WEEKLY_REPORT：週報</remarks>
    public enum OutputType
    {
        DAILY_REPORT_DATA,
        WEEKLY_REPORT
    }

    /// <summary>
    /// データグリッドビューカラムインデックス番号
    /// </summary>
    /// <remarks>CONTROL_NUM：管理番号、DATE_STR：日付、IMPLEMENTATION_CONTENT：実施内容、TOMMOROW_PLAN：翌日予定、TASK：課題</remarks>
    public enum DataGridViewColumnIndex
    {
        CONTROL_NUM,
        DATE_STR,
        IMPLEMENTATION_CONTENT,
        TOMMOROW_PLAN,
        TASK
    }

    /// <summary>
    /// 日報データリストフォームクラス
    /// </summary>
    /// <remarks>作成した日報データを表示するフォーム</remarks>
    public partial class DailyReportDataListForm : Form
    {
        internal List<DailyReportEntity> dailyReportDataList;
        private string csvDailyReportDataPath;
        private int currentDailyReportDataIndex;
        private int pageCount;

        private static string tmpWeeklyReportStr = "【日付】" + Environment.NewLine +
                                                "{RepFstStr}" + Environment.NewLine +
                                                "【実施内容】" + Environment.NewLine +
                                                "{RepScdStr} " + Environment.NewLine +
                                                "【翌日予定】" + Environment.NewLine +
                                                "{RepThdStr} " + Environment.NewLine +
                                                "【課題】" + Environment.NewLine +
                                                "{RepFthStr}" + Environment.NewLine +
                                                Environment.NewLine;
        private static readonly int createReportDataFirstColNum = 1;
        private static readonly string csvDailyReportParentFolderName = @"\data";
        private static readonly string initDialogBoxFolderPath = @"c:\";
        private static readonly char delimiter = ',';
        private static readonly string dialogBoxCaption = "確認";
        private static readonly string dailyReportCsvFileNameWithExt = @"\daily_report_data.csv";
        private static readonly string weeklyReportTxtFileNameWithExt = @"\weekly_report.txt";
        private static readonly string winCharCode = "Shift_JIS";
        private static readonly string regExpTarMngNum = @"^([1-9]{1}[0-9]{0,}){1}(\,{1}[1-9]{1}[0-9]{0,}){0,4}?$";
        private static readonly string repFstStr = "{RepFstStr}";
        private static readonly string repScdStr = "{RepScdStr}";
        private static readonly string repThdStr = "{RepThdStr}";
        private static readonly string repFthStr = "{RepFthStr}";
        private static readonly int pageSize = 20;
        private static readonly char slashChar = '/';
        private static readonly string replaceErrMsgFirstArgStr = "{FIRSTARG}";
        private static readonly string readingPointStr = "、";

        public string CsvDailyReportDataPath { get => csvDailyReportDataPath; set => csvDailyReportDataPath = value; }
        public DataGridView DataGridView1 { get => this.dataGridView1; set => this.dataGridView1 = value; }
        public int CurrentDailyReportDataIndex { get => this.currentDailyReportDataIndex; set => this.currentDailyReportDataIndex = value; }
        public int PageCount { get => this.pageCount; set => this.pageCount = value; }
        public static string WinCharCode { get => winCharCode; }
        public static char SlashChar { get => slashChar; }
        public static string ReplaceErrMsgFirstArgStr { get => replaceErrMsgFirstArgStr; }
        public static string ReadingPointStr { get => readingPointStr; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <remarks>日報データリストフォームクラス</remarks>
        public DailyReportDataListForm()
        {
            InitializeComponent();
            CsvDailyReportDataPath = GetCsvDailyReportDataPath();
            this.dailyReportDataList = new List<DailyReportEntity>();
            InitDailyReportDataReader(CsvDailyReportDataPath);
        }

        /// <summary>
        /// 日報データcsvファイルパスを取得
        /// </summary>
        /// <returns>日報データcsvファイルパス</returns>
        private string GetCsvDailyReportDataPath()
        {
            Assembly myAssembly = Assembly.GetEntryAssembly();
            string exeFilePath = myAssembly.Location;
            string dirPath = System.IO.Path.GetDirectoryName(exeFilePath);
            string csvDailyReportDataPath = dirPath + csvDailyReportParentFolderName + dailyReportCsvFileNameWithExt;
            return csvDailyReportDataPath;
        }

        /// <summary>
        /// 日報データリストフォームクラスの初期処理
        /// </summary>
        /// <remarks>日報データcsvファイルを作成または読み込み、日報データリストフォームに表示</remarks>
        /// <param name="csvDailyReportDataPath">日報データcsvファイルパス</param>
        private void InitDailyReportDataReader(string csvDailyReportDataPath)
        {
            if (File.Exists(csvDailyReportDataPath))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = new StreamReader(csvDailyReportDataPath, Encoding.GetEncoding(WinCharCode)))
                using (var csv = new CsvReader(reader, CultureInfo.CurrentCulture))
                {
                    // ヘッダーの有無
                    csv.Configuration.HasHeaderRecord = false;
                    // データ読み出し（IEnumerable<Item>として受け取る）
                    this.dailyReportDataList = csv.GetRecords<DailyReportEntity>().ToList();
                    CurrentDailyReportDataIndex = 0;
                    SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
                }
            }
            else
            {
                string createDataDirectoryPath = System.IO.Path.GetDirectoryName(csvDailyReportDataPath);
                if (!System.IO.Directory.Exists(createDataDirectoryPath))
                {
                    Directory.CreateDirectory(createDataDirectoryPath);
                }
                File.Create(csvDailyReportDataPath).Close();
            }
        }

        /// <summary>
        /// ページングした日報データをデータグリッドビューにセット
        /// </summary>
        /// <param name="dailyReports">日報データリスト</param>
        internal void SetPagingDailyReportDataToDataGridView(List<DailyReportEntity> dailyReports)
        {
            var dailyReportList = new List<List<DailyReportEntity>>();
            var counter = 0;
            var dailyReportsByPageSize = new List<DailyReportEntity>();
            SetPageCountProperty(dailyReports);
            foreach (DailyReportEntity dailyReport in dailyReports)
            {
                dailyReportsByPageSize.Add(dailyReport);
                counter++;
                if (counter % pageSize == 0)
                {
                    dailyReportList.Add(dailyReportsByPageSize);
                    dailyReportsByPageSize = null;
                    dailyReportsByPageSize = new List<DailyReportEntity>();
                }
                if (counter % pageSize != 0 & counter == dailyReports.Count)
                {
                    dailyReportList.Add(dailyReportsByPageSize);
                }
            }
            if (dailyReportList.Count > 0)
            {
                DataGridView1.Rows.Clear();
                foreach (DailyReportEntity dailyReport in dailyReportList.ToArray()[CurrentDailyReportDataIndex])
                {
                    DataGridView1.Rows.Add(dailyReport.ControlNum, dailyReport.DateStr, DailyReportEntity.ReplaceToNewLineStr(dailyReport.ImplementationContent), DailyReportEntity.ReplaceToNewLineStr(dailyReport.TomorrowPlan), DailyReportEntity.ReplaceToNewLineStr(dailyReport.Task));
                }
            }
            this.bindingNavigatorCountItem.Text = SlashChar.ToString() + this.pageCount.ToString();
            this.bindingNavigatorPositionItem.Text = (CurrentDailyReportDataIndex + 1).ToString();
        }

        /// <summary>
        /// ページカウントプロパティーをセット
        /// </summary>
        /// <param name="dailyReports">日報データリスト</param>
        internal void SetPageCountProperty(List<DailyReportEntity> dailyReports)
        {
            var dailyReportsCount = dailyReports.Count;
            PageCount = 1;
            if (dailyReportsCount > 0)
            {
                PageCount = dailyReportsCount / pageSize;
                int calculateSurplus = dailyReportsCount % pageSize;
                if (calculateSurplus != 0)
                {
                    PageCount++;
                }
            }
        }

        /// <summary>
        /// データグリッドビューダブルクリック時、イベント
        /// </summary>
        /// <remarks>ダブルクリックした日報データを更新</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void ClickCreateData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex != 0)
            {
                using (var inputDailyReportForm = new InputDailyReportForm())
                {
                    inputDailyReportForm.DailyReportDataListForm = this;
                    inputDailyReportForm.CreateDataMode = CreateDataMode.UPDATE;
                    inputDailyReportForm.CreateDataControlNum = Int32.Parse(DataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                    inputDailyReportForm.TextBox1Text = DataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    inputDailyReportForm.TextBox2Text = DataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    inputDailyReportForm.TextBox3Text = DataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    inputDailyReportForm.TextBox4Text = DataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    inputDailyReportForm.ShowDialog();
                }
            }
        }

        /// <summary>
        ///「日報データ新規追加」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void ButtonAppendDailyReportData_Click(object sender, EventArgs e)
        {
            using (var inputDailyReportForm = new InputDailyReportForm())
            {
                inputDailyReportForm.DailyReportDataListForm = this;
                inputDailyReportForm.CreateDataMode = CreateDataMode.APPEND;
                if (this.dailyReportDataList.Count > 0)
                {
                    inputDailyReportForm.CreateDataControlNum = GetMaxControlNum(this.dailyReportDataList) + 1;
                }
                else
                {
                    inputDailyReportForm.CreateDataControlNum = createReportDataFirstColNum;
                }
                inputDailyReportForm.ShowDialog();
            }
        }

        /// <summary>
        /// 日報データリストの最大管理番号取得
        /// </summary>
        /// <remarks>日報データ新規作成時、必要</remarks>
        /// <param name="dailyReports">日報データリスト</param>
        /// <returns>最大管理番号</returns>
        private int GetMaxControlNum(List<DailyReportEntity> dailyReports)
        {
            var controlNumList = new List<int>();
            foreach (var dailyReport in dailyReports)
            {
                controlNumList.Add(Int32.Parse(dailyReport.ControlNum));
            }
            return controlNumList.Max();
        }

        /// <summary>
        /// 日報データリストクラス読み込み時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void DailyReportDataListForm_Load(object sender, EventArgs e)
        {
            ClearDataGridViewFocus();
        }

        /// <summary>
        /// データグリッドビューフォーカスをクリア
        /// </summary>
        private void ClearDataGridViewFocus()
        {
            this.Activate();
            this.dataGridView1.CurrentCell = null;
            this.dataGridView1.RowHeadersVisible = false;
        }

        /// <summary>
        /// データグリッドビューのスタイルを変更
        /// </summary>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        /// <param name="color">ARGB色</param>
        /// <param name="cursor">マウスポインターに使用されるイメージ</param>
        private void ChangeDataGridViewStyle(DataGridViewCellEventArgs e, Color color, Cursor cursor)
        {
            DataGridView1.Cursor = cursor;
            DataGridView1.CurrentCell = null;
            var dataGridViewRowStyle = DataGridView1.Rows[e.RowIndex].DefaultCellStyle;
            dataGridViewRowStyle.BackColor = color;
            dataGridViewRowStyle.SelectionBackColor = color;
        }

        /// <summary>
        /// 日報データ出力項目「フォルダダイアログ」ボタンクリック時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonFirstFolderDialog_Click(object sender, EventArgs e)
        {
            SetSelectedDialogBoxFolderPathToTextBox(this.textBox1);
        }

        /// <summary>
        /// 週報出力項目「フォルダダイアログ」ボタンクリック時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonSecondFolderDialog_Click(object sender, EventArgs e)
        {
            SetSelectedDialogBoxFolderPathToTextBox(this.textBox2);
        }

        /// <summary>
        /// フォルダダイアログボックスで選択されたフォルダパスを「テキストボックス」にセット
        /// </summary>
        /// <remarks>フォルダダイアログボックスでテキストボックスにセットする共通メソッド</remarks>
        /// <param name="textBox">テキストボックス</param>
        private void SetSelectedDialogBoxFolderPathToTextBox(TextBox textBox)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.SelectedPath = initDialogBoxFolderPath;
                dialog.Description = Properties.Resources.I0006;
                dialog.ShowNewFolderButton = true;

                DialogResult result = dialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    textBox.Text = dialog.SelectedPath;
                }
            }
        }

        /// <summary>
        /// 「出力」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonOutputDailyReportData_Click(object sender, EventArgs e)
        {
            OutputReportFile(OutputType.DAILY_REPORT_DATA);
        }

        /// <summary>
        /// 「週報出力」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonOutputWeeklyReport_Click(object sender, EventArgs e)
        {
            OutputReportFile(OutputType.WEEKLY_REPORT);
        }

        /// <summary>
        /// 日報データまたは週報を出力
        /// </summary>
        /// <remarks>テキストボックスにセットしたフォルダパスに日報データ（csvファイル）または週報（テキストファイル）を出力する共通メソッド</remarks>
        /// <param name="outputType">出力タイプ</param>
        private void OutputReportFile(OutputType outputType)
        {
            if (!Inputcheck(outputType))
            {
                switch (outputType)
                {
                    case OutputType.DAILY_REPORT_DATA:
                        OutputDailyReportData();
                        break;

                    case OutputType.WEEKLY_REPORT:
                        if (!ValidateWeeklyReport())
                        {
                            OutputWeeklyReport();
                        }
                        break;

                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// 日報データ出力
        /// </summary>
        /// <remarks>日報データ（csvファイル）を出力</remarks>
        private void OutputDailyReportData()
        {
            if (!System.IO.Directory.Exists(this.textBox1.Text))
            {
                MessageBox.Show(Properties.Resources.E0004);
                this.textBox2.ResetText();
                return;
            }
            var appendFlg = false;
            string outputPath = this.textBox1.Text + dailyReportCsvFileNameWithExt;
            string outputDailyReportDataCompMsg = Properties.Resources.I0003;
            if (File.Exists(outputPath))
            {
                outputDailyReportDataCompMsg = Properties.Resources.I0004;
                //上書きコピー
                appendFlg = true;
            }
            File.Copy(CsvDailyReportDataPath, outputPath, appendFlg);
            MessageBox.Show(outputDailyReportDataCompMsg);
        }

        /// <summary>
        /// 週報出力
        /// </summary>
        /// <remarks>週報（テキストファイル）を出力</remarks>
        private void OutputWeeklyReport()
        {
            if (!System.IO.Directory.Exists(this.textBox2.Text))
            {
                MessageBox.Show(Properties.Resources.E0004);
                this.textBox2.ResetText();
                return;
            }
            //上書きする
            var appendFlg = false;
            string outputPath = this.textBox2.Text + weeklyReportTxtFileNameWithExt;
            string outputWeeklyReportCompMsg = Properties.Resources.I0007;
            if (File.Exists(outputPath))
            {
                DialogResult dialogResult = MessageBox.Show(Properties.Resources.W0001, dialogBoxCaption, MessageBoxButtons.YesNo);
                if (dialogResult == System.Windows.Forms.DialogResult.No)
                {
                    MessageBox.Show(Properties.Resources.I0009);
                    //追記する
                    appendFlg = true;
                }
                else if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                {
                    outputWeeklyReportCompMsg = Properties.Resources.I0008;
                }
                else
                {
                    outputWeeklyReportCompMsg = Properties.Resources.I0005;
                    return;
                }
            }
            //出力タイプが週報の場合、書き込み
            using (var writer = new StreamWriter(outputPath, appendFlg, Encoding.GetEncoding(WinCharCode)))
            {
                string[] controlNumArray = this.textBox3.Text.Split(delimiter);
                var outputWeeklyReportDataList = new List<DailyReportEntity>();
                foreach (var dailyReportData in this.dailyReportDataList)
                {
                    if (controlNumArray.Contains(dailyReportData.ControlNum))
                    {
                        var reportData = new DailyReportEntity();
                        reportData.ControlNum = dailyReportData.ControlNum;
                        reportData.DateStr = dailyReportData.DateStr;
                        reportData.ImplementationContent = DailyReportEntity.ReplaceToNewLineStr(dailyReportData.ImplementationContent);
                        reportData.TomorrowPlan = DailyReportEntity.ReplaceToNewLineStr(dailyReportData.Task);
                        reportData.Task = DailyReportEntity.ReplaceToNewLineStr(dailyReportData.TomorrowPlan);
                        outputWeeklyReportDataList.Add(reportData);
                    }
                }
                //日付で昇順に並べる
                outputWeeklyReportDataList = outputWeeklyReportDataList.OrderBy(value => DateTime.Parse(value.DateStr)).ToList();
                //週報文字列生成
                string outputWeeklyReportByDateStr;
                var outputWeeklyReportStr = new StringBuilder();
                foreach (var outputWeeklyReportData in outputWeeklyReportDataList)
                {
                    outputWeeklyReportByDateStr = tmpWeeklyReportStr.Replace(repFstStr, outputWeeklyReportData.DateStr);
                    outputWeeklyReportByDateStr = outputWeeklyReportByDateStr.Replace(repScdStr, outputWeeklyReportData.ImplementationContent);
                    outputWeeklyReportByDateStr = outputWeeklyReportByDateStr.Replace(repThdStr, outputWeeklyReportData.TomorrowPlan);
                    outputWeeklyReportByDateStr = outputWeeklyReportByDateStr.Replace(repFthStr, outputWeeklyReportData.Task);
                    outputWeeklyReportStr.Append(outputWeeklyReportByDateStr);
                }
                writer.Write(outputWeeklyReportStr.ToString());
            }
            MessageBox.Show(outputWeeklyReportCompMsg);
        }

        /// <summary>
        /// 日報データまたは週報出力時、入力チェック
        /// </summary>
        /// <remarks>メッセージボックスにエラーメッセージ出力</remarks>
        /// <param name="outputType">出力タイプ</param>
        /// <returns>判定結果</returns>
        private Boolean Inputcheck(OutputType outputType)
        {
            var errMsgEleList = new List<string>();
            var errMsg = new StringBuilder();
            if (outputType == OutputType.DAILY_REPORT_DATA)
            {
                if (string.IsNullOrEmpty(this.textBox1.Text))
                {
                    string outputDailyReportDataFolderPathStrErrMsgEle = this.label1.Text + this.label2.Text.Substring(0, 8);
                    errMsgEleList.Add(outputDailyReportDataFolderPathStrErrMsgEle);
                }
            }
            if (outputType == OutputType.WEEKLY_REPORT)
            {
                if (string.IsNullOrEmpty(this.textBox2.Text))
                {
                    string outputWeeklyReportFolderPathStrErrMsgEle = this.label4.Text + this.label3.Text.Substring(0, 8);
                    errMsgEleList.Add(outputWeeklyReportFolderPathStrErrMsgEle);
                }

                if (string.IsNullOrEmpty(this.textBox3.Text))
                {
                    string tarMngNumStrErrMsgEle = this.label4.Text + this.label5.Text.Substring(0, 6);
                    errMsgEleList.Add(tarMngNumStrErrMsgEle);
                }
            }
            if (errMsgEleList.Count > 0)
            {
                string partialErrMsg = errMsgEleList.Aggregate((i, j) => i + ReadingPointStr + j);
                errMsg.Append(Properties.Resources.E0001.Replace(ReplaceErrMsgFirstArgStr, partialErrMsg));
            }
            if (errMsg.Length > 0)
            {
                MessageBox.Show(errMsg.ToString());
            }
            return errMsg.Length > 0;
        }

        /// <summary>
        /// バリデーションチェック
        /// </summary>
        /// <remarks>週報出力時、バリデーションチェック</remarks>
        /// <returns>判定結果</returns>
        private Boolean ValidateWeeklyReport()
        {
            var errMsg = new StringBuilder();
            var pattern = regExpTarMngNum;
            if (!Regex.IsMatch(this.textBox3.Text, pattern))
            {
                errMsg.Append(Properties.Resources.E0002);
            }
            //形式が正しくない場合、チェックしない
            if (!errMsg.ToString().Contains(Properties.Resources.E0002))
            {
                if (!DuplicateArrayCheck(this.textBox3.Text.Split(delimiter)))
                {
                    if (errMsg.Length > 0)
                    {
                        errMsg.Append(DailyReportEntity.NewLineStr);
                    }
                    errMsg.Append(Properties.Resources.E0005);
                }

                if (!ContainDailyReportDataDataCheck(this.textBox3.Text.Split(delimiter)))
                {
                    if (errMsg.Length > 0)
                    {
                        errMsg.Append(DailyReportEntity.NewLineStr);
                    }
                    errMsg.Append(Properties.Resources.E0006);
                }
            }

            if (errMsg.Length > 0)
            {
                MessageBox.Show(errMsg.ToString());
            }
            return errMsg.Length > 0;
        }

        /// <summary>
        /// 配列要素重複チェック
        /// </summary>
        /// <remarks>配列要素の重複をチェック</remarks>
        /// <param name="array">配列</param>
        /// <returns>判定結果</returns>
        private Boolean DuplicateArrayCheck(string[] array)
        {
            for (int i = 0; i < array.Length - 1; i++)
            {
                for (int j = i + 1; j < array.Length; j++)
                {
                    if (array[i].Equals(array[j]))
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// 配列要素存在チェック
        /// </summary>
        /// <remarks>選択された対象管理番号がデータリストに存在しているかチェック</remarks>
        /// <param name="array">配列</param>
        /// <returns>判定結果</returns>
        private Boolean ContainDailyReportDataDataCheck(string[] array)
        {
            var dataControlNumList = new List<String>();
            foreach (var dailyReportData in this.dailyReportDataList)
            {
                dataControlNumList.Add(dailyReportData.ControlNum);
            }
            foreach (var element in array)
            {
                if (dataControlNumList.Contains(element))
                {
                    continue;
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// データグリッドビューマウスポインターがセルに入ったときのイベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void DailyReportDataListFormDataGridView_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex != 0)
            {
                ChangeDataGridViewStyle(e, Color.Yellow, Cursors.Hand);
            }
        }

        /// <summary>
        /// データグリッドビューマウスポインターがセルから離れるときのイベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void DailyReportDataListFormDataGridView_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex != 0)
            {
                ChangeDataGridViewStyle(e, Color.Empty, Cursors.Default);
            }
        }

        /// <summary>
        ///「最初へ」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void ButtonMoveFirstItem_Click(object sender, EventArgs e)
        {
            CurrentDailyReportDataIndex = 0;
            SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
            ClearDataGridViewFocus();
        }

        /// <summary>
        ///「最後へ」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void ButtonMoveLastItem_Click(object sender, EventArgs e)
        {
            if (PageCount > 1)
            {
                CurrentDailyReportDataIndex = PageCount - 1;
            }
            SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
            ClearDataGridViewFocus();
        }

        /// <summary>
        ///「次へ」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void ButtonMoveNextItemItem_Click(object sender, EventArgs e)
        {
            if (CurrentDailyReportDataIndex < PageCount - 1)
            {
                CurrentDailyReportDataIndex++;
            }
            SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
            ClearDataGridViewFocus();
        }

        /// <summary>
        ///「前へ」ボタン押下時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void ButtonMovePreviousItemItem_Click(object sender, EventArgs e)
        {
            if (CurrentDailyReportDataIndex > 0)
            {
                CurrentDailyReportDataIndex--;
            }
            SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
            ClearDataGridViewFocus();
        }

        /// <summary>
        ///データグリッドビューヘッダークリック時、イベント
        /// </summary>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void ChangeOrder_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (DataGridView1.Rows.Count > 0)
            {
                switch (e.ColumnIndex)
                {
                    case (int)DataGridViewColumnIndex.CONTROL_NUM:
                        if (DataGridView1.Columns[0].HeaderCell.SortGlyphDirection == SortOrder.Ascending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderBy(value => Int32.Parse(value.ControlNum)).ToList();
                        }
                        else if (DataGridView1.Columns[0].HeaderCell.SortGlyphDirection == SortOrder.Descending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderByDescending(value => Int32.Parse(value.ControlNum)).ToList();
                        }
                        else
                        {
                        }
                        SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
                        break;

                    case (int)DataGridViewColumnIndex.DATE_STR:
                        if (DataGridView1.Columns[1].HeaderCell.SortGlyphDirection == SortOrder.Ascending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderBy(value => DateTime.Parse(value.DateStr)).ToList();
                        }
                        else if (DataGridView1.Columns[1].HeaderCell.SortGlyphDirection == SortOrder.Descending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderByDescending(value => DateTime.Parse(value.DateStr)).ToList();
                        }
                        else
                        {
                        }
                        SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
                        break;

                    case (int)DataGridViewColumnIndex.IMPLEMENTATION_CONTENT:
                        if (DataGridView1.Columns[2].HeaderCell.SortGlyphDirection == SortOrder.Ascending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderBy(value => value.ImplementationContent).ToList();
                        }
                        else if (DataGridView1.Columns[2].HeaderCell.SortGlyphDirection == SortOrder.Descending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderByDescending(value => value.ImplementationContent).ToList();
                        }
                        else
                        {
                        }
                        SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
                        break;

                    case (int)DataGridViewColumnIndex.TOMMOROW_PLAN:
                        if (DataGridView1.Columns[3].HeaderCell.SortGlyphDirection == SortOrder.Ascending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderBy(value => value.TomorrowPlan).ToList();
                        }
                        else if (DataGridView1.Columns[3].HeaderCell.SortGlyphDirection == SortOrder.Descending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderByDescending(value => value.TomorrowPlan).ToList();
                        }
                        else
                        {
                        }
                        SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
                        break;

                    case (int)DataGridViewColumnIndex.TASK:
                        if (DataGridView1.Columns[4].HeaderCell.SortGlyphDirection == SortOrder.Ascending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderBy(value => value.Task).ToList();
                        }
                        else if (DataGridView1.Columns[4].HeaderCell.SortGlyphDirection == SortOrder.Descending)
                        {
                            this.dailyReportDataList = this.dailyReportDataList.OrderByDescending(value => value.Task).ToList();
                        }
                        else
                        {
                        }
                        SetPagingDailyReportDataToDataGridView(this.dailyReportDataList);
                        break;

                    default:
                        break;
                }
            }
        }
    }
}
