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

namespace AutoReportWinApp
{
    /// <summary>
    /// 出力タイプ
    /// </summary>
    /// <remarks>DAILYREPORTDATA：日報データ、WEEKLYREPORT：週報</remarks>
    public enum OutputType
    {
        DAILYREPORTDATA,
        WEEKLYREPORT
    }

    /// <summary>
    /// 日報データリストフォームクラス
    /// </summary>
    /// <remarks>作成した日報データを表示するフォーム</remarks>
    public partial class DailyReportDataListForm : Form
    {
        internal Dictionary<int, DailyReportEntity> CsvDailyReportDataMap;
        private string csvDailyReportDataPath;

        private static string tmpWeeklyReportStr = "【日付】" + Environment.NewLine + "{RepFstStr}" + Environment.NewLine + "【実施内容】{RepScdStr}【翌日予定】{RepThdStr}【課題】{RepFthStr}";
        private static readonly int createReportDataFirstColNum = 1;
        private static readonly string csvDailyReportParentFolderName = @"\data";
        private static readonly string initDialogBoxFolderPath = @"c:\";
        private static readonly string dialogBoxCaption = "確認";
        private static readonly string dailyReportCsvFileNameWithExt = @"\daily_report_data.csv";
        private static readonly string weeklyReportTxtFileNameWithExt = @"\weekly_report.txt";
        private static readonly string winCharCode = "Shift_JIS";
        private static readonly string RegExpTarMngNum = @"^([1-9]{1}[0-9]{0,}){1}(\,{1}[1-9]{1}[0-9]{0,}){0,4}?$";

        public string CsvDailyReportDataPath { get => csvDailyReportDataPath; set => csvDailyReportDataPath = value; }
        public DataGridView DataGridView1 { get => this.dataGridView1; set => this.dataGridView1 = value; }
        public static string WinCharCode { get => winCharCode; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <remarks>日報データリストフォームクラス</remarks>
        public DailyReportDataListForm()
        {
            InitializeComponent();
            this.CsvDailyReportDataMap = new Dictionary<int, DailyReportEntity>();
            CsvDailyReportDataPath = this.GetCsvDailyReportDataPath();
            this.InitDailyReportDataReader(CsvDailyReportDataPath);
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
        /// <param name="createDataFilePath">日報データcsvファイルパス</param>
        private void InitDailyReportDataReader(string csvDailyReportDataPath)
        {
            //日報データが既にある場合
            if (File.Exists(csvDailyReportDataPath))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = new StreamReader(csvDailyReportDataPath, Encoding.GetEncoding(WinCharCode)))
                using (var csv = new CsvReader(reader, CultureInfo.CurrentCulture))
                {
                    int rowIndex = 0;
                    // ヘッダーの有無
                    csv.Configuration.HasHeaderRecord = false;
                    // データ読み出し（IEnumerable<Item>として受け取る）
                    var dailyReports = csv.GetRecords<DailyReportEntity>();
                    foreach (DailyReportEntity dailyReport in dailyReports)
                    {
                        DataGridView1.Rows.Add(dailyReport.ControlNum, dailyReport.DateStr, DailyReportEntity.ReplaceToNewLineStr(dailyReport.ImplementationContent), DailyReportEntity.ReplaceToNewLineStr(dailyReport.TomorrowPlan), DailyReportEntity.ReplaceToNewLineStr(dailyReport.Task));
                        this.CsvDailyReportDataMap.Add(rowIndex, dailyReport);
                        rowIndex++;
                    }
                }
            }
            //日報データがない場合
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
        /// データグリッドビューダブルクリック時、イベント
        /// </summary>
        /// <remarks>更新した日報データをダブルクリック時、イベント</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void ClickCreateData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex != 0)
            {
                using (var inputDailyReportForm = new InputDailyReportForm())
                {
                    inputDailyReportForm.DailyReportDataListForm = this;
                    inputDailyReportForm.CreateDataMode = CreateDataMode.APPEND;
                    //日報データが1件以上ある場合
                    if (this.CsvDailyReportDataMap.Count > 0)
                    {
                        //日報データ作成管理番号をプロパティにセット
                        inputDailyReportForm.CreateDataControlNum = this.GetMaxControlNum(this.CsvDailyReportDataMap) + 1;
                    }
                    //日報データがない場合
                    else
                    {
                        //日報データ作成管理番号をプロパティにセット
                        inputDailyReportForm.CreateDataControlNum = createReportDataFirstColNum;
                    }

                    //日報データを更新する場合
                    if (e.RowIndex != DataGridView1.Rows.Count - 1)
                    {
                        inputDailyReportForm.CreateDataMode = CreateDataMode.UPDATE;
                        inputDailyReportForm.CreateDataControlNum = Int32.Parse(DataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                        inputDailyReportForm.TextBox1Text = DataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                        inputDailyReportForm.TextBox2Text = DataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                        inputDailyReportForm.TextBox3Text = DataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                        inputDailyReportForm.TextBox4Text = DataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    }

                    inputDailyReportForm.ShowDialog();
                }
            }
        }

        /// <summary>
        /// 日報データリストの最大管理番号取得
        /// </summary>
        /// <remarks>日報データ新規作成時、必要</remarks>
        /// <param name="dataReportMap">日報データ辞書</param>
        /// <returns>最大管理番号</returns>
        private int GetMaxControlNum(Dictionary<int, DailyReportEntity> dataReportMap)
        {
            var controlNumList = new List<int>();
            int controlNum = 0;
            foreach (KeyValuePair<int, DailyReportEntity> keyValuePair in dataReportMap)
            {
                controlNum = Int32.Parse(keyValuePair.Value.ControlNum);
                controlNumList.Add(controlNum);
            }
            return controlNumList.Max();
        }

        /// <summary>
        /// 日報データリストクラス読み込み時、イベント
        /// </summary>
        /// <remarks>クラス読み込み時、データグリッドビューフォーカスをクリア</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void DailyReportDataListForm_Load(object sender, EventArgs e)
        {
            //データグリッドビューフォーカスをクリア
            this.Activate();
            this.dataGridView1.CurrentCell = null;
            this.dataGridView1.RowHeadersVisible = false;
            //データグリッドビュー列幅調整
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
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
        /// <remarks>日報データ出力項目「テキストボックス」にフォルダパスをセット</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonFirstFolderDialog_Click(object sender, EventArgs e)
        {
            SetSelectedDialogBoxFolderPathToTextBox(this.textBox1);
        }

        /// <summary>
        /// 週報出力項目「フォルダダイアログ」ボタンクリック時、イベント
        /// </summary>
        /// <remarks>週報出力項目「テキストボックス」にフォルダパスをセット</remarks>
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
                    string selectedPath = dialog.SelectedPath;
                    textBox.Text = selectedPath;
                }
            }
        }

        /// <summary>
        /// 「出力」ボタン押下時、イベント
        /// </summary>
        /// <remarks>日報データ出力項目「出力」ボタン押下時、イベント</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonOutputDailyReportData_Click(object sender, EventArgs e)
        {
            this.OutputReportFile(OutputType.DAILYREPORTDATA);
        }

        /// <summary>
        /// 「週報出力」ボタン押下時、イベント
        /// </summary>
        /// <remarks>週報出力項目「週報出力」ボタン押下時、イベント</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">イベントに関わる引数</param>
        private void ButtonOutputWeeklyReport_Click(object sender, EventArgs e)
        {
            this.OutputReportFile(OutputType.WEEKLYREPORT);
        }

        /// <summary>
        /// 日報データまたは週報を出力
        /// </summary>
        /// <remarks>テキストボックスにセットしたフォルダパスに日報データ（csvファイル）または週報（テキストファイル）を出力</remarks>
        /// <param name="outputType">出力タイプ</param>
        private void OutputReportFile(OutputType outputType)
        {
            if (this.Inputcheck(outputType))
            {
                switch (outputType)
                {
                    //日報データ
                    case OutputType.DAILYREPORTDATA:
                        this.OutputDailyReportData();
                        break;

                    //週報
                    case OutputType.WEEKLYREPORT:
                        this.OutputWeeklyReport();
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
            if (System.IO.Directory.Exists(this.textBox1.Text))
            {
                Boolean appendFlg = false;
                string outputPath;
                outputPath = this.textBox1.Text + dailyReportCsvFileNameWithExt;

                string outputDailyReportDataCompMsg = Properties.Resources.I0003;
                //出力ファイルパスが同名の場合
                if (File.Exists(outputPath))
                {
                    DialogResult dialogResult = MessageBox.Show(Properties.Resources.W0001, dialogBoxCaption, MessageBoxButtons.YesNo);
                    if (dialogResult == System.Windows.Forms.DialogResult.No)
                    {
                        //上書きしない
                        MessageBox.Show(Properties.Resources.I0005);
                        textBox1.ResetText();
                        return;
                    }
                    else if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                    {
                        //上書きする
                        outputDailyReportDataCompMsg = Properties.Resources.I0004;
                        appendFlg = true;
                    }
                    else
                    {
                        textBox1.ResetText();
                        //メソッドを抜ける
                        return;
                    }
                }
                File.Copy(CsvDailyReportDataPath, outputPath, appendFlg);

                MessageBox.Show(outputDailyReportDataCompMsg);
                this.textBox1.ResetText();
            }
            else
            {
                MessageBox.Show(Properties.Resources.E0004);
            }
        }

        /// <summary>
        /// 週報出力
        /// </summary>
        /// <remarks>週報（テキストファイル）を出力</remarks>
        private void OutputWeeklyReport()
        {
            if (System.IO.Directory.Exists(this.textBox2.Text))
            {
                Boolean appendFlg = false;
                string outputPath;
                outputPath = this.textBox2.Text + weeklyReportTxtFileNameWithExt;

                string outputWeeklyReportCompMsg = Properties.Resources.I0007;
                //出力ファイルパスが同名の場合
                if (File.Exists(outputPath))
                {
                    DialogResult dialogResult = MessageBox.Show(Properties.Resources.W0001, dialogBoxCaption, MessageBoxButtons.YesNo);
                    if (dialogResult == System.Windows.Forms.DialogResult.No)
                    {
                        //上書きしない
                        MessageBox.Show(Properties.Resources.I0005);
                        this.textBox2.ResetText();
                        this.textBox3.ResetText();
                        return;
                    }
                    else if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                    {
                        //上書きする
                        outputWeeklyReportCompMsg = Properties.Resources.I0008;
                        appendFlg = true;
                    }
                    else
                    {
                        this.textBox2.ResetText();
                        this.textBox3.ResetText();
                        //メソッドを抜ける
                        return;
                    }
                }
                MessageBox.Show("出力タイプが週報の場合、書き込み");
                MessageBox.Show(outputWeeklyReportCompMsg);
                this.textBox2.ResetText();
                this.textBox3.ResetText();
            }
            else
            {
                MessageBox.Show(Properties.Resources.E0004);
            }
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
            Boolean rtnFlag = true;

            //日報データ
            if (outputType == OutputType.DAILYREPORTDATA)
            {
                if (string.IsNullOrEmpty(this.textBox1.Text))
                {
                    var outputDailyReportDataFolderPathStrErrMsgEle = this.label1.Text + this.label2.Text.Substring(0, 8);
                    errMsgEleList.Add(outputDailyReportDataFolderPathStrErrMsgEle);
                }
            }

            //週報出力
            if (outputType == OutputType.WEEKLYREPORT)
            {
                if (string.IsNullOrEmpty(this.textBox2.Text))
                {
                    var outputWeeklyReportFolderPathStrErrMsgEle = this.label4.Text + this.label3.Text.Substring(0, 8);
                    errMsgEleList.Add(outputWeeklyReportFolderPathStrErrMsgEle);
                }

                if (string.IsNullOrEmpty(this.textBox3.Text))
                {
                    var tarMngNumStrErrMsgEle = this.label4.Text + this.label5.Text.Substring(0, 6);
                    errMsgEleList.Add(tarMngNumStrErrMsgEle);
                }
            }

            if (errMsgEleList.Count > 0)
            {
                string partialErrMsg = errMsgEleList.Aggregate((i, j) => i + InputDailyReportForm.ReadingPointStr + j);
                errMsg.Append(Properties.Resources.E0001.Replace(InputDailyReportForm.ReplaceErrMsgFirstArgStr, partialErrMsg));
            }

            //週報出力
            if (outputType == OutputType.WEEKLYREPORT)
            {
                //テキストボックスがnullまたは空白の場合、形式をチェックしない
                if (!errMsg.ToString().Contains(this.label4.Text + this.label5.Text.Substring(0, 6)))
                {
                    //形式が正しいかチェック
                    var pattern = RegExpTarMngNum;
                    if (!Regex.IsMatch(this.textBox3.Text, pattern))
                    {
                        if (errMsg.Length > 0)
                        {
                            errMsg.Append(DailyReportEntity.NewLineStr);
                        }
                        errMsg.Append(Properties.Resources.E0002);
                    }
                }
            }

            if (errMsg.Length > 0)
            {
                MessageBox.Show(errMsg.ToString());
                rtnFlag = false;
            }
            return rtnFlag;
        }

        /// データグリッドビューマウスポインターがセルに入ったときのイベント
        /// </summary>
        /// <remarks>データグリッドビューの「日付」、「実施内容」、「翌日予定」、「課題」項目データのセルにマウスポインターが入ったとき、データグリッドビューのスタイルを変更</remarks>
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
        /// <remarks>データグリッドビューの「日付」、「実施内容」、「翌日予定」、「課題」項目データのセルにマウスポインターが入ったとき、データグリッドビューのスタイルを変更</remarks>
        /// <param name="sender">イベントを送信したオブジェクト</param>
        /// <param name="e">データグリッドビューイベントに関わる引数</param>
        private void DailyReportDataListFormDataGridView_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex != 0)
            {
                ChangeDataGridViewStyle(e, Color.Empty, Cursors.Default);
            }
        }
    }
}
