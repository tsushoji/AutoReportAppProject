using CsvHelper;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace AutoReportWinApp
{
    public partial class DailyReportDataListForm : Form
    {
        private DataGridView _dataGridView1;
        internal Dictionary<int, DailyReportEntity> _csvDailyReportDataMap;
        private string _csvDailyReportDataPath;
        public DailyReportDataListForm()
        {
            InitializeComponent();
            DataGridView1 = this.dataGridView1;
            this._csvDailyReportDataMap = new Dictionary<int, DailyReportEntity>();
            CsvDailyReportDataPath = this.csvDailyReportDataPathGet();
            this.initDailyReportDataReader(CsvDailyReportDataPath);
        }

        public DataGridView DataGridView1 { get => _dataGridView1; set => _dataGridView1 = value; }
        public string CsvDailyReportDataPath { get => _csvDailyReportDataPath; set => _csvDailyReportDataPath = value; }
        private string csvDailyReportDataPathGet()
        {
            Assembly myAssembly = Assembly.GetEntryAssembly();
            string exeFilePath = myAssembly.Location;
            string directoryName = System.IO.Path.GetDirectoryName(exeFilePath);
            string path = directoryName + SetValue.AppConstants.CsvDailyReportDataPathEnd;
            return path;
        }
        private void initDailyReportDataReader(string createDataFilePath)
        {
            if (File.Exists(createDataFilePath))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = new StreamReader(createDataFilePath, Encoding.GetEncoding(SetValue.AppConstants.WinEncoding)))
                using (var csv = new CsvReader(reader, CultureInfo.CurrentCulture))
                {
                    var rowNumIndex = 0;
                    csv.Configuration.HasHeaderRecord = false; // ヘッダーの有無
                    var dailyReports = csv.GetRecords<DailyReportEntity>(); // データ読み出し（IEnumerable<Item>として受け取る）
                    foreach (DailyReportEntity dailyReport in dailyReports)
                    {
                        DataGridView1.Rows.Add(dailyReport.controlNum, dailyReport.date, DailyReportEntity.replaceToStrWithNewLine(dailyReport.impContent), DailyReportEntity.replaceToStrWithNewLine(dailyReport.schContent), DailyReportEntity.replaceToStrWithNewLine(dailyReport.task));
                        this._csvDailyReportDataMap.Add(rowNumIndex, dailyReport);
                        rowNumIndex++;
                    }
                }
            }
            else
            {
                string createDataDirectoryPath = System.IO.Path.GetDirectoryName(createDataFilePath);
                if (!System.IO.Directory.Exists(createDataDirectoryPath))
                {
                    Directory.CreateDirectory(createDataDirectoryPath);
                }
                File.Create(createDataFilePath).Close();
            }
        }

        private void upDailyReport_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex != 0)
            {
                using (var inputDailyReportForm = new InputDailyReportForm())
                {
                    inputDailyReportForm.DailyReportDataListForm = this;
                    inputDailyReportForm.CreateDataMode = CreateDataMode.APPEND;
                    if (this._csvDailyReportDataMap.Count > 0)
                    {
                        inputDailyReportForm.CreateDataColNum = this.maxColNumGet(this._csvDailyReportDataMap) + 1;
                    }
                    else
                    {
                        inputDailyReportForm.CreateDataColNum = SetValue.AppConstants.CreateReportDataColNumFirst;
                    }

                    if (e.RowIndex != this.DataGridView1.Rows.Count - 1)
                    {
                        inputDailyReportForm.CreateDataMode = CreateDataMode.UPDATE;
                        inputDailyReportForm.CreateDataColNum = Int32.Parse(this.DataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                        inputDailyReportForm.TextBox1Text = this.DataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                        inputDailyReportForm.TextBox2Text = this.DataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                        inputDailyReportForm.TextBox3Text = this.DataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                        inputDailyReportForm.TextBox4Text = this.DataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    }

                    inputDailyReportForm.ShowDialog();
                }
            }
        }
        private int maxColNumGet(Dictionary<int, DailyReportEntity> dataReportMap)
        {
            var colNumList = new List<int>();
            int controlNum = 0;
            foreach (KeyValuePair<int, DailyReportEntity> keyValuePair in dataReportMap)
            {
                controlNum = Int32.Parse(keyValuePair.Value.controlNum);
                colNumList.Add(controlNum);
            }
            return colNumList.Max();
        }

        private void DailyReportDataListForm_Load(object sender, EventArgs e)
        {
            this.Activate();
            this.dataGridView1.CurrentCell = null;
            this.dataGridView1.RowHeadersVisible = false;
        }
        private void ChangeDataGridViewStyle(DataGridViewCellEventArgs e, Color color, Cursor cursor)
        {
            DataGridView1.Cursor = cursor;
            DataGridView1.CurrentCell = null;
            var dataGridViewRowStyle = DataGridView1.Rows[e.RowIndex].DefaultCellStyle;
            dataGridViewRowStyle.BackColor = color;
            dataGridViewRowStyle.SelectionBackColor = color;
        }
        private void buttonFolderDialog_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.SelectedPath = SetValue.AppConstants.FolderDialogBoxInitSelected;
                dialog.Description = Properties.Resources.I0006;
                dialog.ShowNewFolderButton = true;

                DialogResult result = dialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    string selectedPath = dialog.SelectedPath;
                    this.textBox1.Text = selectedPath;
                }
            }
        }

        private void buttonDailyReportDataOutput_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.textBox1.Text))
            {
                MessageBox.Show(Properties.Resources.E0001);
                return;
            }

            if (System.IO.Directory.Exists(this.textBox1.Text))
            {
                Boolean appendFlg = false;
                string outputPath = this.textBox1.Text + SetValue.AppConstants.OutputFilePathEnd;
                string outputDailyReportDataCompMsg = Properties.Resources.I0003;
                if (File.Exists(outputPath))
                {
                    DialogResult dialogResult = MessageBox.Show(Properties.Resources.W0001, SetValue.AppConstants.AppendOutputFileDialogBoxTitle, MessageBoxButtons.YesNo);
                    if (dialogResult == System.Windows.Forms.DialogResult.No)
                    {
                        MessageBox.Show(Properties.Resources.I0005);
                        this.textBox1.ResetText();
                        return;
                    }
                    else if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                    {
                        outputDailyReportDataCompMsg = Properties.Resources.I0004;
                        appendFlg = true;
                    }
                    else
                    {
                        this.textBox1.ResetText();
                        return;
                    }
                }

                File.Copy(CsvDailyReportDataPath, outputPath, appendFlg);
                MessageBox.Show(outputDailyReportDataCompMsg);
                this.textBox1.ResetText();
            }
            else
            {
                MessageBox.Show(Properties.Resources.E0005);
            }
        }

        private void upDailyReport_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex != 0)
            {
                ChangeDataGridViewStyle(e, Color.Yellow, Cursors.Hand);
            }
        }

        private void upDailyReport_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex != 0)
            {
                ChangeDataGridViewStyle(e, Color.Empty, Cursors.Default);
            }
        }
    }
}
