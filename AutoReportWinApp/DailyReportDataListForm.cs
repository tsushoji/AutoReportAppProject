using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoReportWinApp
{
    public partial class DailyReportDataListForm : Form
    {
        private DataGridView _dataGridView1;
        internal Dictionary<int, DailyReport> _csvDailyReportDataMap;
        private string _csvDailyReportDataPath;
        public DailyReportDataListForm()
        {
            InitializeComponent();
            DataGridView1 = this.dataGridView1;
            this._csvDailyReportDataMap = new Dictionary<int, DailyReport>();
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
            string path = directoryName + AppConstants.CsvDailyReportDataPathEnd;
            return path;
        }
        private void initDailyReportDataReader(string createDataFilePath) 
        {
            if (File.Exists(createDataFilePath))
            {
                using (var fileStream = new FileStream(createDataFilePath, FileMode.Open, FileAccess.Read))
                using (var streamReader = new StreamReader(fileStream, Encoding.Default))
                {
                    var rowNumIndex = 0;
                    var dailyReport = new DailyReport();

                    while (streamReader.Peek() >= 0)
                    {
                        string readingLine = streamReader.ReadLine();
                        string[] cols = readingLine.Split(AppConstants.DailyReportDataLineSeparateChar);
                        int rowConNum = Int32.Parse(cols[0]);
                        dailyReport.ControlNum = rowConNum;
                        dailyReport.CsvDailyReportLine = readingLine;
                        this._csvDailyReportDataMap.Add(rowNumIndex, dailyReport);
                        DataGridView1.Rows.Add(cols[0], cols[1], this.replaceRow(cols[2]), this.replaceRow(cols[3]), this.replaceRow(cols[4]));
                        rowNumIndex++;
                    }
                }
            }
            else 
            {
                File.Create(createDataFilePath).Close();
            }
        }
        private string replaceRow(string row) 
        {
            if (row.Contains(AppConstants.UserNewLineStr))
            {
                return row.Replace(AppConstants.UserNewLineStr, AppConstants.NewLineStr);
            }
            return row;
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
                        inputDailyReportForm.CreateDataColNum = AppConstants.CreateReportDataColNumFirst;
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

        private int maxColNumGet(Dictionary<int, DailyReport> dataReportMap)
        {
            var colNumList = new List<int>();
            foreach (KeyValuePair<int, DailyReport> keyValuePair in dataReportMap)
            {
                colNumList.Add(keyValuePair.Value.ControlNum);
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
            for (var i = 0; i < DataGridView1.Columns.Count; i++)
            {
                DataGridView1.Rows[e.RowIndex].Cells[i].Style.BackColor = color;
                DataGridView1.Rows[e.RowIndex].Cells[i].Style.SelectionBackColor = color;
            }
        }
        private void buttonFolderDialog_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.SelectedPath = AppConstants.FolderDialogBoxInitSelected;
                dialog.Description = AppConstants.FolderDialogBoxExplation;
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
            if (System.IO.Directory.Exists(this.textBox1.Text))
            {
                Boolean appendFlg = false;
                string outputPath = this.textBox1.Text + AppConstants.OutputFilePathEnd;
                string outputDailyReportDataCompMsg = AppConstants.OutputDailyReportDataCompMsg;
                if (File.Exists(outputPath))
                {
                    DialogResult dialogResult = MessageBox.Show(AppConstants.AppendOutputFileDialogBoxMsg, AppConstants.AppendOutputFileDialogBoxTitle, MessageBoxButtons.YesNo);
                    if (dialogResult == System.Windows.Forms.DialogResult.No)
                    {
                        MessageBox.Show(AppConstants.CancelMsg);
                        this.textBox1.ResetText();
                        return;
                    }
                    else if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                    {
                        outputDailyReportDataCompMsg = AppConstants.OutputDailyReportDataAppendCompMsg;
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
                MessageBox.Show(AppConstants.NotOutputFolderPathMsg);
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
