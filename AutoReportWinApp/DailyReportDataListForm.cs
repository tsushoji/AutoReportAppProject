using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoReportWinApp
{
    public partial class DailyReportDataListForm : Form
    {
        private StartMenuForm _startMenuForm;
        private int _createDataColNum = AppConstants.CreateDataFirstColNum;
        public DailyReportDataListForm()
        {
            InitializeComponent();
            this.initDailyReportDataReader(StartMenuForm.CreateDataFilePath);
        }

        public StartMenuForm StartMenuForm { get => _startMenuForm; set => _startMenuForm = value; }
        public int CreateDataColNum { get => _createDataColNum; set => _createDataColNum = value; }
        public void initDailyReportDataReader(string createDataFilePath) 
        {
            if (File.Exists(createDataFilePath))
            {
                var colNumList = new List<int>();
                this.dataGridView1.Rows.Clear();

                using (var fileStream = new FileStream(createDataFilePath, FileMode.Open, FileAccess.Read))
                using (var streamReader = new StreamReader(fileStream, Encoding.Default))
                {
                    while (streamReader.Peek() >= 0)
                    {
                        string[] cols = streamReader.ReadLine().Split(',');
                        this.dataGridView1.Rows.Add(cols[0], cols[1], this.replaceRow(cols[2]), this.replaceRow(cols[3]), this.replaceRow(cols[4]));
                        colNumList.Add(Int32.Parse(cols[0]));
                    }
                }
                this.CreateDataColNum = colNumList.Max() + 1;
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

        private void upDailyReport_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                using (var inputDailyReportForm = new InputDailyReportForm())
                {
                    inputDailyReportForm.DailyReportDataListForm = this;
                    inputDailyReportForm.CreateDataMode = CreateDataMode.APPEND;

                    if (e.RowIndex != this.dataGridView1.Rows.Count - 1) 
                    {
                        inputDailyReportForm.CreateDataMode = CreateDataMode.UPDATE;
                        inputDailyReportForm.CreateDataUpdateColNum = Int32.Parse(this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                        inputDailyReportForm.TextBox1Text = this.dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                        inputDailyReportForm.TextBox2Text = this.dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                        inputDailyReportForm.TextBox3Text = this.dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                        inputDailyReportForm.TextBox4Text = this.dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    }
                    
                    inputDailyReportForm.ShowDialog();
                }
            }
        }

        private void DailyReportDataListForm_Load(object sender, EventArgs e)
        {
            this.Activate();
        }

        private void DailyReportDataListForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.StartMenuForm.Close();
        }
    }
}
