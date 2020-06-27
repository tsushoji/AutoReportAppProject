using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoReportWinApp
{
    public partial class StartMenuForm : Form
    {
        private static string _createDataFilePath;
        public StartMenuForm()
        {
            InitializeComponent();
        }

        public static string CreateDataFilePath { get => _createDataFilePath; set => _createDataFilePath = value; }
        private void buttonStart_Click(object sender, EventArgs e)
        {
            if (this.inputcheck(this.textBox1.Text)) 
            {
                StartMenuForm.CreateDataFilePath = this.textBox1.Text;

                using (var dailyReportDataListForm = new DailyReportDataListForm())
                {
                    this.Hide();
                    dailyReportDataListForm.StartMenuForm = this;
                    dailyReportDataListForm.ShowDialog();
                }
            }
        }

        private Boolean inputcheck(string inputCreateDataFilePath)
        {
            if ((inputCreateDataFilePath == null) || (inputCreateDataFilePath.Length == 0))
            {
                MessageBox.Show(AppConstants.NotInputFilePathMsg);
                return false;
            }

            return true;
        }

        private void buttonFileDialog_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog()) 
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    this.textBox1.Text = openFileDialog.FileName;
                }
            }
        }
    }
}
