﻿namespace AutoReportWinApp
{
    partial class InputDailyReportForm
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.buttonCalendar = new System.Windows.Forms.Button();
            this.buttonCreate = new System.Windows.Forms.Button();
            this.buttonForDataList = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(33, 106);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 34);
            this.label1.TabIndex = 0;
            this.label1.Text = "日付：";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox1.Location = new System.Drawing.Point(217, 106);
            this.textBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(372, 36);
            this.textBox1.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(33, 166);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(168, 34);
            this.label2.TabIndex = 2;
            this.label2.Text = "実施内容：";
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox2.Location = new System.Drawing.Point(217, 166);
            this.textBox2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(725, 174);
            this.textBox2.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(33, 374);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(168, 34);
            this.label3.TabIndex = 4;
            this.label3.Text = "翌日予定：";
            // 
            // textBox3
            // 
            this.textBox3.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox3.Location = new System.Drawing.Point(217, 374);
            this.textBox3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(725, 174);
            this.textBox3.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(33, 582);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 34);
            this.label4.TabIndex = 6;
            this.label4.Text = "課題：";
            // 
            // textBox4
            // 
            this.textBox4.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox4.Location = new System.Drawing.Point(217, 582);
            this.textBox4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(725, 174);
            this.textBox4.TabIndex = 7;
            // 
            // buttonCalendar
            // 
            this.buttonCalendar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonCalendar.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.buttonCalendar.Location = new System.Drawing.Point(864, 106);
            this.buttonCalendar.Margin = new System.Windows.Forms.Padding(4);
            this.buttonCalendar.Name = "buttonCalendar";
            this.buttonCalendar.Size = new System.Drawing.Size(80, 38);
            this.buttonCalendar.TabIndex = 8;
            this.buttonCalendar.Text = "カレンダー";
            this.buttonCalendar.UseVisualStyleBackColor = true;
            this.buttonCalendar.Click += new System.EventHandler(this.buttonCalendar_Click);
            // 
            // buttonCreate
            // 
            this.buttonCreate.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonCreate.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.buttonCreate.Location = new System.Drawing.Point(741, 788);
            this.buttonCreate.Margin = new System.Windows.Forms.Padding(4);
            this.buttonCreate.Name = "buttonCreate";
            this.buttonCreate.Size = new System.Drawing.Size(203, 50);
            this.buttonCreate.TabIndex = 9;
            this.buttonCreate.Text = "データ作成";
            this.buttonCreate.UseVisualStyleBackColor = true;
            this.buttonCreate.Click += new System.EventHandler(this.buttonCreateData_Click);
            // 
            // buttonForDataList
            // 
            this.buttonForDataList.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonForDataList.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.buttonForDataList.Location = new System.Drawing.Point(717, 30);
            this.buttonForDataList.Margin = new System.Windows.Forms.Padding(4);
            this.buttonForDataList.Name = "buttonForDataList";
            this.buttonForDataList.Size = new System.Drawing.Size(227, 50);
            this.buttonForDataList.TabIndex = 10;
            this.buttonForDataList.Text = "データリストへ";
            this.buttonForDataList.UseVisualStyleBackColor = true;
            this.buttonForDataList.Click += new System.EventHandler(this.buttonForDataList_Click);
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("MS UI Gothic", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(603, 106);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(253, 38);
            this.label5.TabIndex = 11;
            this.label5.Text = "ex)2020/06/24";
            // 
            // InputDailyReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(992, 889);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.buttonForDataList);
            this.Controls.Add(this.buttonCreate);
            this.Controls.Add(this.buttonCalendar);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InputDailyReportForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "日報作成入力フォーム";
            this.Load += new System.EventHandler(this.InputDailyReportForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Button buttonCalendar;
        private System.Windows.Forms.Button buttonCreate;
        private System.Windows.Forms.Button buttonForDataList;
        private System.Windows.Forms.Label label5;
    }
}
