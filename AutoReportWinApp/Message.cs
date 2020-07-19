using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace AutoReportWinApp
{
    public class Message
	{
		IDictionary<string, string> dicMsg = new Dictionary<string, string>();
        private readonly string notReadDataMsg = "メッセージ設定ファイルを対象パスに配置してください。";
        private readonly string notReadDataMsgID = "メッセージ設定ファイルにメッセージIDがありません。";
        private readonly string notReadDataMsgArgs = "メッセージIDの引数が設定されていません。";
        private readonly string SettingMsgPathEnd = @"\setting\Message.txt";
        private readonly string ReplaceArgsFirst = "{ARGFIRST}";
        public Message()
		{
            string SettingMsgPath = this.getSettingMsgPathPath();
            this.initSettingMsgReader(SettingMsgPath);
        }
        private string getSettingMsgPathPath()
        {
            Assembly myAssembly = Assembly.GetEntryAssembly();
            string exeFilePath = myAssembly.Location;
            string directoryName = System.IO.Path.GetDirectoryName(exeFilePath);
            return directoryName + this.SettingMsgPathEnd;
        }
        private void initSettingMsgReader(string readDataFilePath)
        {
            if (File.Exists(readDataFilePath))
            {
                using (var fileStream = new FileStream(readDataFilePath, FileMode.Open, FileAccess.Read))
                using (var streamReader = new StreamReader(fileStream, Encoding.Default))
                {
                    var rowNumIndex = 0;
                    var dailyReport = new DailyReport();

                    while (streamReader.Peek() >= 0)
                    {
                        string readingLine = streamReader.ReadLine();
                        string[] cols = readingLine.Split(SpecialStr.AppConstants.CommaChar);
                        dicMsg.Add(cols[0], cols[1]);
                        rowNumIndex++;
                    }
                }
            }
            else
            {
                MessageBox.Show(notReadDataMsg);
            }
        }
        public string get(string messageId)
		{
            string rtnMsg = notReadDataMsgID;
            if (this.dicMsg.TryGetValue(messageId, out string getValue))
            {
                rtnMsg = getValue;
            }
            return rtnMsg;
        }

		public string get(string messageId, string argFirst)
		{
            string rtnMsg = notReadDataMsgID;
            string getValue;
            if (this.dicMsg.TryGetValue(messageId, out getValue))
            {
                rtnMsg = notReadDataMsgArgs;
                if (getValue.Contains(this.ReplaceArgsFirst))
                {
                    rtnMsg = getValue.Replace(this.ReplaceArgsFirst, argFirst);
                }
            }
			return rtnMsg;
		}
	}
}
