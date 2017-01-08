using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoRefresh
{
    public partial class Form2 : Form
    {
        private Thread freshThread = null;

        private static string webSiteName = "58";

        private static int intervalMin = 25000;

        public Form2()
        {
            InitializeComponent();
            this.comboBox1.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            webSiteName = this.comboBox1.SelectedIndex == 0 ? "ganji" : "58";
            intervalMin = 25000;

            try
            {
                intervalMin = Convert.ToInt32(this.textBox1.Text);
            }
            catch
            {

            }
            this.WindowState = FormWindowState.Minimized;
            intervalMin = intervalMin <= 25 ? 25000 : intervalMin * 1000;
            //新启动一个线程 处理刷新
            freshThread = new Thread(new ThreadStart(DoRefresh));
            freshThread.Start();

            //DoRefresh(webSiteName, intervalMin);
            //System.Environment.Exit(0);
        }

        public static List<string> webSiteList = new List<string>() { "ganji", "58" };
        private static void DoRefresh()
        {
            List<int> intList = new List<int>();
            List<int> rList = new List<int>();
            string path = string.Format(@"D:\AutoRefresh\{0}.mcr", webSiteName);
            string logPath = string.Format(@"D:\AutoRefresh\Log\{0}_{1}.LOG", webSiteName, DateTime.Now.ToString("yyyy-MM-dd"));
            if (!File.Exists(logPath))
            {
                File.Create(logPath);
            }

            string content = ReadLog(logPath);
            List<string> striparr = content.Split(new string[] { "\r\n" }, StringSplitOptions.None).ToList();
            striparr = striparr.Where(s => !string.IsNullOrEmpty(s)).ToList();

            DataTable dt = GetDataTableByExcel(@"D:\AutoRefresh\refresh.xlsx", webSiteName);
            List<string> killProcess = new List<string>() { "firefox", "macrorecorder", "flashplayerplugin_23_0_0_207" };
            if (dt != null && dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    intList.Add(i);
                }
                rList = randomList(intList);

                WriteLog(logPath, GenerateLogHead(webSiteName));

                int existAuthCodeTimes = 0;

                foreach (int r in rList)
                {
                    try
                    {
                        Process[] processesBegin = Process.GetProcesses();
                        foreach (var item in processesBegin)
                        {
                            if (killProcess.Contains(item.ProcessName.ToLower()))
                            {
                                item.Kill();
                            }
                        }


                        if (!string.IsNullOrEmpty(dt.Rows[r]["Username"].ToString().Trim()) && !string.IsNullOrEmpty(dt.Rows[r]["Password"].ToString().Trim()))
                        {
                            if (!striparr.Contains(dt.Rows[r]["Username"].ToString().Trim()))
                            {
                                GenerateFile(path, dt.Rows[r]["Username"].ToString().Trim(), dt.Rows[r]["Password"].ToString().Trim(), intervalMin, existAuthCodeTimes);
                            }
                        }

                        string clipStr = ClipboardAsync.GetText();

                        if (!clipStr.Contains("验证码"))
                        {
                            WriteLog(logPath, dt.Rows[r]["Username"].ToString().Trim() + "\r\n");
                        }
                        else
                        {
                            if (existAuthCodeTimes < 5)
                            {
                                existAuthCodeTimes = existAuthCodeTimes + 1;
                            }
                            else
                            {
                                existAuthCodeTimes = 0;
                            }
                        }

                        Process[] processesEnd = Process.GetProcesses();
                        foreach (var item in processesEnd)
                        {
                            if (killProcess.Contains(item.ProcessName.ToLower()))
                            {
                                item.Kill();
                            }
                        }

                    }
                    catch (Exception e)
                    {

                    }
                }
                WriteLog(logPath, GenerateLogEnd(webSiteName, dt.Rows.Count));
            }
        }

        private static void GenerateFile(string path, string userName, string passWord, int intervalMin, int existAuthCodeTimes)
        {
            bool isPassword = false;
            string content = string.Empty;
            try
            {
                using (StreamReader sr = new StreamReader(path, System.Text.Encoding.GetEncoding("utf-8")))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {

                        if (line.Contains("TYPE TEXT"))
                        {
                            line = string.Format("TYPE TEXT : {0}", isPassword ? passWord : userName);
                            isPassword = !isPassword;
                        } 
                        content += line + "\r\n";
                    }
                    sr.Close();
                }
                File.WriteAllText(path, content);
                System.Diagnostics.Process.Start(path); //打开此文件。  

                //Application.DoEvents();

                //出现五次验证码 等待10分钟
                if (existAuthCodeTimes == 5)
                {
                    intervalMin = 10000;
                }
                Thread.Sleep(intervalMin);

            }
            catch
            {

            }
        }

        private static DataTable GetDataTableByExcel(string filePath, string webSiteName)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            try
            {
                if (!string.IsNullOrEmpty(filePath))
                {
                    FileInfo file = new FileInfo(filePath);
                    string ConnStr = string.Empty;
                    string extension = file.Extension;
                    switch (extension)
                    {
                        case ".xls":
                            ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                            break;
                        case ".xlsx":
                            ConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                            break;
                        default:
                            ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                            break;
                    }

                    OleDbConnection Conn = new OleDbConnection(ConnStr);
                    Conn.Open();
                    //填充数据
                    string sql = string.Format("select * from [{0}$]", webSiteName);
                    OleDbDataAdapter da = new OleDbDataAdapter(sql, ConnStr);
                    da.Fill(ds);
                }

                if (ds != null && ds.Tables.Count > 0)
                {
                    dt = ds.Tables[0];
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return dt;
        }

        private static string GenerateLogHead(string webSiteName)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("---------------------------------------------------------------\r\n");
            sb.AppendFormat("-------------{0} 开始刷新 {1}-----------------\r\n", DateTime.Now, webSiteName);

            return sb.ToString();
        }

        private static string GenerateLogEnd(string webSiteName, int rowNums)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("-------------{0} 刷新完成 {1}-----------------\r\n", DateTime.Now, webSiteName);
            sb.AppendFormat("-------------共刷新 {0} 条数据 ----------------------------------\r\n", rowNums.ToString());

            return sb.ToString();
        }

        private static void WriteLog(string path, string writeLog)
        {
            StreamReader sr = new StreamReader(path, System.Text.Encoding.GetEncoding("utf-8"));
            string content = sr.ReadToEnd().ToString();
            sr.Close();
            content += writeLog;
            File.WriteAllText(path, content);
        }

        private static string ReadLog(string path)
        {
            StreamReader sr = new StreamReader(path, System.Text.Encoding.GetEncoding("utf-8"));
            string content = sr.ReadToEnd().ToString();
            sr.Close();
            return content;
        }

        //private static void CleanLog(string path)
        //{
        //    File.WriteAllText(path, "");
        //}


        private static List<int> randomList(List<int> ContentList)
        {
            Random random = new Random();
            List<int> newList = new List<int>();
            foreach (int item in ContentList)
            {
                newList.Insert(random.Next(newList.Count), item);
            }
            return newList;

        }

        //暂停
        private void button2_Click(object sender, EventArgs e)
        {
            if (freshThread != null && freshThread.IsAlive)
            {
                freshThread.Abort();
            }
            Process[] processes = Process.GetProcessesByName("MacroRecorder");
            if (processes.Length > 0)
            {
                processes[0].Kill();
            }
        }

        //停止
        private void button3_Click(object sender, EventArgs e)
        {
            Process[] processes = Process.GetProcessesByName("MacroRecorder");
            if (processes.Length > 0)
            {
                processes[0].Kill();
            }
            System.Environment.Exit(0);
        }


        class ClipboardAsync
        {
            private string _getText;
            private void ThGetText(object format)
            {
                try
                {
                    _getText = format == null ? Clipboard.GetText() : Clipboard.GetText((TextDataFormat)format);
                }
                catch
                {
                    _getText = null;
                }
            }

            public static string GetText()
            {
                var instance = new ClipboardAsync();
                var staThread = new Thread(instance.ThGetText);
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
                return instance._getText;
            }

            public static string GetText(TextDataFormat format)
            {
                var instance = new ClipboardAsync();
                var staThread = new Thread(instance.ThGetText);
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start(format);
                staThread.Join();
                return instance._getText;
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Process[] processes = Process.GetProcessesByName("MacroRecorder");
            if (processes.Length > 0)
            {
                processes[0].Kill();
            }
            System.Environment.Exit(0);
        }
    }
}
