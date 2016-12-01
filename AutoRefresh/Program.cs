using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Threading;

namespace AutoRefresh
{
    class Program
    {
        public static List<string> webSiteList = new List<string>() { "ganji", "58" };

        static void Main(string[] args)
        {
            DoRefresh();
        }

        private static void DoRefresh()
        {
            foreach (string webSiteName in webSiteList)
            {
                string path = string.Format(@"C:\Users\kj01\Desktop\{0}.mcr", webSiteName);
                string logPath = @"C:\Users\kj01\Desktop\log.txt";
                DataTable dt = GetDataTableByExcel(@"C:\Users\kj01\Desktop\refresh.xlsx", webSiteName);
                if (dt != null && dt.Rows.Count > 0)
                {

                    WriteLog(logPath, GenerateLogHead(webSiteName));

                    foreach (DataRow dr in dt.Rows)
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(dr["Username"].ToString().Trim()) && !string.IsNullOrEmpty(dr["Password"].ToString().Trim()))
                            {
                                GenerateFile(path, dr["Username"].ToString().Trim(), dr["Password"].ToString().Trim());
                            }
                        }
                        catch
                        {

                        }
                    }

                    WriteLog(logPath, GenerateLogEnd(webSiteName,dt.Rows.Count));
                }

            }
        }

        private static void GenerateFile(string path, string userName, string passWord)
        {
            int i = 1;
            string content = string.Empty;
            try
            {
                using (StreamReader sr = new StreamReader(path, System.Text.Encoding.GetEncoding("utf-8")))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (i == 14)
                        {
                            line = string.Format("TYPE TEXT : {0}", userName);
                        }
                        if (i == 16)
                        {
                            line = string.Format("TYPE TEXT : {0}", passWord);
                        }
                        content += line + "\r\n";
                        i++;
                    }
                    sr.Close();
                }
                File.WriteAllText(path, content);
                System.Diagnostics.Process.Start(path); //打开此文件。 

                Thread.Sleep(25000);
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

    }
}
