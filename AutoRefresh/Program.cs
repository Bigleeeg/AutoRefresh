using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutoRefresh
{
    class Program
    {
        static void Main(string[] args)
        {
            LogIn2("西稍门z02", "liu321");
            //GetCookie("", "", "西稍门z59", "liu321");
        }

        private static void DoTest()
        {
            string path = @"C:\Users\kj01\Desktop\test.mcr";  //测试一个word文档


            DataTable dt = GetDataTableByExcel(@"C:\Users\kj01\Desktop\111.xlsx");
            if (dt != null && dt.Rows.Count > 0)
            {
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

                Thread.Sleep(40000);
            }
            catch
            {

            }
        }


        private static DataTable GetDataTableByExcel(string filePath)
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
                    string sql = string.Format("select * from [{0}$]", "Sheet1");
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

        private static void LogIn(string userName, string passWord)
        {

            string postString = @"callback=jQuery18204269285847678401_1480334545434&username={0}&password={1}&checkCode=&setcookie=14&second=&parentfunc=&redirect_in_iframe=&next=%2F&__hash__=yUspR%2FZe9QvSlCpmWojpC7wXsPfslKE2WcOxGLPSM8f5THDnVvbleSinwqABKrcA&_=1480335250600";//这里即为传递的参数，可以用工具抓包分析，也可以自己分析，主要是form里面每一个name都要加进来  
            postString = string.Format(postString, userName, passWord);
            byte[] postData = Encoding.UTF8.GetBytes(postString);//编码，尤其是汉字，事先要看下抓取网页的编码方式  
            string url = "https://passport.ganji.com/login.php?";//地址  
            WebClient webClient = new WebClient();
            //webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");//采取POST方式必须加的header，如果改为GET方式的话就去掉这句话即可  
            byte[] responseData = webClient.UploadData(url, "GET", postData);//得到返回字符流  
            string srcString = Encoding.UTF8.GetString(responseData);//解码 


            string url1 = "http://sync.ganji.com.cn/passport/sync.php?postData=%7B%22source%22%3A%22passport%22%2C%22act%22%3A%22login%22%2C%22site_id%22%3A5%2C%22sscode%22%3A%22SeyjCzz2tkWn%2B9oiSeVGRrlr%22%2C%22cookie_expire%22%3A1481542016%7D";
            postData = new byte[1];
            byte[] responseData1 = webClient.UploadData(url1, "POST", postData);//得到返回字符流  
            string srcString1 = Encoding.UTF8.GetString(responseData);//解码 

        }


        private static void LogIn1(string userName, string passWord)
        {
            string postString = @"username={0}&password={1}";//这里即为传递的参数，可以用工具抓包分析，也可以自己分析，主要是form里面每一个name都要加进来  
            postString = string.Format(postString, userName, passWord);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://passport.ganji.com/login.php");
            request.CookieContainer = new CookieContainer();
            CookieContainer cookie = request.CookieContainer;//如果用不到Cookie，删去即可  
                                                             //以下是发送的http头，随便加，其中referer挺重要的，有些网站会根据这个来反盗链  
            request.Referer = "https://passport.ganji.com/login.php?next=/";
            request.Accept = "Accept:text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            request.Headers["Accept-Language"] = "zh-CN,zh;q=0.";
            request.Headers["Accept-Charset"] = "GBK,utf-8;q=0.7,*;q=0.3";
            request.UserAgent = "User-Agent:Mozilla/5.0 (Windows NT 5.1) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.835.202 Safari/535.1";
            request.KeepAlive = true;
            //上面的http头看情况而定，但是下面俩必须加  
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "POST";

            Encoding encoding = Encoding.UTF8;//根据网站的编码自定义  
            byte[] postData = encoding.GetBytes(postString);//postDataStr即为发送的数据，格式还是和上次说的一样  
            request.ContentLength = postData.Length;
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(postData, 0, postData.Length);

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream responseStream = response.GetResponseStream();
            //如果http头中接受gzip的话，这里就要判断是否为有压缩，有的话，直接解压缩即可  
            //if (response.Headers["Content-Encoding"] != null && response.Headers["Content-Encoding"].ToLower().Contains("gzip"))
            //{
            //    responseStream = new GZipStream(responseStream, CompressionMode.Decompress);
            //}



            StreamReader streamReader = new StreamReader(responseStream, encoding);
            string retString = streamReader.ReadToEnd();

            //string ss = response.Headers["Set-Cookie"].ToString();
            //foreach (Cookie cookieItem in response.Cookies)
            //{
            //    cookie.Add(cookieItem);
            //}

            streamReader.Close();
            responseStream.Close();

            var xxx = cookie;

            //return retString;
        }



        private static void LogIn2(string userName, string passWord)
        {
            //string postString = @"username={0}&password={1}";//这里即为传递的参数，可以用工具抓包分析，也可以自己分析，主要是form里面每一个name都要加进来  
            //postString = string.Format(postString, userName, passWord);

            string postString = @"";

            CookieContainer cookie = new CookieContainer();

            cookie.SetCookies(new Uri("http://xa.ganji.com/"), "id58=c5/nn1g9JJTBaMvOHSX6Ag==");
             

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://xa.ganji.com/");
            //request.CookieContainer = new CookieContainer();
            //CookieContainer cookie = request.CookieContainer;//如果用不到Cookie，删去即可  
            //以下是发送的http头，随便加，其中referer挺重要的，有些网站会根据这个来反盗链  


            //request.CookieContainer = GetCookie("", "", "", "");

            request.CookieContainer = cookie;

            request.Referer = "http://xa.ganji.com/";
            request.Accept = "Accept:text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            request.Headers["Accept-Language"] = "zh-CN,zh;q=0.";
            request.Headers["Accept-Charset"] = "GBK,utf-8;q=0.7,*;q=0.3";
            request.UserAgent = "User-Agent:Mozilla/5.0 (Windows NT 5.1) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.835.202 Safari/535.1";
            request.KeepAlive = true;
            //上面的http头看情况而定，但是下面俩必须加  
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "POST";

            Encoding encoding = Encoding.UTF8;//根据网站的编码自定义  
            byte[] postData = encoding.GetBytes(postString);//postDataStr即为发送的数据，格式还是和上次说的一样  
            request.ContentLength = postData.Length;
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(postData, 0, postData.Length);

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream responseStream = response.GetResponseStream();
            //如果http头中接受gzip的话，这里就要判断是否为有压缩，有的话，直接解压缩即可  
            //if (response.Headers["Content-Encoding"] != null && response.Headers["Content-Encoding"].ToLower().Contains("gzip"))
            //{
            //    responseStream = new GZipStream(responseStream, CompressionMode.Decompress);
            //}



            StreamReader streamReader = new StreamReader(responseStream, encoding);
            string retString = streamReader.ReadToEnd();

            //string ss = response.Headers["Set-Cookie"].ToString();
            //foreach (Cookie cookieItem in response.Cookies)
            //{
            //    cookie.Add(cookieItem);
            //}

            streamReader.Close();
            responseStream.Close();

            var xxx = cookie;

            //return retString;
        }



        static CookieContainer GetCookie(string postString, string postUrl, string userName, string passWord)
        {
            postString = @"source=ganji_pc_login&xxzltr=3$$-1|-1$$ganji_uuid,2186244710381245062971|GANJISESSID,bd2f006392ab6b5dffdf19203ad099d6$$https://passport.ganji.com/login.php?next=/$$1480474928250$$9$$http://xa.ganji.com/$$1480474929566,358,193$$-1$$usename,1480474932131|usepassword,1480474935879$$usename,1480474935879$$21,23";
            postUrl = "https://cdata.58.com/btData";

            CookieContainer cookie = new CookieContainer(); 

            //https客户端证书  

            //HttpWebRequest request;
            ////blabla
            //X509Certificate cer = X509Certificate.CreateFromCertFile(“你的cer证书文件”);
            //request.ClientCertificates.Add(cer);
            ////blabla
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);

            HttpWebRequest httpRequset = (HttpWebRequest)HttpWebRequest.Create(postUrl);//创建http 请求
            httpRequset.Referer = "https://passport.ganji.com/login.php?next=/";
            httpRequset.CookieContainer = cookie;//设置cookie
            httpRequset.Method = "POST";//POST 提交
            httpRequset.KeepAlive = true;
            httpRequset.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko";
            httpRequset.Accept = "text/html, application/xhtml+xml, image/jxr, */*";
            httpRequset.Headers["Accept-Encoding"] = "gzip, deflate";
            httpRequset.Headers["Accept-Language"] = "en-US,en;q=0.8,zh-Hans-CN;q=0.5,zh-Hans;q=0.3";
            httpRequset.KeepAlive = true;
            httpRequset.ContentType = "application/x-www-form-urlencoded";//以上信息在监听请求的时候都有的直接复制过来
            httpRequset.Host = "cdata.58.com";
            httpRequset.AllowAutoRedirect = false;
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(postString);
            httpRequset.ContentLength = bytes.Length;
            Stream stream = httpRequset.GetRequestStream();

            stream.Write(bytes, 0, bytes.Length);
            stream.Close();//以上是POST数据的写入

            HttpWebResponse httpResponse = (HttpWebResponse)httpRequset.GetResponse();//获得 服务端响应
            string ss = httpResponse.Headers.Get("Set-Cookie").ToString();
            var xxx = httpResponse.Cookies;

            Stream responseStream = httpResponse.GetResponseStream();
            Encoding encoding = Encoding.UTF8;//根据网站的编码自定义  
            StreamReader streamReader = new StreamReader(stream, encoding);
            string retString = streamReader.ReadToEnd();


            return cookie;//拿到cookie
        }

        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true; //总是接受  
        }
    }
}
