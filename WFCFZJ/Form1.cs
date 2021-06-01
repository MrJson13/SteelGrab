using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Collections;
using System.IO;
using System.Reflection;

using System.Net;

namespace GrabSteel
{
    public partial class form1 : Form
    {
        public form1()
        {
            InitializeComponent();
        }
        public string XdPath = Application.StartupPath.ToString();
        /// <summary>
        /// 接口地址
        /// </summary>
        public string ApiUrl = "https://search.mysteel.com/searchapi/search/getMarketPrice";

        private void form1_Load(object sender, EventArgs e)
        {
            pager1.PageIndex = 1;
            QueryAllData();

            mytimer.Enabled = true;
            mytimer.Interval = 60000;//执行间隔时间,单位为毫秒;此时时间间隔为60秒
            mytimer.Start();   //定时器开始工作
        }
        /// <summary>
        /// 手动抓取【全部】
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnZQ_Click(object sender, EventArgs e)
        {
            btnZQ.Text = "抓取中...";

            btnZQ.Enabled = false;
            btxExe.Enabled = false;
            btnSearch.Enabled = false;
            ck_Auto.Enabled = false;
            textBox1.Enabled = false;

            //清空数据表
            GrabSteel.Common.SqlData.SelectDataTable("", " truncate table mySteel_GrabData ");
            //条件
            var endtime = DateTime.Now.ToString("yyyy-MM-dd 00:00:00");
            //Post参数
            string postString = "{\"query\":\"成都市钢材\",\"startTime\":\"2020-10-01 00:00:00\",\"endTime\":\"" + endtime + "\",\"sortType\":\"complex\",\"platform\":\"pc\",\"pageNo\":1,\"pageSize\":999}";

            //返回的json字符串
            var jsonStr = sendHttpRequest(ApiUrl, postString);

            Newtonsoft.Json.Linq.JObject jobject =
                (Newtonsoft.Json.Linq.JObject)Newtonsoft.Json.JsonConvert.DeserializeObject(jsonStr);

            var records = jobject["dataList"];

            //先登录 然后请求页面
            var LoginUrl = "https://passport.mysteel.com/loginJson.jsp?callback=loginJsn&my_username=minghua151515&my_password=49fd7e0d4867373b0821d78d4ce3d05f8&callbackJsonp=loginJsn&jumpPage=https%3A%2F%2Fwww.mysteel.com%2F&site=www.mysteel.com&my_rememberStatus=true&vcode=&_=1618296487431";

            CookieContainer cookie = new CookieContainer();
            //登录并获取cookie
            HttpPost(LoginUrl, "", ref cookie);

            for (int j = 0; j < records.Count(); j++)
            {
                //if (records[j]["resultType"].ToString() == "market")
                //{
                //    var strId = records[j]["marketBean"]["id"].ToString();
                //    var strTitle = records[j]["marketBean"]["title"].ToString();
                //    var strUrl = records[j]["marketBean"]["linkUrl"].ToString();
                //    var strWebDate = records[j]["publishTime"].ToString();
                //    var strContent = Post2(records[j]["marketBean"]["linkUrl"].ToString(), cookie);

                //    var strUrlLink = strUrl.Replace("//", "https://"); ;

                //    string strsql = "insert into mySteel_GrabData(id,title,url,content,webdate) "
                //        + " values('" + strId + "','" + strTitle + "','" + strUrlLink + "','" + strContent.Replace("'", "''").Trim() + "','" + strWebDate + "')";
                //    //判断是否存在重复记录
                //    string strsql2 = "select * from mySteel_GrabData where id='" + strId + "'";
                //    DataTable dt = GrabSteel.Common.SqlData.SelectDataTable("", strsql2);
                //    if (dt.Rows.Count == 0)
                //    {
                //        GrabSteel.Common.SqlData.InsDelUpdData("", strsql);
                //    }
                //}

                if (records[j]["resultType"].ToString() == "aggregate")
                {
                    var strId = records[j]["aggregateBean"]["id"].ToString();
                    var strTitle = records[j]["aggregateBean"]["title"].ToString();
                    var strUrl = records[j]["aggregateBean"]["linkUrl"].ToString();
                    var strWebDate = records[j]["publishTime"].ToString();
                    strTitle = Convert.ToDateTime(strWebDate).ToString("yyyy年MM月") + strTitle;
                    var strContent = Post(records[j]["aggregateBean"]["linkUrl"].ToString(), cookie);

                    var strUrlLink = strUrl.Replace("//", "https://"); ;

                    string strsql = "insert into mySteel_GrabData(id,title,url,content,webdate) "
                        + " values('" + strId + "','" + strTitle + "','" + strUrlLink + "','" + strContent.Replace("'", "''").Trim() + "','" + strWebDate + "')";
                    //判断是否存在重复记录
                    string strsql2 = "select * from mySteel_GrabData where id='" + strId + "'";
                    DataTable dt = GrabSteel.Common.SqlData.SelectDataTable("", strsql2);
                    if (dt.Rows.Count == 0)
                    {
                        GrabSteel.Common.SqlData.InsDelUpdData("", strsql);
                    }
                }
            }

            btnZQ.Enabled = true;
            btxExe.Enabled = true;
            btnSearch.Enabled = true;
            ck_Auto.Enabled = true;
            textBox1.Enabled = true;
            btnZQ.Text = "全部抓取...";

            pager1.PageIndex = 1;
            //调用查询全部的方法
            QueryAllData();

            MessageBox.Show("成功导入");
        }
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btxExe_Click(object sender, EventArgs e)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory.Replace("\\bin\\Debug", "\\") + "file\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
            string sql = " select grabID,title,url,webdate,createTime from mySteel_GrabData where 1=1 ";
            string txttitle = textBox1.Text.Trim();
            if (txttitle != "")
            {
                sql += " and title like '%" + txttitle + "%' ";
            }
            DataTable myTable = GrabSteel.Common.SqlData.SelectDataTable("", sql);
            if (dt2csv(myTable, path, "成都钢铁价格信息", "编号,标题,链接,发布时间,抓取时间"))
            {
                MessageBox.Show("导出成功,文件位置:" + path);
            }
            else
            {
                MessageBox.Show("导出失败");
            }
        }
        /// <summary>
        /// 导出报表为Csv
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="strFilePath">物理路径</param>
        /// <param name="tableheader">表头</param>
        /// <param name="columname">字段标题,逗号分隔</param>
        public bool dt2csv(DataTable dt, string strFilePath, string tableheader, string columname)
        {
            try
            {
                string strBufferLine = "";
                StreamWriter strmWriterObj = new StreamWriter(strFilePath, false, System.Text.Encoding.UTF8);
                strmWriterObj.WriteLine(tableheader);
                strmWriterObj.WriteLine(columname);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    strBufferLine = "";
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j > 0)
                            strBufferLine += ",";
                        strBufferLine += dt.Rows[i][j].ToString();
                    }
                    strmWriterObj.WriteLine(strBufferLine);
                }
                strmWriterObj.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// List转DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="collection"></param>
        /// <returns></returns>
        public DataTable ToDataTable<T>(IEnumerable<T> collection)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray());
            if (collection.Count() > 0)
            {
                for (int i = 0; i < collection.Count(); i++)
                {
                    ArrayList tempList = new ArrayList();
                    foreach (PropertyInfo pi in props)
                    {
                        object obj = pi.GetValue(collection.ElementAt(i), null);
                        tempList.Add(obj);
                    }
                    object[] array = tempList.ToArray();
                    dt.LoadDataRow(array, true);
                }
            }
            return dt;
        }
        private void QueryAllData()
        {
            try
            {
                string txttitle = textBox1.Text.Trim();
                string sql = " select count(1) from mySteel_GrabData where 1=1 ";
                if (txttitle != "")
                {
                    sql += " and title like '%" + txttitle + "%' ";
                }
                pager1.PageSize = 20;
                int[] arr = cofig.config.CountStartEnd(pager1.PageIndex, pager1.PageSize);
                object o = GrabSteel.Common.SqlData.ExecuteDataSql("", sql);
                int total = o == null ? 0 : (int)o;
                pager1.RecordCount = total;
                pager1.Page();

                string sql1 = "SELECT  [grabID],[title],[url],[webdate],[createTime] FROM(select *,ROW_NUMBER() over(order by webdate desc) rows from mySteel_GrabData where title like '%{0}%') t where rows between " + arr[0] + " and " + arr[1] + " ";
                sql1 = string.Format(sql1, textBox1.Text);
                DataTable myTable = GrabSteel.Common.SqlData.SelectDataTable("", sql1);
                dataGridView1.DataSource = myTable;

                //不允许添加行
                dataGridView1.AllowUserToAddRows = false;
                //背景为白色
                dataGridView1.BackgroundColor = Color.White;
                //只允许选中单行
                dataGridView1.MultiSelect = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("查询错误！" + ex.Message);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            pager1.PageIndex = 1;
            QueryAllData();
        }
        /// <summary>
        /// 执行的定时方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mytimer_Tick(object sender, EventArgs e)
        {
            //如果当前时间是1点00分
            DateTime dt = DateTime.Now;
            string[] arrTime = GrabSteel.cofig.config.GetTime(XdPath).Split(':');
            if (dt.Hour == Convert.ToInt32(arrTime[0]) && dt.Minute == Convert.ToInt32(arrTime[1]))
            {
                //执行定时抓取方法
                //条件-起始时间
                var beginTime = DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd 00:00:00");//抓取7天内的
                var endtime = DateTime.Now.ToString("yyyy-MM-dd 00:00:00");
                //Post参数
                string postString = "{\"query\":\"成都市钢材\",\"startTime\":\"" + beginTime + "\",\"endTime\":\"" + endtime + "\",\"sortType\":\"complex\",\"platform\":\"pc\",\"pageNo\":1,\"pageSize\":999}";
                //返回的json字符串
                var jsonStr = sendHttpRequest(ApiUrl, postString);

                Newtonsoft.Json.Linq.JObject jobject =
                    (Newtonsoft.Json.Linq.JObject)Newtonsoft.Json.JsonConvert.DeserializeObject(jsonStr);

                var records = jobject["dataList"];

                //先登录 然后请求页面
                var LoginUrl = "https://passport.mysteel.com/loginJson.jsp?callback=loginJsn&my_username=minghua151515&my_password=49fd7e0d4867373b0821d78d4ce3d05f8&callbackJsonp=loginJsn&jumpPage=https%3A%2F%2Fwww.mysteel.com%2F&site=www.mysteel.com&my_rememberStatus=true&vcode=&_=1618296487431";

                CookieContainer cookie = new CookieContainer();
                //登录并获取cookie
                HttpPost(LoginUrl, "", ref cookie);

                for (int j = 0; j < records.Count(); j++)
                {
                    if (records[j]["resultType"].ToString() == "aggregate")
                    {
                        var strId = records[j]["aggregateBean"]["id"].ToString();
                        var strTitle = records[j]["aggregateBean"]["title"].ToString();
                        var strUrl = records[j]["aggregateBean"]["linkUrl"].ToString();
                        var strWebDate = records[j]["publishTime"].ToString();
                        strTitle = Convert.ToDateTime(strWebDate).ToString("yyyy年MM月") + strTitle;
                        var strContent = Post(records[j]["aggregateBean"]["linkUrl"].ToString(),cookie);

                        var strUrlLink = strUrl.Replace("//", "https://"); ;

                        string strsql = "insert into mySteel_GrabData(id,title,url,content,webdate) "
                            + " values('" + strId + "','" + strTitle + "','" + strUrlLink + "','" + strContent.Replace("'", "''").Trim() + "','" + strWebDate + "')";
                        //判断是否存在重复记录
                        string strsql2 = "select * from mySteel_GrabData where id='" + strId + "'";
                        DataTable dt2 = GrabSteel.Common.SqlData.SelectDataTable("", strsql2);
                        if (dt2.Rows.Count == 0)
                        {
                            GrabSteel.Common.SqlData.InsDelUpdData("", strsql);
                        }
                    }
                }
                QueryAllData();//查询看看是否新增了数据
            }
        }
        /// <summary>
        /// 窗口关闭 停止定时器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            mytimer.Enabled = false;
            mytimer.Stop();
        }

        /// <summary>
        /// 定时服务开启/关闭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ck_Auto_CheckedChanged(object sender, EventArgs e)
        {
            if (ck_Auto.Checked == true)
            {
                mytimer.Enabled = true;
                mytimer.Start();
            }
            else
            {
                mytimer.Enabled = false;
                mytimer.Stop();
            }
        }

        private void pager1_OnPageChanged(object sender, EventArgs e)
        {
            QueryAllData();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "链接" && e.RowIndex >= 0)
            {
                DataGridViewColumn column = dataGridView1.Columns[e.ColumnIndex];
                int row = this.dataGridView1.CurrentRow.Index;
                string projectPath = dataGridView1.Rows[row].Cells["链接"].Value.ToString();
                System.Diagnostics.Process.Start(projectPath);
            }
        }
        #region my-code-Center
        public static string HttpPost(string Url, string postDataStr, ref CookieContainer cookie)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
            request.Method = "POST";
            request.ContentType = "application/json;charset=UTF-8";
            //如果服务器返回错误，他还会继续再去请求，不会使用之前的错误数据，做返回数据
            request.ServicePoint.Expect100Continue = false;

            byte[] postData = Encoding.UTF8.GetBytes(postDataStr);
            request.ContentLength = postData.Length;
            request.CookieContainer = cookie;
            Stream myRequestStream = request.GetRequestStream();
            myRequestStream.Write(postData, 0, postData.Length);
            myRequestStream.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            response.Cookies = cookie.GetCookies(response.ResponseUri);
            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();

            return retString;
        }
        public static string HttpGet(string Url, string postDataStr, CookieContainer cookie)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url + (postDataStr == "" ? "" : "?") + postDataStr);

            request.Method = "GET";
            request.ContentType = "text/html;charset=UTF-8";
            request.CookieContainer = cookie;
            //如果服务器返回错误，他还会继续再去请求，不会使用之前的错误数据，做返回数据
            request.ServicePoint.Expect100Continue = false;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("gb2312"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();

            return retString;
        }
        static string Post(string url,CookieContainer cookie)
        {
            var newUrl = url.Replace("//", "https://");

            string _innerHTML = HttpGet(newUrl, "", cookie);

            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(_innerHTML);

            var htmlNode1 = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='text']");
            var htmlNode2 = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='text']/p[position()=2]");
            htmlNode1.RemoveChild(htmlNode2);
            return htmlNode1.InnerHtml;
        }

        static string Post2(string url, CookieContainer cookie)
        {
            var newUrl = url.Replace("//", "https://");

            string _innerHTML = HttpGet(newUrl, "", cookie);

            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(_innerHTML);

            var nodes = htmlDoc.DocumentNode.SelectNodes("//*");
            if (nodes != null)
            {
                nodes.ToList().ForEach(node => {
                    // 筛选属性
                    node.Attributes.ToList().ForEach(attr => {
                        // 移除脚本属性onclick
                        if (attr.Name.StartsWith("on")) attr.Remove();
                        // 移除外部链接
                        if (node.Name == "a" && attr.Name == "href") attr.Remove();
                        // 移除特定节点
                        if (attr.Value == "select_box") node.Remove();
                        if (attr.Value == "selectCandition") node.Remove();
                    });

                });
            }

            var htmlNode1 = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='text22222']");
            var removeNode1 = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='text22222']/div[@id='company-rec']");
            var removeNode2 = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='text22222']/p[position()=1]");
            var removeNode3 = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='text22222']/p[position()=2]");
            var removeNode4 = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='text22222']/p[position()=3]");
            htmlNode1.RemoveChild(removeNode1);
            htmlNode1.RemoveChild(removeNode2);
            htmlNode1.RemoveChild(removeNode3);
            htmlNode1.RemoveChild(removeNode4);
            return htmlNode1.InnerHtml;
        }

        public static string SendRequest(string url, string method, string auth, string reqParams)
        {
            //这是发送Http请求的函数，可根据自己内部的写法改造
            HttpWebRequest myReq = null;
            HttpWebResponse response = null;
            string result = string.Empty;
            try
            {
                myReq = (HttpWebRequest)WebRequest.Create(url);
                myReq.Method = method;
                myReq.ContentType = "application/json;";
                myReq.KeepAlive = false;

                //basic 验证下面这句话不能少
                if (!String.IsNullOrEmpty(auth))
                {
                    myReq.Headers.Add("Authorization", "Basic " + auth);

                    myReq.Headers.Add("token", "4252714");
                    //myReq.Headers.Add("Cookie", "Hm_lvt_1c4432afacfa2301369a5625795031b8=1617700884,1617765296,1617775143; qimo_xstKeywords_5d36a9e0-919c-11e9-903c-ab24dbab411b=; qimo_seokeywords_5d36a9e0-919c-11e9-903c-ab24dbab411b=%E6%9C%AA%E7%9F%A5; qimo_seosource_5d36a9e0-919c-11e9-903c-ab24dbab411b=%E5%85%B6%E4%BB%96%E7%BD%91%E7%AB%99; href=https%3A%2F%2Fsearch.mysteel.com%2Fprice.html%3Fkw%3D%25E9%2592%25A2%25E6%259D%2590%26st%3D%26et%3D; accessId=5d36a9e0-919c-11e9-903c-ab24dbab411b; marketPriceHistory=%5B%22%E6%88%90%E9%83%BD%E5%B8%82%E9%92%A2%E6%9D%90%22%2C%22%E9%92%A2%E6%9D%90%22%5D; _last_loginuname=Jason113; _rememberStatus=false; _login_token=cbb9484ed9ca2dd621a9bbbb61959d8b; _login_uid=6219744; _login_mid=6865482; _login_ip=218.89.234.232; cbb9484ed9ca2dd621a9bbbb61959d8b=33%3D5%2622%3D5%2611%3D5%2634%3D5%2635%3D5%2613%3D5%2636%3D5%2637%3D5%2638%3D5%261%3D5%262%3D5%264%3D10%2640%3D5%2641%3D5%2642%3D5%2631%3D5%2632%3D5%26catalog%3D020105%2C020206%2C0204%2C0222%2C0223%2C0205%2C0209%2C1001%2C1007%2C010205%2C010202%2C0203%2C0220%2C02%2C1006%2C0202%2C0201%2C0223%2C0205%2C0222%2C0213%2C1002; openChat5d36a9e0-919c-11e9-903c-ab24dbab411b=true; _last_ch_r_t=1617779683846; Hm_lpvt_1c4432afacfa2301369a5625795031b8=1617779684; pageViewNum=5");
                }

                if (method == "POST" || method == "PUT")
                {
                    byte[] bs = Encoding.UTF8.GetBytes(reqParams);
                    myReq.ContentLength = bs.Length;
                    using (Stream reqStream = myReq.GetRequestStream())
                    {
                        reqStream.Write(bs, 0, bs.Length);
                        reqStream.Close();
                    }
                }

                response = (HttpWebResponse)myReq.GetResponse();
                HttpStatusCode statusCode = response.StatusCode;
                if (Equals(response.StatusCode, HttpStatusCode.OK))
                {
                    using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        result = reader.ReadToEnd();
                    }
                }
            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpStatusCode errorCode = ((HttpWebResponse)e.Response).StatusCode;
                    string statusDescription = ((HttpWebResponse)e.Response).StatusDescription;
                    using (StreamReader sr = new StreamReader(((HttpWebResponse)e.Response).GetResponseStream(), Encoding.UTF8))
                    {
                        result = sr.ReadToEnd();
                    }
                }
                else
                {
                    result = e.Message;
                }
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
                if (myReq != null)
                {
                    myReq.Abort();
                }
            }

            return result;
        }

        public static string sendHttpRequest(string url, string reqparam)
        {
            string auth = Base64Encode("key:secret");
            return SendRequest(url, "POST", auth, reqparam);
        }

        private static string Base64Encode(string value)
        {
            byte[] bytes = Encoding.Default.GetBytes(value);
            return Convert.ToBase64String(bytes);
        }
        public static void AddLogToTXT(string logstring)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory.Replace("\\bin\\Debug", "\\") + "log\\" + "operaLog" + DateTime.Now.ToString("HHmmss") + ".txt";
            if (!System.IO.File.Exists(path))
            {
                FileStream stream = System.IO.File.Create(path);
                stream.Close();
                stream.Dispose();
            }
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(logstring);
            }
        }
        public static string ReplaceHtmlTag(string html, int length = 0)
        {
            string strText = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", "");
            strText = System.Text.RegularExpressions.Regex.Replace(strText, "&[^;]+;", "");

            if (length > 0 && strText.Length > length)
                return strText.Substring(0, length);

            return strText;
        }
        #endregion
    }
}
