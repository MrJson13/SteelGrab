using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Web;
using System.IO;
using System.Web.Script.Serialization;

namespace GrabSteel.cofig
{
    public class config
    {
        // <summary>
        /// 获取浏览器请求参数
        /// </summary>
        /// <returns></returns>
        public static string GetBrower(string jdaddress)
        {
            Random rd = new Random();
            JavaScriptSerializer jsshd = new JavaScriptSerializer();
            jdaddress = jdaddress.Replace("\\bin\\Debug", "\\") + ("cofig\\headers.json");

            List<GrabSteel.cofig.Headers> lmhd = jsshd.Deserialize<List<GrabSteel.cofig.Headers>>(File.ReadAllText(jdaddress));//System.Web.HttpContext.Current.Server.MapPath("../json/headers.json")
            return lmhd[rd.Next(0, lmhd.Count)].UserAgent;
        }
        /// <summary>
        /// 获取请求地址
        /// </summary>
        /// <param name="jdaddress"></param>
        /// <returns></returns>
        public static string GetRequestUrl(string jdaddress)
        {
            Random rd = new Random();
            JavaScriptSerializer jsshd = new JavaScriptSerializer();
            jdaddress = jdaddress.Replace("\\bin\\Debug", "\\") + ("cofig\\link.json");
            List < GrabSteel.cofig.Link> link = jsshd.Deserialize< List<GrabSteel.cofig.Link>>(File.ReadAllText(jdaddress));
            return link[rd.Next(0, link.Count)].Url;
        }
        /// <summary>
        /// 获取设置时间
        /// </summary>
        /// <param name="jdaddress"></param>
        /// <returns></returns>
        public static string GetTime(string jdaddress)
        {
            Random rd = new Random();
            JavaScriptSerializer jsshd = new JavaScriptSerializer();
            jdaddress = jdaddress.Replace("\\bin\\Debug", "\\") + ("cofig\\link.json");
            List<GrabSteel.cofig.Link> link = jsshd.Deserialize<List<GrabSteel.cofig.Link>>(File.ReadAllText(jdaddress));
            return link[0].SetTime;
        }
        /// <summary>
        /// 计算分页开始/结束条数
        /// </summary>
        /// <param name="_nowpage">当前页</param>
        /// <param name="_perpage">每页条数</param>
        /// <returns></returns>
        public static int[] CountStartEnd(int nowpage, int perpage)
        {
            int[] _arr = new int[2];
            _arr[0] = (nowpage - 1) * perpage + 1;
            _arr[1] = _arr[0] + perpage - 1;
            return _arr;
        }
    }
}
