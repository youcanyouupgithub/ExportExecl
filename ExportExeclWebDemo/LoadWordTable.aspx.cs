using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using HtmlTableToExecl;

namespace ExportExeclWebDemo
{
    public partial class LoadWordTable : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            List<string> doclist=new List<string>();
            doclist.AddRange(new List<string>() {
                "D:/dish/doc/shang/1.docx", "D:/dish/doc/shang/2.docx", "D:/dish/doc/shang/3.docx",
                "D:/dish/doc/shang/4.docx", "D:/dish/doc/shang/5.docx", "D:/dish/doc/shang/6.docx" });
            doclist.AddRange(new List<string>() {
                "D:/dish/doc/xia/1.docx", "D:/dish/doc/xia/2.docx", "D:/dish/doc/xia/3.docx" });
            doclist.AddRange(new List<string>() {
                "D:/dish/doc/wan/1.docx", "D:/dish/doc/wan/2.docx", "D:/dish/doc/wan/3.docx",
                "D:/dish/doc/wan/4.docx", "D:/dish/doc/wan/5.docx", "D:/dish/doc/wan/6.docx",
                "D:/dish/doc/wan/7.docx", "D:/dish/doc/wan/8.docx", "D:/dish/doc/wan/9.docx"});


            List<string> imglist = new List<string>();
            imglist.AddRange(new List<string>() {
                "D:/dish/images/shang/1/", "D:/dish/images/shang/2/", "D:/dish/images/shang/3/",
                "D:/dish/images/shang/4/", "D:/dish/images/shang/5/", "D:/dish/images/shang/6/" });
            imglist.AddRange(new List<string>() {
                "D:/dish/images/xia/1/", "D:/dish/images/xia/2/", "D:/dish/images/xia/3/" });
            imglist.AddRange(new List<string>() {
                "D:/dish/images/wan/1/", "D:/dish/images/wan/2/", "D:/dish/images/wan/3/",
                "D:/dish/images/wan/4/", "D:/dish/images/wan/5/", "D:/dish/images/wan/6/",
                "D:/dish/images/wan/7/", "D:/dish/images/wan/8/", "D:/dish/images/wan/9/" });
            
            for (int i = 0; i < doclist.Count; i++)
            {
                //导入菜谱
                string msg = WordHelper.ExcuteWord(doclist[i], imglist[i]);
                Response.Write(msg);
                Response.Write($"</br>{doclist[i]}");
            }

            Response.End();
        }
    }
}