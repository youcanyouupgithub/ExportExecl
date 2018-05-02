using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using DAL;
using Model;
using NPOI.XWPF.UserModel;

namespace HtmlTableToExecl
{
    /// <summary>
    /// Word处理工具
    /// </summary>
    public class WordHelper
    {
        /// <summary>
        /// 读取表
        /// </summary>
        /// <returns></returns>
        public static string ExcuteWord(string docpath, string imgpath)
        {
            D_Impdish dImpdish = new D_Impdish();


            StringBuilder sb = new StringBuilder();
            using (FileStream stream = File.OpenRead(docpath))
            {
                XWPFDocument doc = new XWPFDocument(stream);
                var tables = doc.Tables;
                foreach (var table in tables)    //遍历表格  
                {
                    E_Impdish eImpdish = new E_Impdish();
                    if (docpath.IndexOf("wan/9.docx") > -1)
                    {
                        eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText; //菜品名称
                        eImpdish.pic = table.Rows[0].GetCell(2).Paragraphs[0].ParagraphText.Replace("\t", "");      //图片
                        eImpdish.caix = table.Rows[1].GetCell(1).Paragraphs[0].ParagraphText;     //菜系
                        eImpdish.weix = table.Rows[2].GetCell(1).Paragraphs[0].ParagraphText;     //味型
                        eImpdish.diz = "";//table.Rows[3].GetCell(1).Paragraphs[0].ParagraphText;      //地质
                        eImpdish.prjf = "";// table.Rows[4].GetCell(1).Paragraphs[0].ParagraphText;     //烹饪技法
                        eImpdish.zhul = GetAllParagraphText(table.Rows[4].GetCell(1).Paragraphs);     //主料
                        eImpdish.ful = GetAllParagraphText(table.Rows[5].GetCell(1).Paragraphs);      //辅料
                        eImpdish.tiaol = GetAllParagraphText(table.Rows[6].GetCell(1).Paragraphs);    //调料
                        eImpdish.pengrjf = GetAllParagraphText(table.Rows[10].GetCell(1).Paragraphs);  //烹饪方法

                        eImpdish.jishuyd = "";//table.Rows[9].GetCell(1).Paragraphs[0].ParagraphText;  //技术要点
                    }
                    else
                    {
                        eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText; //菜品名称
                        eImpdish.pic = table.Rows[0].GetCell(2).Paragraphs[0].ParagraphText.Replace("\t", "");      //图片
                        eImpdish.caix = table.Rows[1].GetCell(1).Paragraphs[0].ParagraphText;     //菜系
                        eImpdish.weix = table.Rows[2].GetCell(1).Paragraphs[0].ParagraphText;     //味型
                        eImpdish.diz = table.Rows[3].GetCell(1).Paragraphs[0].ParagraphText;      //地质
                        eImpdish.prjf = table.Rows[4].GetCell(1).Paragraphs[0].ParagraphText;     //烹饪技法

                        eImpdish.zhul = GetAllParagraphText(table.Rows[5].GetCell(1).Paragraphs);     //主料
                        eImpdish.ful = GetAllParagraphText(table.Rows[6].GetCell(1).Paragraphs);      //辅料
                        eImpdish.tiaol = GetAllParagraphText(table.Rows[7].GetCell(1).Paragraphs);    //调料
                        eImpdish.pengrjf = GetAllParagraphText(table.Rows[8].GetCell(1).Paragraphs);  //烹饪方法
                        eImpdish.jishuyd = GetAllParagraphText(table.Rows[9].GetCell(1).Paragraphs);  //技术要点
                    }

                    //查找对应图片，并进行拷贝重命名
                    var imgname = Guid.NewGuid();

                    //JPG
                    var img = imgpath + eImpdish.pic + ".jpg";
                    var newimgpath = @"D:/dish/upload/" + imgname + ".jpg";
                    FileInfo file = new FileInfo(img);
                    if (file.Exists)
                    {
                        file.CopyTo(newimgpath, true);
                        eImpdish.newpic = "{%upimgpath%}/upload/20180501/" + imgname + ".jpg";
                    }

                    //PNG
                    img = imgpath + eImpdish.pic + ".png";
                    newimgpath = @"D:/dish/upload/" + imgname + ".png";
                    file = new FileInfo(img);
                    if (file.Exists)
                    {
                        file.CopyTo(newimgpath, true);
                        eImpdish.newpic = "{%upimgpath%}/upload/20180501/" + imgname + ".png";
                    }

                    var id = dImpdish.Add(eImpdish);
                    if (id > 0)
                    {
                        sb.Append($"OK：{eImpdish.dishname}</br>");
                    }
                    else
                    {
                        sb.Append($"NO：{eImpdish.dishname}=》添加失败</br>");
                    }
                }
                return sb.ToString();
            }
        }

        /// <summary>
        /// 获取所有段落文本
        /// </summary>
        public static string GetAllParagraphText(IList<XWPFParagraph> Paragraphs)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var item in Paragraphs)
            {
                sb.AppendLine(item.ParagraphText);
            }
            return sb.ToString();
        }

        /// <summary>
        /// 读取段落
        /// </summary>
        /// <returns></returns>
        public static string ExcuteWordText()
        {
            StringBuilder sb = new StringBuilder();
            using (FileStream stream = File.OpenRead("d:/test.docx"))
            {
                XWPFDocument doc = new XWPFDocument(stream);
                foreach (var para in doc.Paragraphs)
                {
                    string text = para.ParagraphText; //获得文本  

                    var runs = para.Runs;
                    string styleid = para.Style;

                    for (int i = 0; i < runs.Count; i++)
                    {
                        var run = runs[i];
                        text = run.ToString(); //获得run的文本  
                        sb.Append(text + ",");
                    }
                }
            }
            return sb.ToString();
        }

    }
}
