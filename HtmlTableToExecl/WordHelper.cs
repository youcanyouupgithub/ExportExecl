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
        public static string ExcuteWord(string strfilepath)
        {
            D_Impdish dImpdish=new D_Impdish();


            StringBuilder sb = new StringBuilder();
            using (FileStream stream = File.OpenRead(strfilepath))
            {
                XWPFDocument doc = new XWPFDocument(stream);
                var tables = doc.Tables;
                foreach (var table in tables)    //遍历表格  
                {
                    E_Impdish eImpdish = new E_Impdish();

                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;
                    eImpdish.pic = table.Rows[0].GetCell(2).Paragraphs[0].ParagraphText;
                    eImpdish.caix = table.Rows[1].GetCell(1).Paragraphs[0].ParagraphText;

                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;
                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;
                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;
                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;
                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;
                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;
                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;
                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;
                    eImpdish.dishname = table.Rows[0].GetCell(1).Paragraphs[0].ParagraphText;

                    dImpdish.Add(eImpdish);
                    /*
                    foreach (var row in table.Rows)    //遍历行  
                    {
                        var c0 = row.GetCell(0);        //获得单元格0  
                        foreach (var para in c0.Paragraphs)
                        {
                            string text = para.ParagraphText;
                            //处理段落      
                            sb.Append(text + ",");
                        }
                    }
                    */
                }
                return sb.ToString();
            }
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
