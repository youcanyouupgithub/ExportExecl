using HtmlTableToExecl;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.IO;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;

namespace ExportExeclWeb
{
    public partial class ExportExecl : Page
    {
        //public XWPFParagraph SetCellText(XWPFDocument doc, XWPFTable table, string setText, ParagraphAlignment align, int textPos)
        //{
        //    CT_P para = new CT_P();
        //    XWPFParagraph pCell = new XWPFParagraph(para, table.Body);
        //    //pCell.Alignment = ParagraphAlignment.LEFT;//字体  
        //    pCell.Alignment = align;
        //    XWPFRun r1c1 = pCell.CreateRun();
        //    r1c1.SetText(setText);
        //    //r1c1.FontSize = 12;
        //    //r1c1.SetFontFamily("华文楷体", FontCharRange.None);//设置雅黑字体  
        //    //r1c1.SetTextPosition(textPos);//设置高度  
        //    return pCell;
        //}

        protected void Page_Load(object sender, EventArgs e)
        {
            //string filepath = Server.MapPath("/upfile/分样单模板.docx");
            //using (FileStream stream = File.OpenRead(filepath))
            //{
            //    XWPFDocument doc = new XWPFDocument(stream);
            //    //遍历表格
            //    var table = doc.Tables[0];
            //    for (int i = 2; i < 20; i++)
            //    {
            //        if (table.GetRow(i) == null)
            //        {
            //            CT_Row row = new CT_Row();
            //            row.AddNewTc();
            //            row.AddNewTc();
            //            row.AddNewTc();
            //            row.AddNewTc();
            //            row.AddNewTc();
            //            row.AddNewTc();
            //            row.AddNewTc();
            //            table.AddRow(new XWPFTableRow(row, table));
            //        }
            //        table.GetRow(i).GetCell(0).SetParagraph(SetCellText(doc, table, "附件：", ParagraphAlignment.CENTER, 80));
            //        //table.GetRow(i).GetCell(0).SetText("1");
            //        table.GetRow(i).GetCell(1).SetText("2");
            //        table.GetRow(i).GetCell(2).SetText("3");
            //        table.GetRow(i).GetCell(3).SetText("4");
            //        table.GetRow(i).GetCell(4).SetText("5");
            //        table.GetRow(i).GetCell(5).SetText("6");
            //        table.GetRow(i).GetCell(6).SetText("7");
            //    }

            //    using (MemoryStream ms = new MemoryStream())
            //    {
            //        doc.Write(ms);
            //        Response.ContentType = "application/vnd.ms-word";
            //        Response.ContentEncoding = Encoding.UTF8;
            //        Response.Charset = "";
            //        Response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode("分样单20170724.docx", Encoding.UTF8));
            //        Response.BinaryWrite(ms.GetBuffer());
            //        Response.End();
            //    }
            //}
        }

        protected void Btn_Export_Click(object sender, EventArgs e)
        {
            string value = this.HidHtmlTableJson.Value;
            HtmlTableExport.RenderToExcel(value, base.Server.MapPath("~/upfile/数据文档"+DateTime.Now.ToString("yyyyMMddhhmmss")+".xls"));
        }


    }
}
