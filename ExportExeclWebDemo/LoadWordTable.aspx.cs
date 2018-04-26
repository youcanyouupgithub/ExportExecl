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
            List<string> doclist=new List<string>(){ "d:/test.docx" };
            foreach (var strfilepath in doclist)
            {
                WordHelper.ExcuteWord(strfilepath);
            }

        }
    }
}