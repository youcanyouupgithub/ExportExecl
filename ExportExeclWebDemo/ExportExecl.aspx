<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExportExecl.aspx.cs" Inherits="ExportExeclWeb.ExportExecl" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <style type="text/css">
        * {
            padding: 0px;
            margin-left: 0px;
            margin-right: 0px;
            margin-top: 0px;
        }
    </style>
    <script src="Js/jquery-1.9.1.js" type="text/javascript"></script>
    <script src="Js/HtmlTableToExecl.js"></script>
    <script type="text/javascript">
        function GetHtmlTableJson(TableID) {
            var HtmlTableJson = "";
            HtmlTableJson += "[";

            HtmlTableJson += GetTrJson(TableID,0);

            HtmlTableJson += "," + GetTrJson(TableID, $("#" + TableID+" tr").length+1); //Table 2

            HtmlTableJson += "]";
            $("#HtmlTableJson").html(HtmlTableJson);
            $("#HidHtmlTableJson").val(HtmlTableJson);
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:HiddenField ID="HidHtmlTableJson" runat="server" Value="" />
        <div id="HtmlTableJson">
        </div>
        <div style="margin-top: 20px; width: 300px;">
            <table border="1" cellpadding="0" cellspacing="0" id="Tab_1" width="100%">
                <thead>
                    <tr>
                        <th colspan="3">合并单元格的表头</th>
                    </tr>
                    <tr>
                        <th>序号</th>
                        <th>标题</th>
                        <th>内容</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>2</td>
                        <td>测试2</td>
                        <td>来一个</td>
                    </tr>
                    <tr>
                        <td>1</td>
                        <td colspan="2">测试1</td>
                    </tr>
                    <tr>
                        <td rowspan="2">2</td>
                        <td>测试1231</td>
                        <td>测123123试2</td>
                    </tr>
                    <tr>
                        <td colspan="2">aa山东分公司地方a11</td>
                    </tr>
                    <tr>
                        <td>5</td>
                        <td rowspan="2">12</td>
                        <td>测试</td>
                    </tr>
                    <tr>
                        <td>5</td>
                        <td>测试</td>
                    </tr>
                    <tr>
                        <td>5</td>
                        <td>测试</td>
                        <td rowspan="3">测试</td>
                    </tr>
                    <tr>
                        <td colspan="2">5</td>
                    </tr>
                    <tr>
                        <td>5</td>
                        <td>6</td>
                    </tr>
                </tbody>
            </table>
            <br />
            第一步:点击该按钮生成一下json<br />
            <input type="button" id="Export" onclick="GetHtmlTableJson('Tab_1')" value="导出Execl" />
            <br />
            第二步：导出Execl（注：导出路径可以自己在后台修改）<br />
            <asp:Button ID="Btn_Export" runat="server" Text="导出Execl" OnClick="Btn_Export_Click" />
            <br />
            <br />
            <br />
            <span style="color: Red;">注：嵌套表格导出还在开发中...
        <br />
                支持：合并单元格、合并行等复杂表格导出。
            </span>
        </div>
    </form>
</body>
</html>
