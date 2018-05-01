using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Comp;
using Dapper;
using Model;
using System.Data;
using System.Data.SqlClient;

namespace DAL
{
    /// <summary>
    /// 菜品导入
    /// </summary>
    public class D_Impdish
    {
        /// <summary>
        /// 添加
        /// </summary>
        public int Add(E_Impdish model)
        {
            string sql = @"INSERT INTO tb_impdish
                ([dishname],[caix],[weix],[diz],[prjf],[pic],[zhul],[ful],[tiaol],[pengrjf],[jishuyd],[newpic]) 
                VALUES 
                (@dishname,@caix,@weix,@diz,@prjf,@pic,@zhul,@ful,@tiaol,@pengrjf,@jishuyd,@newpic);select @@IDENTITY;";
            using (IDbConnection conn = new SqlConnection(DapperHelper.GetConStr()))
            {
                return Convert.ToInt32(conn.ExecuteScalar(sql, model));
            }
        }
    }
}

