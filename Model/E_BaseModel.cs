using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// 实体公共属性
    /// </summary>
    public class E_BaseModel
    {
        /// <summary>
        /// 页面索引
        /// </summary>
        public int? pageindex { get; set; }

        /// <summary>
        /// 页面数据行数
        /// </summary>
        public int pagesize { get; set; }

        /// <summary>
        /// 是否查询count
        /// </summary>
        public bool iscountsearch { get; set; } = false;

        /// <summary>
        /// 自动生成序号
        /// </summary>
        public int rowid { get; set; }

        /// <summary>
        /// 开始日期
        /// </summary>
        public DateTime? starttime { get; set; }

        /// <summary>
        /// 结束日期
        /// </summary>
        public DateTime? endtime { get; set; }

    }
}
