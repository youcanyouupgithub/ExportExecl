
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Model
{

    /// <summary>
    /// 菜品导入
    /// </summary>
    public class E_Impdish: E_BaseModel
    {

        /// <summary>
        /// 菜品名称
        /// </summary>
        public string dishname { get; set; }

        /// <summary>
        /// 菜系
        /// </summary>
        public string caix { get; set; }

        /// <summary>
        /// 味型
        /// </summary>
        public string weix { get; set; }

        /// <summary>
        /// 地质
        /// </summary>
        public string diz { get; set; }

        /// <summary>
        /// 烹饪技法
        /// </summary>
        public string prjf { get; set; }

        /// <summary>
        /// 图片
        /// </summary>
        public string pic { get; set; }

        /// <summary>
        /// 主料
        /// </summary>
        public string zhul { get; set; }

        /// <summary>
        /// 辅料
        /// </summary>
        public string ful { get; set; }

        /// <summary>
        /// 调料
        /// </summary>
        public string tiaol { get; set; }

        /// <summary>
        /// 烹饪技法
        /// </summary>
        public string pengrjf { get; set; }

        /// <summary>
        /// 技术要点
        /// </summary>
        public string jishuyd { get; set; }

        /// <summary>
        /// 图片路径
        /// </summary>
        public string newpic { get; set; }

    }
}

