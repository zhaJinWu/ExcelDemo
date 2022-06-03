using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Advanced.Common
{
    public class ExcelDataResource
    {
        /// <summary>
        /// Sheet页名称
        /// </summary>
        public string SheetName { get; set; }
        /// <summary>
        /// 标题所在行
        /// </summary>
        public int HeadIndex { get; set; }
        /// <summary>
        /// 每一行sheet的数据
        /// </summary>
        public List<object> SheetDataResource { get; set; }
    }

    public class UseInfo
    {
        [Title(Title = "用户Id")]
        public int UseId { get; set; }

        [Title(Title = "用户名称")]
        public string UseName { get; set; }

        [Title(Title = "用户年龄")]
        public int Age { get; set; }

        [Title(Title = "用户类型")]
        public int UserType { get; set; }

        [Title(Title = "描述")]
        public string Description { get; set; }
    }
}
