using System;
using System.Collections.Generic;
using System.Text;

namespace EasyEpplus.MyAttribute
{
    /// <summary>
    /// Excel文件特性
    /// <para name = "author">Peihui</para>
    /// <para name = "QQ">129303542</para>
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelExportAttribute : Attribute
    {
        /// <summary>
        /// Sheet名称
        /// </summary>
        public string SheetName { get; set; } = "Sheet1";
        /// <summary>
        /// 设置表头背景颜色,格式如：#E0FFFF
        /// </summary>
        public string HeaderColor { get; set; } = "#E0FFFF";
        /// <summary>
        /// 最大的行数
        /// </summary>
        public int MaxRowNumberOnASheet { get; set; }
        /// <summary>
        /// 字体，如幼圆
        /// </summary>
        public string TableStyle { get; set; }
        /// <summary>
        /// 字体颜色
        /// </summary>
        public string FontColor { get; set; } = "#000000";
        /// <summary>
        /// 字体大小
        /// </summary>
        public float? FontSize { set; get; }
        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; }
        /// <summary>
        /// 表头是否加粗
        /// </summary>
        public bool IsHeadBold { set; get; } = false;
    }
}
