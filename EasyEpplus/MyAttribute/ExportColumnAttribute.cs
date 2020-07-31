using System;
using System.Collections.Generic;
using System.Text;

namespace EasyEpplus.MyAttribute
{
    /// <summary>
    /// Excel导出列特性
    /// <para name = "author">Peihui</para>
    /// <para name = "QQ">129303542</para>
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportColumnAttribute : Attribute
    {
        /// <summary>
        /// 列名
        /// </summary>
        public string ColumnName { set; get; }
        /// <summary>
        /// 设置列格式
        /// </summary>
        public string Format { get; set; }
        /// <summary>
        /// 字体大小
        /// </summary>
        public float? FontSize { set; get; }
        /// <summary>
        /// 字体颜色
        /// </summary>
        public string FontColor { set; get; }
        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool IsBold { set; get; } = true;
        /// <summary>
        /// 是否自适应
        /// </summary>
        public bool IsAutoFit { set; get; } = true;
        /// <summary>
        /// 是否选择
        /// </summary>
        public bool IsSelect { get; set; } = true;
    }
}
