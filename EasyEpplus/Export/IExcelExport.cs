using System;
using System.Collections.Generic;
using System.Text;

namespace EasyEpplus.Export
{
    /// <summary>
    /// Excel文件导出接口
    /// <para name = "author">Peihui</para>
    /// <para name = "QQ">129303542</para>
    /// </summary>
    public interface IExcelExport
    {
        /// <summary>
        /// 根据特性导出Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">数据</param>
        /// <param name="sheetName">工作簿名</param>
        /// <returns></returns>
        byte[] ExportExcelByAttribute<T>(List<T> data, string sheetName = "") where T : class;
    }
}
