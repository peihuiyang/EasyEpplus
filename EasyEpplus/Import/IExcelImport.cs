using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EasyEpplus.Import
{
    /// <summary>
    /// Excel文件导入接口
    /// <para name = "author">Peihui</para>
    /// <para name = "QQ">129303542</para>
    /// </summary>
    public interface IExcelImport
    {
        /// <summary>
        /// 将Excel文件流转化为List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <returns></returns>
        List<T> ImportExcelToList<T>(Stream stream) where T : new();
    }
}
