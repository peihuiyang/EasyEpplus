using EasyEpplus.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EasyEpplus.Import
{
    /// <summary>
    /// Excel文件导入实现类
    /// <para name = "author">Peihui</para>
    /// <para name = "QQ">129303542</para>
    /// </summary>
    public class ExcelImport : IExcelImport
    {
        public List<T> ImportExcelToList<T>(Stream stream) where T : new()
        {
            var dataTable = EPPlusHelper.WorksheetToTable(stream);
            List<T> dataList = EPPlusHelper.GetDtoFromDataTable<T>(dataTable);
            return dataList;
        }
    }
}
