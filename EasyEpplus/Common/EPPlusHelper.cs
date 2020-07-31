using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;

namespace EasyEpplus.Common
{
    /// <summary>
    /// 使用EPPlus操作Excel
    /// <para name = "author">Peihui</para>
    /// <para name = "QQ">129303542</para>
    /// </summary>
    public class EPPlusHelper
    {
        private static string GetString(object obj)
        {
            try
            {
                return obj.ToString();
            }
            catch (Exception ex)
            {
                ex.Message.ToString();
                return "";
            }
        }
        /// <summary>
        /// 将指定的Excel的文件转换成DataTable （Excel的第一个sheet）
        /// </summary>
        /// <param name="fullFielPath"></param>
        /// <returns></returns>
        public static DataTable WorksheetToTable(string fullFielPath)
        {
            try
            {
                FileInfo existingFile = new FileInfo(fullFielPath);

                ExcelPackage package = new ExcelPackage(existingFile);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];//选定 指定页

                return WorksheetToTable(worksheet);
            }
            catch (Exception)
            {
                throw;
            }
        }
        /// <summary>
        /// 将指定Excel文件流转换为DataTable
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static DataTable WorksheetToTable(Stream stream)
        {
            try
            {
                ExcelPackage package = new ExcelPackage(stream);
                //获取指定Sheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                return WorksheetToTable(worksheet);
            }
            catch (Exception)
            {
                throw;
            }
        }
        /// <summary>
        /// 将worksheet转成datatable
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static DataTable WorksheetToTable(ExcelWorksheet worksheet)
        {
            //获取worksheet的行数
            int rows = worksheet.Dimension.End.Row;
            //获取worksheet的列数
            int cols = worksheet.Dimension.End.Column;

            DataTable dt = new DataTable(worksheet.Name);
            DataRow dr = null;
            for (int i = 1; i <= rows; i++)
            {
                if (i > 1)
                    dr = dt.Rows.Add();

                for (int j = 1; j <= cols; j++)
                {
                    //默认将第一行设置为datatable的标题
                    if (i == 1)
                        dt.Columns.Add(GetString(worksheet.Cells[i, j].Value));
                    //剩下的写入datatable
                    else
                        dr[j - 1] = GetString(worksheet.Cells[i, j].Value);
                }
            }
            return dt;
        }
        /// <summary>
        /// 将DataTable转为DTO对象
        /// </summary>
        /// <typeparam name="T">目标类型</typeparam>
        /// <param name="dataTable">原DT</param>
        /// <returns>转换后的实体列表</returns>
        public static List<T> GetDtoFromDataTable<T>(DataTable dataTable) where T : new()
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                return null;
            }
            List<T> modelList = new List<T>();
            foreach (DataRow dr in dataTable.Rows)
            {
                //T model = (T)Activator.CreateInstance(typeof(T));  
                T model = new T();
                for (int i = 0; i < dr.Table.Columns.Count; i++)
                {
                    PropertyInfo propertyInfo = model.GetType().GetProperty(dr.Table.Columns[i].ColumnName);
                    if (propertyInfo != null && dr[i] != DBNull.Value)
                        propertyInfo.SetValue(model, dr[i], null);
                }

                modelList.Add(model);
            }
            return modelList;
        }
        /// <summary>
        /// 将List数据转换为DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static DataTable ListToDataTable<T>(List<T> data) where T : class
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable dataTable = new DataTable();
            for (int i = 0; i < properties.Count; i++)
            {
                PropertyDescriptor property = properties[i];
                dataTable.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
            }
            object[] values = new object[properties.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = properties[i].GetValue(item);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }
    }
}
