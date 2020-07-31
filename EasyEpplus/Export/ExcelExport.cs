using EasyEpplus.Common;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EasyEpplus.Export
{
    /// <summary>
    /// Excel文件导出实现类
    /// <para name = "author">Peihui</para>
    /// <para name = "QQ">129303542</para>
    /// </summary>
    public class ExcelExport : IExcelExport
    {
        #region 设置成员变量
        /// <summary>
        /// 工作簿字典
        /// </summary>
        public Dictionary<string, string> excelDistionary;
        /// <summary>
        /// 选择字典
        /// </summary>
        public Dictionary<int, string> selectDistionary;
        /// <summary>
        /// 开始行
        /// </summary>
        public int BeginRow;
        #endregion
        public ExcelExport()
        {
            BeginRow = 1;
        }

        public byte[] ExportExcelByAttribute<T>(List<T> data, string sheetName = "") where T : class
        {
            #region 定义变量
            DataTable dataTable = EPPlusHelper.ListToDataTable(data);
            byte[] result;

            #endregion
            using (ExcelPackage package = new ExcelPackage())
            {
                this.GetClassAttribute<T>();
                // 设置工作簿
                ExcelWorksheet workSheet = this.SetWorksheet(package, dataTable.Columns.Count, sheetName);
                workSheet = this.SetColumnAttribute<T>(workSheet, dataTable);
                //赋值
                workSheet.Cells[string.Format("A{0}", BeginRow + 1)].LoadFromDataTable(dataTable, false);

                // 清除非选中列
                foreach (var key in this.selectDistionary.Keys)
                {
                    if (!Convert.ToBoolean(this.selectDistionary.FirstOrDefault(v => v.Key == key).Value))
                    {
                        workSheet.DeleteColumn(key);
                    }
                }
                result = package.GetAsByteArray();
            }
            return result;
        }
        /// <summary>
        /// 设置列特性
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        private ExcelWorksheet SetColumnAttribute<T>(ExcelWorksheet workSheet, DataTable dataTable) where T : class
        {
            selectDistionary = new Dictionary<int, string>();
            // 循环标识
            int i = 1;
            foreach (PropertyInfo p in typeof(T).GetProperties())
            {
                var attributes = p.CustomAttributes.ToArray()[0].NamedArguments;
                foreach (var item in attributes)
                {
                    #region 列特性
                    switch (item.MemberName)
                    {
                        case "ColumnName":
                            workSheet.Cells[1, i].Value = item.TypedValue.Value.ToString();
                            break;
                        case "Format":
                            workSheet.Cells[2, i, dataTable.Rows.Count + 1, i].Style.Numberformat.Format = item.TypedValue.Value.ToString();
                            break;
                        case "FontSize":
                            workSheet.Cells[2, i, dataTable.Rows.Count + 1, i].Style.Font.Size = Convert.ToSingle(item.TypedValue.Value.ToString());
                            break;
                        case "FontColor":
                            workSheet.Cells[2, i, dataTable.Rows.Count + 1, i].Style.Font.Color.SetColor(ColorTranslator.FromHtml(item.TypedValue.Value.ToString()));
                            break;
                        case "IsBold":
                            workSheet.Cells[2, i, dataTable.Rows.Count + 1, i].Style.Font.Bold = Convert.ToBoolean(item.TypedValue.Value.ToString());
                            break;
                        case "IsAutoFit":
                            if (Convert.ToBoolean(excelDistionary.FirstOrDefault(v => v.Key == "IsAutoFit").Value))
                                workSheet.Column(i).AutoFit();
                            break;
                        case "IsSelect":
                            selectDistionary.Add(i, item.TypedValue.Value.ToString());
                            break;
                    }
                    #endregion
                }
                i++;
            }
            return workSheet;
        }

        /// <summary>
        /// 设置工作簿基本样式内容
        /// </summary>
        /// <param name="package"></param>
        /// <param name="columnCount"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private ExcelWorksheet SetWorksheet(ExcelPackage package, int columnCount, string sheetName)
        {
            string SheetName = "Sheet1";
            string HeaderColor = "#E0FFFF";
            string FontColor = "#000000";
            if (excelDistionary.ContainsKey("SheetName"))
            {
                SheetName = this.excelDistionary.FirstOrDefault(v => v.Key == "SheetName").Value;
            }
            ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(sheetName == "" ? SheetName : sheetName);
            #region 设置表头样式
            using (ExcelRange r = workSheet.Cells[BeginRow, 1, BeginRow, columnCount])
            {
                if (excelDistionary.ContainsKey("HeaderColor"))
                {
                    HeaderColor = this.excelDistionary.FirstOrDefault(v => v.Key == "HeaderColor").Value;
                }
                if (excelDistionary.ContainsKey("FontColor"))
                {
                    FontColor = this.excelDistionary.FirstOrDefault(v => v.Key == "FontColor").Value;
                }
                if (excelDistionary.ContainsKey("FontColor"))
                {
                    FontColor = this.excelDistionary.FirstOrDefault(v => v.Key == "FontColor").Value;
                }
                if (excelDistionary.ContainsKey("FontSize"))
                {
                    r.Style.Font.Size = Convert.ToSingle(this.excelDistionary.FirstOrDefault(v => v.Key == "FontSize").Value);
                }
                r.Style.Font.Color.SetColor(ColorTranslator.FromHtml(FontColor));
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                // 设置表头颜色
                r.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(HeaderColor));
            }
            #endregion

            #region 设置数据单元格样式
            using (ExcelRange r = workSheet.Cells[BeginRow + 1, 1, BeginRow + columnCount, columnCount])
            {
                r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                r.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                r.Style.Border.Top.Color.SetColor(Color.Black);
                r.Style.Border.Bottom.Color.SetColor(Color.Black);
                r.Style.Border.Left.Color.SetColor(Color.Black);
                r.Style.Border.Right.Color.SetColor(Color.Black);
            }
            #endregion

            return workSheet;
        }


        /// <summary>
        /// 获取Excel特性并存为字典
        /// </summary>
        /// <typeparam name="T"></typeparam>
        private void GetClassAttribute<T>() where T : class
        {
            this.excelDistionary = new Dictionary<string, string>();
            var attributes = typeof(T).CustomAttributes.ToArray()[0].NamedArguments;
            foreach (var item in attributes)
            {
                excelDistionary.Add(item.MemberName, item.TypedValue.Value.ToString());
            }
        }
    }
}
