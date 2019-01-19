using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Framework.Common
{
    public class ExcelHelper
    {
        /// <summary>
        /// 根据模板导出excel
        /// </summary>
        /// <typeparam name="T1">表头实体类型</typeparam>
        /// <typeparam name="T2">表体实体类型</typeparam>
        /// <param name="path">模板文件路径</param>
        /// <param name="savePath">保存文件路径</param>
        /// <param name="ds">数据源</param>
        public static void ExportExcel<T1, T2>(string path,
            string savePath,
            List<SheetDataSource<T1, T2>> list) where T1 : class where T2 : class
        {
            var workbook = new XSSFWorkbook(path);

            for (int index = 0; index < list.Count; index++)
            {
                var sheetDataSource = list[index];
                string sheetName = sheetDataSource.SheetName;
                var ds = sheetDataSource.DataSource;

                var placeholderList = GetPlaceholder(path, sheetName);

                var head = ds.FieldData;

                var listData = ds.ListData;
                var listEntityType = listData.FirstOrDefault().GetType();

                var sheet = workbook.GetSheet(sheetName);

                //表头部分
                placeholderList.Where(o => o.PlaceholderType == PlaceholderType.Field).ToList().ForEach(o =>
                {
                    int rownum = o.RowIndex;
                    int cellnum = o.ColIndex;

                    //反射通过字段名称获取实体字段对应的值
                    var item = head.GetType().GetProperty(o.Placeholder, BindingFlags.Instance | BindingFlags.Public);
                    object value = item.GetValue(head, null);

                    if (value != null)
                    {
                        if (item.PropertyType.IsValueType || item.PropertyType.Name.StartsWith("String"))
                        {
                            sheet.GetRow(rownum).Cells[cellnum].SetCellValue(value.ToString());
                        }
                    }
                    else
                    {
                        sheet.GetRow(rownum).Cells[cellnum].SetCellValue("");
                    }

                });

                //表体部分
                var listPlaceholder = placeholderList.Where(o => o.PlaceholderType == PlaceholderType.List).ToList();
                int rowIndx = 0;
                for (int i = 0; i < listData.Count; i++)
                {
                    if (i == 0) //列表占位符行，列表占位符行已存在所以无需创建
                    {
                        listPlaceholder.ForEach(o =>
                        {
                            int rownum = o.RowIndex;
                            rowIndx = rownum;
                            int cellnum = o.ColIndex;

                            var item = listEntityType.GetProperty(o.Placeholder, BindingFlags.Instance | BindingFlags.Public);
                            object value = item.GetValue(listData[i], null);

                            if (value != null)
                            {
                                if (item.PropertyType.IsValueType || item.PropertyType.Name.StartsWith("String"))
                                {
                                    sheet.GetRow(rownum).Cells[cellnum].SetCellValue(value.ToString());
                                }
                            }
                            else
                            {
                                sheet.GetRow(rownum).Cells[cellnum].SetCellValue(string.Empty);
                            }

                        });
                    }
                    else//占位符之后的行进行新建
                    {
                        var row = sheet.CreateRow(rowIndx);
                        listPlaceholder.ForEach(o =>
                        {
                            int rownum = rowIndx;
                            int cellnum = o.ColIndex;

                            var item = listEntityType.GetProperty(o.Placeholder, BindingFlags.Instance | BindingFlags.Public);
                            object value = item.GetValue(listData[i], null);

                            if (value != null)
                            {
                                if (item.PropertyType.IsValueType || item.PropertyType.Name.StartsWith("String"))
                                {
                                    var cell = sheet.GetRow(rownum).CreateCell(cellnum);
                                    cell.CellStyle = sheet.GetRow(rownum - 1).Cells[cellnum].CellStyle;
                                    cell.SetCellValue(value.ToString());
                                }
                            }
                            else
                            {

                            }

                        });
                    }

                    rowIndx++;

                }

                
            }

            FileStream sw = File.Create(savePath);
            workbook.Write(sw);
            sw.Close();
        }

        /// <summary>
        /// 根据模板获取占位符信息
        /// 包括表头、表体占位符信息
        /// </summary>
        /// <param name="path">模板文件路径</param>
        /// <returns></returns>
        private static List<PlaceholderDetail> GetPlaceholder(string path, string sheetName)
        {
            List<PlaceholderDetail> placeholderList = new List<PlaceholderDetail>();

            Regex fieldRegex = new Regex(string.Format(@"{0}(\w+){0}", PlaceholderDetail.FormatDic[PlaceholderType.Field]));
            Regex listRegex = new Regex(string.Format(@"{0}(\w+){0}", PlaceholderDetail.FormatDic[PlaceholderType.List]));

            var workbook = new XSSFWorkbook(path);

            var sheet = workbook.GetSheet(sheetName);
            for (int rownum = sheet.FirstRowNum; rownum < sheet.LastRowNum + 1; rownum++)
            {
                var row = sheet.GetRow(rownum);
                for (int cellnum = 0; cellnum < row.Cells.Count; cellnum++)
                {
                    string cellValue = row.GetCell(cellnum).StringCellValue;

                    var fieldMatch = fieldRegex.Match(cellValue);
                    if (fieldMatch.Success)
                    {
                        PlaceholderDetail placeholder = new PlaceholderDetail();
                        placeholder.PlaceholderType = PlaceholderType.Field;
                        placeholder.RowIndex = rownum;
                        placeholder.ColIndex = cellnum;
                        placeholder.Placeholder = fieldMatch.Groups[1].Value;

                        placeholderList.Add(placeholder);
                    }
                    else
                    {
                        var listMatch = listRegex.Match(cellValue);
                        if (listMatch.Success)
                        {
                            PlaceholderDetail placeholder = new PlaceholderDetail();
                            placeholder.PlaceholderType = PlaceholderType.List;
                            placeholder.RowIndex = rownum;
                            placeholder.ColIndex = cellnum;
                            placeholder.Placeholder = listMatch.Groups[1].Value;

                            placeholderList.Add(placeholder);
                        }
                    }
                }
            }

            return placeholderList;
        }
    }

    /// <summary>
    /// 占位符明细
    /// </summary>
    public class PlaceholderDetail
    {
        /// <summary>
        /// 占位符类型：用于区分表头字段或者表体字段
        /// </summary>
        public PlaceholderType PlaceholderType { get; set; }

        /// <summary>
        /// 占位符所在的行
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 占位符所在的列
        /// </summary>
        public int ColIndex { get; set; }

        /// <summary>
        /// 占位符字段名称（和实体字段类字段要同名，否则反射获取不到值）
        /// </summary>
        public string Placeholder { get; set; }

        /// <summary>
        /// 表头、表体占位符
        /// </summary>
        public static Dictionary<PlaceholderType, string> FormatDic = new Dictionary<PlaceholderType, string>()
        {
            {PlaceholderType.Field,"#F#" },
            { PlaceholderType.List,"#L#"}
        };
    }

    public enum PlaceholderType
    {
        Field = 0,//单值字段
        List = 1//列表字段
    }

    /// <summary>
    /// 数据源
    /// </summary>
    /// <typeparam name="T1">表头实体类型</typeparam>
    /// <typeparam name="T2">表体实体类型</typeparam>
    public class DataSource<T1, T2>
        where T1 : class
        where T2 : class
    {
        public T1 FieldData { get; set; }
        public List<T2> ListData { get; set; }
    }

    public class SheetDataSource<T1, T2>
        where T1 : class
        where T2 : class
    {
        public string SheetName { get; set; }
        public DataSource<T1, T2> DataSource { get; set; }
    }
}
