using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using CR.Core.Extention;

namespace CR.Core.Excel
{
    public class ExcelEngine
    {

        private Workbook _workbook;

        public ExcelEngine(string path)
        {
            if (!File.Exists(path))
                throw new Exception("文件不存在");

            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            _workbook = new Workbook(fs);
        }

        public ExcelEngine(Stream stream)
        {
            _workbook = new Workbook(stream);
        }

        public bool IsPropertyAndColumnFit(Type t, IEnumerable<string> cellname)
        {
            //为了防止传入的类一条ExcelColumnAttribute也没有配置
            int attrcount = 0;
            foreach (var prop in t.GetProperties())
            {
                ExcelColumnAttribute attr = prop.GetCustomAttribute(typeof(ExcelColumnAttribute)) as ExcelColumnAttribute;
                if (attr == null)
                    continue;

                attrcount++;
                if (!cellname.Any(o => o == attr.Field))
                    return false;
            }
            return attrcount == 0 ? false : true;
        }

        #region 读取表格

        public List<T> ReadData<T>() where T : new()
        {
            //读取第一页的SHEET，并划分出需要处理的区域来
            Worksheet sheet = _workbook.Worksheets[0];
            int maxcolumn = 0;
            while (sheet.Cells[0, maxcolumn].Type != CellValueType.IsNull) { maxcolumn++; };

            Range range = sheet.Cells.CreateRange(1, 0, sheet.Cells.Rows.Count - 1, maxcolumn);

            //读取实体的变量与列名的对应关系
            Dictionary<string, PropertyInfo> dicentity = GetEntityPropertyMapping<T>();
            if (dicentity == null || dicentity.Count == 0)
                throw new Exception("实体没有配置与EXCEL表头对应的列");

            //读取EXCEL文件中的标题与标题所在的列的INDEX对应的关系`
            Dictionary<string, int> dicexcel = GetTableHeaderMapping();
            if (dicexcel == null || dicexcel.Count == 0)
                throw new Exception("导入的EXCEL文件中未发现表头行");

            //1、使用实体列名与EXCEL列名关系起来，找到实体对应的属性与表格中的列对应
            //2、抽取数据
            List<T> retdata = new List<T>();
            for (int i = 0; i < range.RowCount; i++)
            {
                T t = new T();
                bool isemptyrow = true;
                foreach (var key in dicentity.Keys)
                {
                    if (dicexcel.ContainsKey(key))
                    {
                        dynamic val = getPropertyValue(dicentity[key], range[i, dicexcel[key]]);
                        //如果值不为空，则直接将值赋给属性
                        if (val != null)
                            isemptyrow = false;
                        else
                        {
                            //如果根据属性直接取的值为空，则从属性中取出特性中的默认值
                            //如果默认值中没有定义默认值类型，则直接将默认值赋给属性，否则将先进行转换后再赋值
                            //这主要是因为特性中只允许赋值常量的数据,decimal和datetime类型的属性不能指定为默认值，所以在这里需要做一个类型转换
                            ExcelColumnAttribute attr = dicentity[key].GetCustomAttribute(typeof(ExcelColumnAttribute)) as ExcelColumnAttribute;
                            if (attr.IsRequire)
                            {
                                throw new Exception(attr.Field + "不能为空");
                            }

                            if (attr != null && attr.Default != null && attr.DefaultType != null)
                                val = Convert.ChangeType(attr.Default, attr.DefaultType);
                            else
                                val = attr.Default;
                        }
                        dicentity[key].SetValue(t, val);
                    }
                }

                if (!isemptyrow)
                    retdata.Add(t);
            }
            return retdata;
            //return default(T);
        }

        public ExtracterComplexModel ReadComplexData(IReadHandler handler)
        {
            return handler.AnalyeExcelData(_workbook);
        }

        #endregion

        #region 导出表格

        public Workbook ExportData(IWriteHandler handler, Dictionary<string, object> param)
        {
            var wb = handler.ExecuteWork(_workbook, param);
            return wb;
        }

        #endregion

        public byte[] ConvertToByte<T>(IEnumerable<T> data) where T : new()
        {
            //读取第一页的SHEET，并划分出需要处理的区域来
            Worksheet sheet = _workbook.Worksheets[0];
            int maxcolumn = 0;
            while (sheet.Cells[0, maxcolumn].Type != CellValueType.IsNull) { maxcolumn++; };

            //读取实体的变量与列名的对应关系
            Dictionary<string, PropertyInfo> dicentity = GetEntityPropertyMapping<T>();
            if (dicentity == null || dicentity.Count == 0)
                throw new Exception("实体没有配置与EXCEL表头对应的列");

            //读取EXCEL文件中的标题与标题所在的列的INDEX对应的关系`
            Dictionary<string, int> dicexcel = GetTableHeaderMapping();
            if (dicexcel == null || dicexcel.Count == 0)
                throw new Exception("EXCEL模板中未发现对应列");

            sheet.Cells.InsertRows(2, data.Count() - 1);
            Range range = sheet.Cells.CreateRange(1, 0, data.Count(), maxcolumn);
            for (int i = 0; i < data.Count(); i++)
            {
                var rowdata = data.ElementAt(i);
                foreach (var key in dicentity.Keys)
                {
                    if (dicexcel.ContainsKey(key))
                    {
                        range[i, dicexcel[key]].Value = dicentity[key].GetValue(rowdata);
                    }
                }
            }
            return _workbook.SaveToStream().ToArray();
        }

        /// <summary>
        /// 取出属性在单元格中的值
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        private object getPropertyValue(PropertyInfo prop, Cell cell)
        {
            if (cell.Value == null)
                return null;

            Func<PropertyInfo, object> func = (p) => {
                switch (p.PropertyType.Name.ToLower())
                {
                    case "string":
                        return cell.StringValue;
                    case "int32":
                    case "int64":
                        return cell.IntValue;
                    case "decimal":
                        return (cell.Value ?? "").ToString().ToDecimal();
                    case "double":
                        return cell.DoubleValue;
                    case "float":
                        return cell.FloatValue;
                    case "datetime":
                        return cell.DateTimeValue;
                    default:
                        return cell.Value;
                }
            };

            if (prop.PropertyType.Name.ToLower() == "nullable`1")
            {
                return func(prop.PropertyType.GetProperties()[1]);
            }
            else
            {
                return func(prop);
            }
        }

        /// <summary>
        /// 获取类型的列名与属性信息的对应关系
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private Dictionary<string, PropertyInfo> GetEntityPropertyMapping<T>() where T : new()
        {
            Dictionary<string, PropertyInfo> dic = new Dictionary<string, PropertyInfo>();
            foreach (PropertyInfo prop in typeof(T).GetProperties())
            {
                Attribute attr = prop.GetCustomAttribute(typeof(ExcelColumnAttribute));
                if (attr == null)
                    continue;

                ExcelColumnAttribute colattr = attr as ExcelColumnAttribute;
                dic.Add(colattr.Field, prop);
            }

            return dic;
        }

        /// <summary>
        /// 取出表格中列名与索引对应关系
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, int> GetTableHeaderMapping()
        {
            Worksheet sheet = _workbook.Worksheets[0];
            Range r = sheet.Cells.CreateRange(0, 0, 1, 1000);

            Dictionary<string, int> dicmap = new Dictionary<string, int>();
            for (int i = 0; i < r.ColumnCount; i++)
            {
                Cell c = r.GetCellOrNull(0, i);
                if (c == null || c.Value == null)
                    break;
                dicmap.Add(c.StringValue, i);
            }
            return dicmap;
        }
    }
}
