using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace LambdaIO.NPOI
{
    public static class XSSFSheetExtensions
    {
        public static int Fill<TObject>(this IEnumerable<TObject> source, DefaultOutputMapper<TObject> outputMapper, XSSFSheet sheet)
        {
            return Fill(source, outputMapper, sheet, XSSFSheetOutputOption.Default);
        }
        public static int Fill<TObject>(this IEnumerable<TObject> source, DefaultOutputMapper<TObject> outputMapper, XSSFSheet sheet, XSSFSheetOutputOption option)
        {
            IDataFormat format = sheet.Workbook.CreateDataFormat();
            short dateFormat = format.GetFormat(option.DatetimeFormat);
            ICellStyle dateCellStyle = sheet.Workbook.CreateCellStyle();
            dateCellStyle.DataFormat = dateFormat;

            var rowAction = CreateOutputAction(outputMapper, dateCellStyle);

            int columnIndex = 0;

            int rowIndex = 0;
            IRow headerRow = sheet.CreateRow(rowIndex);
            foreach (var item in outputMapper)
            {
                headerRow.CreateCell(columnIndex,CellType.String).SetCellValue(item.Key);
                columnIndex++;
            }
            rowIndex++;

            foreach (var item in source)
            {
                var row = (XSSFRow)sheet.CreateRow(rowIndex);
                rowAction(item, row, rowIndex);
                rowIndex++;
            }
            if (option.AutoSizeColumn)
            {
                for (int i = 0; i < columnIndex; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
            }

            return rowIndex;
        }
        public static Action<TObject, XSSFRow, int> CreateOutputAction<TObject>(DefaultOutputMapper<TObject> outputMapper, ICellStyle dateCellStyle)
        {
            var dataRowType = typeof(XSSFRow);
            var rowParameter = Expression.Parameter(dataRowType, "row");

            var dateCellStyleExpression = Expression.Constant(dateCellStyle);

            var exps = new List<Expression>(outputMapper.Count);

            var numberTypes = new[]{
                typeof(byte),typeof(sbyte),
                typeof(ushort),typeof(short),
                typeof(uint),typeof(int),
                typeof(ulong),typeof(long),
                typeof(float),
                typeof(decimal),
            };

            int columnIndex = 0;
            foreach (var item in outputMapper)
            {
                Type valueType = item.Value.Item1;
                var valueExpression = item.Value.Item2;

                var columnIndexExpression = Expression.Constant(columnIndex);

                Expression cellExpression = null;

                Type cellValueType = null;
                Expression cellValueExpression = null;
                ICellStyle cellStyle = null;

                var nullableUnderlyingType = Nullable.GetUnderlyingType(valueType);
                if (nullableUnderlyingType != null)
                {
                    if (numberTypes.Contains(nullableUnderlyingType))
                    {
                        cellValueType = typeof(double?);
                        cellValueExpression = Expression.Convert(valueExpression, typeof(double?));
                    }
                    else if (valueType == typeof(DateTime))
                    {
                        cellValueType = typeof(DateTime?);
                        cellValueExpression = valueExpression;
                        cellStyle = dateCellStyle;
                    }
                    else if (valueType == typeof(DateTimeOffset))
                    {
                        cellValueType = typeof(DateTimeOffset?);
                        cellValueExpression = valueExpression;
                        cellStyle = dateCellStyle;
                    }
                    else if (valueType == typeof(bool))
                    {
                        cellValueType = typeof(bool?);
                        cellValueExpression = valueExpression;
                    }
                    else
                    {
                        cellValueType = typeof(object);
                        cellValueExpression = Expression.Convert(valueExpression, typeof(object));
                    }
                }
                else
                {
                    if (valueType == typeof(string))
                    {
                        cellValueType = typeof(string);
                        cellValueExpression = valueExpression;
                    }
                    else if (valueType == typeof(char))
                    {
                        cellValueType = typeof(string);
                        cellValueExpression = Expression.Convert(valueExpression, typeof(string));
                    }
                    else if (valueType == typeof(DateTime))
                    {
                        cellValueType = typeof(DateTime);
                        cellValueExpression = valueExpression;
                        cellStyle = dateCellStyle;
                    }
                    else if (valueType == typeof(DateTimeOffset))
                    {
                        cellValueType = typeof(DateTimeOffset);
                        cellValueExpression = valueExpression;
                        cellStyle = dateCellStyle;
                    }
                    else if (valueType == typeof(bool))
                    {
                        cellValueType = typeof(bool);
                        cellValueExpression = valueExpression;
                    }
                    else if (numberTypes.Contains(valueType))
                    {
                        cellValueType = typeof(double);
                        cellValueExpression = Expression.Convert(valueExpression, typeof(double));
                    }
                    else if (valueType == typeof(DBNull))
                    {
                        cellValueExpression = null;
                    }
                    else
                    {
                        cellValueType = typeof(object);
                        cellValueExpression = Expression.Convert(valueExpression, typeof(object));
                    }
                }

                if (cellValueExpression != null)
                {
                    if (cellStyle == null)
                    {
                        var method = (from m in typeof(XSSFSheetExtensions).GetMethods(BindingFlags.Static | BindingFlags.NonPublic)
                                      where m.Name == nameof(CreateCellAndSetValue)
                                      let q = m.GetParameters()
                                      where q.Length == 3 && q[2].ParameterType == cellValueType
                                      select m).Single();

                        cellExpression = Expression.Call(method, rowParameter, columnIndexExpression, cellValueExpression);
                    }
                    else
                    {
                        var method = (from m in typeof(XSSFSheetExtensions).GetMethods(BindingFlags.Static | BindingFlags.NonPublic)
                                      where m.Name == nameof(CreateCellAndSetValue)
                                      let q = m.GetParameters()
                                      where q.Length == 4 && q[2].ParameterType == cellValueType && q[3].ParameterType == typeof(ICellStyle)
                                      select m).Single();

                        cellExpression = Expression.Call(method, rowParameter, columnIndexExpression, cellValueExpression, dateCellStyleExpression);
                    }
                }
                else
                {
                    var method = (from m in typeof(XSSFSheetExtensions).GetMethods(BindingFlags.Static | BindingFlags.NonPublic)
                                  where m.Name == nameof(CreateBlankCell)
                                  let q = m.GetParameters()
                                  where q.Length == 2
                                  select m).Single();
                    cellExpression = Expression.Call(method, rowParameter, columnIndexExpression);
                }
                exps.Add(cellExpression);
                columnIndex++;
            }

            var exps2one = Expression.Block(exps);
            var resultExpression = Expression.Lambda<Action<TObject, XSSFRow, int>>(exps2one, new[] { outputMapper.ObjectParameterExpression, rowParameter, outputMapper.IndexParameterExpression });

            var resultAction = resultExpression.Compile();
            return resultAction;
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, string value)
        {
            var result = (XSSFCell)row.CreateCell(columnIndex, CellType.String);
            result.SetCellValue(value);
            return result;
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, double value)
        {
            var result = (XSSFCell)row.CreateCell(columnIndex, CellType.Numeric);
            result.SetCellValue(value);
            return result;
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, double? value)
        {
            var result = (XSSFCell)row.CreateCell(columnIndex, CellType.Numeric);
            if (value.HasValue)
            {
                result.SetCellValue(value.Value);
            }
            return result;
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, DateTime value, ICellStyle dateCellStyle)
        {
            var result = (XSSFCell)row.CreateCell(columnIndex, CellType.Numeric);
            result.SetCellValue(value);
            if (dateCellStyle != null)
            {
                result.CellStyle = dateCellStyle;
            }
            return result;
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, DateTime? value, ICellStyle dateCellStyle)
        {
            var result = (XSSFCell)row.CreateCell(columnIndex, CellType.Numeric);
            if (value.HasValue)
            {
                result.SetCellValue(value.Value);
            }
            if (dateCellStyle != null)
            {
                result.CellStyle = dateCellStyle;
            }
            return result;
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, DateTimeOffset value, ICellStyle dateCellStyle)
        {
            return CreateCellAndSetValue(row, columnIndex, value.DateTime, dateCellStyle);
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, DateTimeOffset? value, ICellStyle dateCellStyle)
        {
            return CreateCellAndSetValue(row, columnIndex, value?.DateTime, dateCellStyle);
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, bool value)
        {
            var result = (XSSFCell)row.CreateCell(columnIndex, CellType.Numeric);
            result.SetCellValue(value);
            return result;
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, bool? value)
        {
            var result = (XSSFCell)row.CreateCell(columnIndex, CellType.Numeric);
            if (value.HasValue)
            {
                result.SetCellValue(value.Value);
            }
            return result;
        }
        private static XSSFCell CreateCellAndSetValue(XSSFRow row, int columnIndex, object value)
        {
            var result = (XSSFCell)row.CreateCell(columnIndex, CellType.String);
            if (value != null)
            {
                result.SetCellValue(value.ToString());
            }
            return result;
        }
        private static XSSFCell CreateBlankCell(XSSFRow row, int columnIndex)
        {
            var result = (XSSFCell)row.CreateCell(columnIndex, CellType.Blank);
            return result;
        }
    }
}
