using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace LambdaIO.NPOI
{
    public static class XSSFSheetExtensions
    {
        public static int Fill<TObject>(this IEnumerable<TObject> source, DefaultOutputMapper<TObject> outputMapper, XSSFSheet sheet)
        {
            var rowAction = CreateOutputAction(outputMapper);

            int columnIndex = 0;

            IRow headerRow = sheet.CreateRow(0);
            foreach (var item in outputMapper)
            {
                headerRow.CreateCell(columnIndex).SetCellValue(item.Key);
                columnIndex++;
            }

            int rowIndex = 0;
            foreach (var item in source)
            {
                var row = (XSSFRow)sheet.CreateRow(rowIndex);
                rowAction(item, row, rowIndex);
                rowIndex++;
            }
            return rowIndex;
        }
        public static Action<TObject, XSSFRow, int> CreateOutputAction<TObject>(DefaultOutputMapper<TObject> outputMapper)
        {
            var dataRowType = typeof(XSSFRow);
            var rowParameter = Expression.Parameter(dataRowType, "row");

            var rowReplaceParameterExpressionVisitor = new ReplaceParameterExpressionVisitor(rowParameter);

            var exps = new List<Expression>(outputMapper.Count);

            int columnIndex = 0;
            foreach (var item in outputMapper)
            {
                var columnIndexExpression = Expression.Constant(columnIndex);

                Expression setCellValueExpression = null;
                CellType cellDataType;

                switch (Type.GetTypeCode(item.Value.Item1))
                {
                    case TypeCode.Byte:
                    case TypeCode.SByte:
                    case TypeCode.UInt16:
                    case TypeCode.UInt32:
                    case TypeCode.UInt64:
                    case TypeCode.Int16:
                    case TypeCode.Int32:
                    case TypeCode.Int64:
                    case TypeCode.Decimal:
                    case TypeCode.Single:
                        {
                            cellDataType = CellType.Numeric;

                            var valueExpression = Expression.Convert(item.Value.Item2, typeof(double));

                            Expression<Action<XSSFRow>> tempExp = row => row.CreateCell(-1, cellDataType).SetCellValue(default(double));
                            var replaceConstantExpressionVisitor = ReplaceConstantExpressionVisitor.Create(-1, columnIndexExpression, default(double), valueExpression);

                            setCellValueExpression = replaceConstantExpressionVisitor.Visit(tempExp.Body);

                            setCellValueExpression = rowReplaceParameterExpressionVisitor.Visit(setCellValueExpression);
                        }
                        break;
                    case TypeCode.Double:
                        {
                            cellDataType = CellType.Numeric;

                            var valueExpression = item.Value.Item2;

                            Expression<Action<XSSFRow>> tempExp = row => row.CreateCell(-1, cellDataType).SetCellValue(default(double));
                            var replaceConstantExpressionVisitor = ReplaceConstantExpressionVisitor.Create(-1, columnIndexExpression, default(double), valueExpression);

                            setCellValueExpression = replaceConstantExpressionVisitor.Visit(tempExp.Body);

                            setCellValueExpression = rowReplaceParameterExpressionVisitor.Visit(setCellValueExpression);
                        }
                        break;
                    case TypeCode.Empty:
                    case TypeCode.DBNull:
                        {
                            cellDataType = CellType.Blank;

                            Expression<Action<XSSFRow>> tempExp = row => row.CreateCell(-1, cellDataType);

                            setCellValueExpression = rowReplaceParameterExpressionVisitor.Visit(tempExp.Body);
                        }
                        break;
                    case TypeCode.Char:
                    case TypeCode.String:
                        {
                            cellDataType = CellType.String;

                            var valueExpression = Expression.Convert(item.Value.Item2, typeof(string));

                            Expression<Action<XSSFRow>> tempExp = row => row.CreateCell(-1, cellDataType).SetCellValue("");
                            var replaceConstantExpressionVisitor = ReplaceConstantExpressionVisitor.Create(-1, columnIndexExpression, "", valueExpression);

                            setCellValueExpression = replaceConstantExpressionVisitor.Visit(tempExp.Body);

                            setCellValueExpression = rowReplaceParameterExpressionVisitor.Visit(setCellValueExpression);
                        }
                        break;
                    case TypeCode.DateTime:
                        {
                            cellDataType = CellType.Numeric;

                            var valueExpression = Expression.Convert(item.Value.Item2, typeof(DateTime));

                            Expression<Action<XSSFRow>> tempExp = row => row.CreateCell(-1, cellDataType).SetCellValue(default(DateTime));
                            var replaceConstantExpressionVisitor = ReplaceConstantExpressionVisitor.Create(-1, columnIndexExpression, default(DateTime), valueExpression);

                            setCellValueExpression = replaceConstantExpressionVisitor.Visit(tempExp.Body);

                            setCellValueExpression = rowReplaceParameterExpressionVisitor.Visit(setCellValueExpression);
                        }
                        break;
                    case TypeCode.Boolean:
                        {
                            cellDataType = CellType.Boolean;

                            var valueExpression = item.Value.Item2;

                            Expression<Action<XSSFRow>> tempExp = row => row.CreateCell(-1, cellDataType).SetCellValue(default(bool).ToString());
                            var replaceConstantExpressionVisitor = ReplaceConstantExpressionVisitor.Create(-1, columnIndexExpression, default(bool), valueExpression);

                            setCellValueExpression = replaceConstantExpressionVisitor.Visit(tempExp.Body);

                            setCellValueExpression = rowReplaceParameterExpressionVisitor.Visit(setCellValueExpression);
                        }
                        break;
                    //case TypeCode.Object:
                    default:
                        {
                            cellDataType = CellType.String;

                            var valueExpression = item.Value.Item2;

                            Expression<Action<XSSFRow>> tempExp = row => row.CreateCell(-1, cellDataType).SetCellValue(new object().ToString());
                            var replaceConstantExpressionVisitor = ReplaceConstantExpressionVisitor.Create(-1, columnIndexExpression, new object(), valueExpression);

                            setCellValueExpression = replaceConstantExpressionVisitor.Visit(tempExp.Body);

                            setCellValueExpression = rowReplaceParameterExpressionVisitor.Visit(setCellValueExpression);
                        }
                        break;
                }
                if (setCellValueExpression != null)
                {
                    exps.Add(setCellValueExpression);
                }
                columnIndex++;
            }

            var exps2one = Expression.Block(exps);
            var resultExpression = Expression.Lambda<Action<TObject, XSSFRow, int>>(exps2one, new[] { outputMapper.ObjectParameterExpression, rowParameter, outputMapper.IndexParameterExpression });

            var resultAction = resultExpression.Compile();
            return resultAction;
        }
    }
}
