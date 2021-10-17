using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace LambdaIO
{
    public static class DataTableExtensions
    {
        public static int Fill<TObject>(this IEnumerable<TObject> source, DefaultOutputMapper<TObject> outputMapper, DataTable dataTable)
        {
            if (dataTable.Columns.Count > 0 || dataTable.Rows.Count > 0)
            {
                throw new NotSupportedException($"{nameof(dataTable)}的Columns和Rows都必须为空！");
            }

            foreach (var item in outputMapper)
            {
                dataTable.Columns.Add(item.Key, item.Value.Item1);
            }

            var rowAction = CreateOutputAction<TObject>(outputMapper);

            int rowIndex = 0;
            foreach (var item in source)
            {
                var row = dataTable.NewRow();
                rowAction(item, row, rowIndex);
                dataTable.Rows.Add(row);
                rowIndex++;
            }
            return rowIndex;
        }
        public static DataTable ToDataTable<TObject>(this IEnumerable<TObject> source, DefaultOutputMapper<TObject> outputMapper)
        {
            var dataTable = new DataTable();
            Fill(source, outputMapper, dataTable);
            return dataTable;
        }
        public static Action<TObject, DataRow, int> CreateOutputAction<TObject>(DefaultOutputMapper<TObject> outputMapper)
        {
            var dataRowType = typeof(DataRow);
            var rowParameter = Expression.Parameter(dataRowType, "row");
            var valueType = typeof(object);

            PropertyInfo indexer = (from p in dataRowType.GetDefaultMembers().OfType<PropertyInfo>()
                                    where p.PropertyType == valueType
                                    let q = p.GetIndexParameters()
                                    where q.Length == 1 && q[0].ParameterType == typeof(int)
                                    select p).Single();
            Console.WriteLine(indexer);
            var exps = new List<Expression>(outputMapper.Count);



            int columnIndex = 0;
            foreach (var item in outputMapper)
            {
                var columnIndexExpression = Expression.Constant(columnIndex);

                var indexerExpression = Expression.Property(rowParameter, indexer, columnIndexExpression);
                Expression valueExpression;
                valueExpression = item.Value.Item1 == valueType ? item.Value.Item2 : Expression.Convert(item.Value.Item2, valueType);

                var assignExpression = Expression.Assign(indexerExpression, valueExpression);

                exps.Add(assignExpression);
                columnIndex++;
            }

            var exps2one = Expression.Block(exps);
            var resultExpression = Expression.Lambda<Action<TObject, DataRow, int>>(exps2one, new[] { outputMapper.ObjectParameterExpression, rowParameter, outputMapper.IndexParameterExpression });

            var resultAction = resultExpression.Compile();
            return resultAction;
        }
    }
}
