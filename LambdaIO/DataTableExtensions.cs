using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace LambdaIO
{
    public static class DataTableExtensions
    {
        public static DataTable ToDataTable<TObject>(this IEnumerable<TObject> source, DefaultOutputMapper<TObject> outputMapper)
        {
            var dataTable = new DataTable();

            foreach (var item in outputMapper)
            {
                dataTable.Columns.Add(item.Key, item.Value.Item1);
            }

            var rowAction = CreateOutputAction<TObject>(outputMapper);

            foreach (var item in source)
            {
                var row = dataTable.NewRow();
                rowAction(item, row);
                dataTable.Rows.Add(row);
            }
            return dataTable;
        }
        public static Action<TObject, DataRow> CreateOutputAction<TObject>(DefaultOutputMapper<TObject> outputMapper)
        {
            var dataRowType = typeof(DataRow);
            var rowParameter = Expression.Parameter(dataRowType, "row");
            var valueType = typeof(object);

            //PropertyInfo indexer = (from p in dataRowType.GetDefaultMembers().OfType<PropertyInfo>()
            //                        where p.PropertyType == valueType
            //                        let q = p.GetIndexParameters()
            //                        where q.Length == 1 && q[0].ParameterType == typeof(string)
            //                        select p).Single();
            PropertyInfo indexer = (from p in dataRowType.GetDefaultMembers().OfType<PropertyInfo>()
                                    where p.PropertyType == valueType
                                    let q = p.GetIndexParameters()
                                    where q.Length == 1 && q[0].ParameterType == typeof(int)
                                    select p).Single();
            Console.WriteLine(indexer);
            var exps = new List<Expression>(outputMapper.Count);
            var dbNullExpression = Expression.Constant(DBNull.Value);
            int columnIndex = 0;
            foreach (var item in outputMapper)
            {
                
                var keyExpression = Expression.Constant(columnIndex);//  Expression.Constant(item.Key);
                var indexerExpression = Expression.Property(rowParameter, indexer, keyExpression);
                Expression valueExpression;
                valueExpression = item.Value.Item1 == valueType ? item.Value.Item2 : Expression.Convert(item.Value.Item2, valueType);
                //if (item.Value.Item1.IsValueType)
                //{
                //    valueExpression = item.Value.Item1 == valueType ? item.Value.Item2 : Expression.Convert(item.Value.Item2, valueType);
                //}
                //else
                //{
                //    //valueExpression = item.Value.Item1 == valueType ? item.Value.Item2 : Expression.Convert(Expression.Coalesce( item.Value.Item2, dbNullExpression,), valueType);
                //    //valueExpression = item.Value.Item1 == valueType ? item.Value.Item2 : Expression.Convert(Expression.Coalesce(item.Value.Item2, dbNullExpression,), valueType);
                //    if (item.Value.Item1 == valueType)
                //    {
                //        //valueExpression = Expression.Condition(Expression.Equal(item.Value.Item2, Expression.Constant(null)), item.Value.Item2, dbNullExpression);
                //        valueExpression = item.Value.Item1 == valueType ? item.Value.Item2 : Expression.Convert(item.Value.Item2, valueType);
                //    }
                //    else
                //    {
                //        //valueExpression = Expression.Condition(Expression.Equal(item.Value.Item2, Expression.Constant(null)), Expression.Convert(dbNullExpression, valueType), Expression.Convert(item.Value.Item2, valueType));
                //        valueExpression = item.Value.Item1 == valueType ? item.Value.Item2 : Expression.Convert(item.Value.Item2, valueType);
                //    }
                //}

                var assignExpression = Expression.Assign(indexerExpression, valueExpression);
                Console.WriteLine(assignExpression);
                exps.Add(assignExpression);
                columnIndex++;
            }

            var exps2one = Expression.Block(exps);
            var outputExpression = Expression.Lambda<Action<TObject, DataRow>>(exps2one, new[] { outputMapper.ParameterExpression, rowParameter });

            var outputAction = outputExpression.Compile();
            return outputAction;
        }
    }
}
