using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace LambdaIO
{
    public class OutputMapper<TObject, TKey> : IEnumerable<KeyValuePair<TKey, Tuple<Type, Expression>>>
    {
        protected readonly Dictionary<TKey, Tuple<Type, Expression>> _keyAndGetValueExpressions;
        protected readonly ParameterExpression _objectParameterExpression;
        protected readonly ParameterExpression _indexParameterExpression;

        protected readonly ReplaceParameter2ExpressionVisitor _replaceParameterExpressionVisitor;

        public OutputMapper()
        {
            _keyAndGetValueExpressions = new Dictionary<TKey, Tuple<Type, Expression>>();
            _objectParameterExpression = Expression.Parameter(typeof(TObject), typeof(TObject).Name.ToLower());
            _indexParameterExpression = Expression.Parameter(typeof(int), "index");
            _replaceParameterExpressionVisitor = new ReplaceParameter2ExpressionVisitor(_objectParameterExpression, _indexParameterExpression);

        }
        public ParameterExpression ObjectParameterExpression => _objectParameterExpression;
        public int Count => _keyAndGetValueExpressions.Count;
        public ParameterExpression IndexParameterExpression => _indexParameterExpression;
        public Expression ReplaceParameter(Expression expression)
        {
            return _replaceParameterExpressionVisitor.Visit(expression);
        }
        private OutputMapper<TObject, TKey> Add(TKey key, Type valueType, Expression getValueExpression)
        {
            var result = _replaceParameterExpressionVisitor.Visit(getValueExpression);
            _keyAndGetValueExpressions.Add(key, Tuple.Create(valueType, result));
            return this;
        }
        public OutputMapper<TObject, TKey> Add<TValue>(TKey key, Expression<Func<TObject, TValue>> getValueExpression)
        {
            return Add(key, typeof(TValue), getValueExpression.Body);
        }
        public OutputMapper<TObject, TKey> Add<TValue>(TKey key, Expression<Func<TObject, int, TValue>> getValueExpression)
        {
            return Add(key, typeof(TValue), getValueExpression.Body);
        }
        public OutputMapper<TObject, TKey> AddFunc<TValue>(TKey key, Func<TObject, TValue> getValue)
        {
            var getValueExpression = Expression.Invoke(Expression.Constant(getValue), _objectParameterExpression);

            return Add(key, typeof(TValue), getValueExpression);
        }
        public OutputMapper<TObject, TKey> AddFunc<TValue>(TKey key, Func<TObject, int, TValue> getValue)
        {
            var getValueExpression = Expression.Invoke(Expression.Constant(getValue), _objectParameterExpression);

            return Add(key, typeof(TValue), getValueExpression);
        }

        IEnumerator<KeyValuePair<TKey, Tuple<Type, Expression>>> IEnumerable<KeyValuePair<TKey, Tuple<Type, Expression>>>.GetEnumerator()
        {
            return _keyAndGetValueExpressions.GetEnumerator();
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _keyAndGetValueExpressions.GetEnumerator();
        }

    }
}
