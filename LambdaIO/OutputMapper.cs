using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;

namespace LambdaIO
{
    public class OutputMapper<TObject, TKey> : IEnumerable<KeyValuePair<TKey, Tuple<Type, Expression>>>
    {
        protected readonly Dictionary<TKey, Tuple<Type, Expression>> _keyAndGetValueExpressions;
        protected readonly ParameterExpression _parameterExpression;
        protected readonly ReplaceParameterVisitor _replaceParameterVisitor;
        protected readonly Type _objectType;
        public OutputMapper()
        {
            _objectType = typeof(TObject);
            _keyAndGetValueExpressions = new Dictionary<TKey, Tuple<Type, Expression>>();
            _parameterExpression = Expression.Parameter(_objectType, _objectType.Name.ToLower());
            _replaceParameterVisitor = new ReplaceParameterVisitor(_parameterExpression);
        }
        private OutputMapper<TObject, TKey> Add(TKey key, Type valueType, Expression getValueExpression)
        {
            var result = _replaceParameterVisitor.Visit(getValueExpression);
            _keyAndGetValueExpressions.Add(key, Tuple.Create(valueType, result));
            return this;
        }
        public OutputMapper<TObject, TKey> Add<TValue>(TKey key, Expression<Func<TObject, TValue>> getValueExpression)
        {
            if (getValueExpression.Body is MemberExpression memberExpression && getValueExpression.Parameters[0] == memberExpression.Expression)
            {
                return Add(key, typeof(TValue), getValueExpression.Body);
            }
            else if (getValueExpression.Body is MethodCallExpression methodCallExpression && getValueExpression.Parameters[0] == methodCallExpression.Object)
            {
                return Add(key, typeof(TValue), getValueExpression.Body);
            }
            else
            {
                return Add(key, typeof(TValue), getValueExpression);
            }
        }
        public OutputMapper<TObject, TKey> AddFunc<TValue>(TKey key, Func<TObject, TValue> getValue)
        {
            var getValueExpression = Expression.Invoke(Expression.Constant(getValue), _parameterExpression);

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
        public ParameterExpression ParameterExpression
        {
            get => _parameterExpression;
        }
        public int Count
        {
            get => _keyAndGetValueExpressions.Count;
        }
    }
}
