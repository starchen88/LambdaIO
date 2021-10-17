using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;

namespace LambdaIO
{
    public class ReplaceConstantExpressionVisitor
    {
        public static ReplaceConstantExpressionVisitor<TValue> Create<TValue>(TValue oldValue, Expression newExpression)
        {
            return new ReplaceConstantExpressionVisitor<TValue>(oldValue, newExpression);
        }
        public static ReplaceConstantExpressionVisitor<TValue1, TValue2> Create<TValue1, TValue2>(TValue1 oldValue1, Expression newExpression1, TValue2 oldValue2, Expression newExpression2)
        {
            return new ReplaceConstantExpressionVisitor<TValue1, TValue2>(oldValue1, newExpression1, oldValue2, newExpression2);
        }
    }
    public class ReplaceConstantExpressionVisitor<TValue> : ExpressionVisitor
    {
        private readonly TValue _oldValue;
        private readonly Expression _newExpression;
        public ReplaceConstantExpressionVisitor(TValue oldValue, Expression newExpression)
        {
            _oldValue = oldValue;
            _newExpression = newExpression;
        }
        protected override Expression VisitConstant(ConstantExpression node)
        {
            if (node is ConstantExpression constantExpression && constantExpression.Value is TValue value && value.Equals(_oldValue))
            {
                return _newExpression;
            }
            else
            {
                return node;
            }
        }
    }
    public class ReplaceConstantExpressionVisitor<TValue1, TValue2> : ExpressionVisitor
    {
        private readonly TValue1 _oldValue1;
        private readonly Expression _newExpression1;
        private readonly TValue2 _oldValue2;
        private readonly Expression _newExpression2;
        public ReplaceConstantExpressionVisitor(TValue1 oldValue1, Expression newExpression1, TValue2 oldValue2, Expression newExpression2)
        {
            _oldValue1 = oldValue1;
            _newExpression1 = newExpression1;

            _oldValue2 = oldValue2;
            _newExpression2 = newExpression2;
        }
        protected override Expression VisitConstant(ConstantExpression node)
        {
            if (node is ConstantExpression constantExpression1 && constantExpression1.Value is TValue1 value1 && value1.Equals(_oldValue1))
            {
                return _newExpression1;
            }
            else if (node is ConstantExpression constantExpression2 && constantExpression2.Value is TValue2 value2 && value2.Equals(_oldValue2))
            {
                return _newExpression2;
            }
            else
            {
                return node;
            }
        }
    }
}
