using System;
using System.Linq.Expressions;

namespace LambdaIO
{
    public class ReplaceParameterExpressionVisitor : ExpressionVisitor
    {
        private readonly Expression _parameter;
        public ReplaceParameterExpressionVisitor(Expression parameter)
        {
            _parameter = parameter;
        }
        protected override Expression VisitParameter(ParameterExpression node)
        {
            if (node.Type == _parameter.Type)
            {
                return _parameter;
            }
            else
            {
                return node;
            }
        }
    }

    public class ReplaceParameter2ExpressionVisitor : ExpressionVisitor
    {
        private readonly Expression _parameter1;
        private readonly Expression _parameter2;

        public ReplaceParameter2ExpressionVisitor(ParameterExpression parameter1, ParameterExpression parameter2)
        {
            if (parameter1.Type == parameter2.Type)
            {
                throw new NotSupportedException($"{nameof(parameter1)}与{nameof(parameter2)}的{nameof(parameter1.Type)}不能相同！");
            }
            _parameter1 = parameter1;
            _parameter2 = parameter2;
        }
        protected override Expression VisitParameter(ParameterExpression node)
        {
            if (node.Type == _parameter1.Type)
            {
                return _parameter1;
            }
            else if (node.Type == _parameter2.Type)
            {
                return _parameter2;
            }
            else
            {
                return node;
            }
        }
    }
}