using System;
using System.Linq.Expressions;

namespace LambdaIO
{
    public class ReplaceParameterVisitor : ExpressionVisitor
    {
        private readonly Type _type;
        private readonly Expression _parameter;
        public ReplaceParameterVisitor(Expression parameter)
        {
            _type = parameter.Type;
            _parameter = parameter;
        }
        protected override Expression VisitParameter(ParameterExpression node)
        {
            if (node.Type == _type)
            {
                return _parameter;
            }
            else
            {
                return node;
            }
        }
    }
}
