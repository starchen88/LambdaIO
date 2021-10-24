using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Linq;
using System.Linq.Expressions;

namespace LambdaIO
{
    public class DefaultOutputMapper<TObject> : OutputMapper<TObject, string>
    {
        public OutputMapper<TObject, string> AddProperties(Func<PropertyInfo, bool> predicate = null, Func<PropertyInfo, string> getKey = null)
        {
            var objectType = typeof(TObject);
            var properties = objectType.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var property in properties)
            {
                if (property.CanRead && (predicate == null || predicate(property)))
                {
                    string key;
                    if (getKey != null)
                    {
                        key = getKey(property);
                        if (key == null)
                        {
                            throw new ArgumentException($"{nameof(getKey)}不能返回空");
                        }
                    }
                    else
                    {
                        key = property.Name;
                    }
                    var getValueExpression = Expression.Call(ObjectParameterExpression, property.GetGetMethod());
                    Add(key, property.PropertyType, getValueExpression);
                }
            }
            return this;
        }
    }
}
