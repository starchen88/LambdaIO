using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LambdaIO.Test
{
    public class Foo
    {
        public int IntProp { get; set; }
        public string StringProp { get; set; }
        public int GetIntProp() => IntProp;
        public Foo FooProp { get; set; }
        public int? NullableIntProp { get; set; }
        public FooEnum? FooEnumProp { get; set; }
        public DateTimeOffset DateTimeOffsetProp { get; set; }
    }
    public enum FooEnum
    {
        A,
        B,
        C
    }
}
