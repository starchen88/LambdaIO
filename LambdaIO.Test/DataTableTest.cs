using System;
using Xunit;

namespace LambdaIO.Test
{
    public class DataTableTest
    {
        [Fact]
        public void TestOutput()
        {
            var mapper = new DefaultOutputMapper<Foo>();
            mapper.Add("Index", (it, i) => i);
            mapper.Add("IntProp", it => it.GetIntProp());
            mapper.Add("StringProp", it => it.StringProp);

            var foos = new[]{
                new Foo
                {
                    IntProp = 1,
                    StringProp = "’≈3"
                },new Foo
                {
                    IntProp = 2,
                    StringProp = "¿Ó4"
                },new Foo
                {
                    IntProp = 3,
                    StringProp = null
                }
            };
            var result = foos.ToDataTable(mapper);

            Assert.Equal(mapper.Count, result.Columns.Count);
            Assert.Equal(foos.Length, result.Rows.Count);

            for (int i = 0; i < foos.Length; i++)
            {
                var foo = foos[i];
                var row = result.Rows[i];
                Assert.Equal(row["Index"], i);
                Assert.Equal(row["IntProp"], foo.IntProp);
                Assert.Equal(row["StringProp"], (object)foo.StringProp ?? DBNull.Value);
            }
        }
    }
}
