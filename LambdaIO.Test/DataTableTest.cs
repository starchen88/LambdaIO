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
            mapper.Add("Id", it => it.GetIntProp());
            mapper.Add("Name", it => it.StringProp);

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
                    IntProp = 2,
                    StringProp = null
                }
            };
            var dt = foos.ToDataTable(mapper);

            Assert.Equal(mapper.Count, dt.Columns.Count);
            Assert.Equal(foos.Length, dt.Rows.Count);

            for (int i = 0; i < foos.Length; i++)
            {
                var foo = foos[i];
                var row = dt.Rows[i];
                Assert.Equal(row["Index"], i);
                Assert.Equal(row["Id"], foo.IntProp);
                Assert.Equal(row["Name"], (object)foo.StringProp ?? DBNull.Value);
            }
        }
    }
}
