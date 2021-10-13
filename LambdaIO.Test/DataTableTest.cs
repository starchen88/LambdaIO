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
            mapper.Add("Id", it => it.GetId());
            mapper.Add("Name", it => it.Name);

            var foos = new[]{
                new Foo
                {
                    Id = 1,
                    Name = "’≈3"
                },new Foo
                {
                    Id = 2,
                    Name = "¿Ó4"
                },new Foo
                {
                    Id = 2,
                    Name = null
                }
            };
            var dt = foos.ToDataTable(mapper);

            Assert.Equal(mapper.Count, dt.Columns.Count);
            Assert.Equal(foos.Length, dt.Rows.Count);

            for (int i = 0; i < foos.Length; i++)
            {
                var foo = foos[i];
                var row = dt.Rows[i];

                Assert.Equal(row["Id"], foo.Id);
                Assert.Equal(row["Name"], (object)foo.Name ?? DBNull.Value);
            }
        }
    }
}
