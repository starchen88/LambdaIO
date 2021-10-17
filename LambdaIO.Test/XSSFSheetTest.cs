using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using LambdaIO;
using LambdaIO.NPOI;
using System.IO;

namespace LambdaIO.Test
{
    public class XSSFSheetTest
    {
        [Fact]
        public void TestOutput()
        {
            var mapper = new DefaultOutputMapper<Foo>();
            mapper.Add("Index", (it, i) => i);
            mapper.Add("Id", it => it.GetIntProp());
            mapper.Add("Name", it => it.StringProp);
            mapper.Add("Parent", it => it.FooProp);

            var foos = new[]{
                new Foo
                {
                    IntProp = 1,
                    StringProp = "张3",
                    FooProp=new Foo()
                },new Foo
                {
                    IntProp = 2,
                    StringProp = "李4",
                    FooProp=new Foo()
                },new Foo
                {
                    IntProp = 2,
                    StringProp = null,
                    FooProp=new Foo()
                }
            };

            var workbook = new XSSFWorkbook();
            var excelSheet = (XSSFSheet)workbook.CreateSheet("Sheet1");


            foos.Fill(mapper, excelSheet);
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "导出Excel", "导出.xlsx");
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            //Assert.Equal(mapper.Count, dt.Columns.Count);
            //Assert.Equal(foos.Length, dt.Rows.Count);

            //for (int i = 0; i < foos.Length; i++)
            //{
            //    var foo = foos[i];
            //    var row = dt.Rows[i];
            //    Assert.Equal(row["Index"], i);
            //    Assert.Equal(row["Id"], foo.Id);
            //    Assert.Equal(row["Name"], (object)foo.Name ?? DBNull.Value);
            //}
        }
    }
}
