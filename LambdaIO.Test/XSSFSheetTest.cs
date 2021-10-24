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
            mapper.Add("IntProp", it => it.GetIntProp());
            mapper.Add("StringProp", it => it.StringProp);
            mapper.Add("FooProp", it => it.FooProp);
            mapper.Add("NullableIntProp", it => it.NullableIntProp);
            mapper.Add("FooEnumProp", it => it.FooEnumProp);
            mapper.Add("DateTimeOffsetProp", it => it.DateTimeOffsetProp);

            var foos = new[]{
                new Foo
                {
                    IntProp = 1,
                    StringProp = "张3",
                    FooProp=new Foo(),
                    NullableIntProp=1,
                    FooEnumProp=null,
                    DateTimeOffsetProp=new DateTimeOffset()
                },new Foo
                {
                    IntProp = 2,
                    StringProp = "李4",
                    FooProp=new Foo(),
                    NullableIntProp=null,
                    FooEnumProp=FooEnum.A,
                    DateTimeOffsetProp=DateTimeOffset.Now
                },new Foo
                {
                    IntProp = 3,
                    StringProp = null,
                    FooProp=new Foo(),
                    NullableIntProp=3,
                    FooEnumProp=FooEnum.B,
                    DateTimeOffsetProp=DateTimeOffset.MaxValue
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
        }
        [Fact]
        public void TestOutputAllProperties()
        {
            var mapper = new DefaultOutputMapper<Foo>();
            mapper.AddProperties();

            var foos = new[]{
                new Foo
                {
                    IntProp = 1,
                    StringProp = "张3",
                    FooProp=new Foo(),
                    NullableIntProp=1,
                    FooEnumProp=null,
                    DateTimeOffsetProp=new DateTimeOffset()
                },new Foo
                {
                    IntProp = 2,
                    StringProp = "李4",
                    FooProp=new Foo(),
                    NullableIntProp=null,
                    FooEnumProp=FooEnum.A,
                    DateTimeOffsetProp=DateTimeOffset.Now
                },new Foo
                {
                    IntProp = 3,
                    StringProp = null,
                    FooProp=new Foo(),
                    NullableIntProp=3,
                    FooEnumProp=FooEnum.B,
                    DateTimeOffsetProp=DateTimeOffset.MaxValue
                }
            };

            var workbook = new XSSFWorkbook();
            var excelSheet = (XSSFSheet)workbook.CreateSheet("Sheet1");

            foos.Fill(mapper, excelSheet);
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "导出Excel", "导出AllProperties.xlsx");
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }
    }
}
