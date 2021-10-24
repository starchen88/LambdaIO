using System;
using System.Collections.Generic;
using System.Text;

namespace LambdaIO.NPOI
{
    public class XSSFSheetOutputOption
    {
        public string DatetimeFormat { get; set; }
        public bool AutoSizeColumn { get; set; }
        public static XSSFSheetOutputOption Default =>
            new XSSFSheetOutputOption
            {
                DatetimeFormat = "yyyy-MM-dd",
                AutoSizeColumn = true
            };
    }
}
