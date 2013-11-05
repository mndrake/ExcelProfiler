using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelProfiler
{
    public class SheetTimerResult
    {
        public string SheetName { get; set; }
        public double SheetCalcTime { get; set; }
        public double UsedRangeCalcTime { get; set; }
    }
}
