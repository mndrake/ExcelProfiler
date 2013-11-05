using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelProfiler
{
    public class WorkbookTimerResult
    {
        public List<SheetTimerResult> SheetTimerResults { get; set; }
        public double RecalcTimerResult { get; set; }
        public double FullCalcTimerResult { get; set; }
    }
}
