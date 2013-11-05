namespace ExcelProfiler
{
	using System;
	using System.Collections.Generic;
	using System.Diagnostics;
	using Microsoft.Office.Interop.Excel;

	public class Profiler
	{
		Workbook _workbook;
		Application _application;

		public Profiler(Workbook workbook)
		{
			_workbook = workbook;
			_application = (Application) workbook.Parent;
		}

		public double SheetTimer(Worksheet sheet)
		{
			MicroStopwatch timer = MicroStopwatch.StartNewMicroStopwatch();
			sheet.Calculate();
			return (double) timer.ElapsedMillisecondsHighResolution;
		}
		
		public double UsedRangeTimer(Worksheet sheet)
		{
            MicroStopwatch timer = MicroStopwatch.StartNewMicroStopwatch();
			if (Convert.ToDouble(_application.Version) >= 12.0)
			{
				sheet.UsedRange.CalculateRowMajorOrder();
			}
			else
			{
				sheet.UsedRange.Calculate();
			}

            return (double)timer.ElapsedMillisecondsHighResolution;
		}

		public double RecalcTimer()
		{
            MicroStopwatch timer = MicroStopwatch.StartNewMicroStopwatch();
			_application.Calculate();
            return (double)timer.ElapsedMillisecondsHighResolution;
		}

		public double FullCalcTimer()
		{
            MicroStopwatch timer = MicroStopwatch.StartNewMicroStopwatch();
			_application.CalculateFull();
            return (double)timer.ElapsedMillisecondsHighResolution;
		}

		public List<SheetTimerResult> GetSheetTimerResults()
		{
			var result = new List<SheetTimerResult>();
			foreach (Worksheet sheet in _workbook.Sheets)
			{
                sheet.Activate();
				result.Add(
					new SheetTimerResult() 
					{
						SheetName = sheet.Name,
						SheetCalcTime = this.SheetTimer(sheet),
						UsedRangeCalcTime = this.UsedRangeTimer(sheet)
					});
			}

			return result;
		}

		public WorkbookTimerResult GetWorkbookTimerResults()
		{
			WorkbookTimerResult result;

			result = new WorkbookTimerResult()
			{
				FullCalcTimerResult = FullCalcTimer(),
				RecalcTimerResult = RecalcTimer(),
				SheetTimerResults = GetSheetTimerResults()
			};

			return result;
		}

        public string WorkbookName
        {
            get
            {
                return _workbook.Name;
            }
        }
	}
}