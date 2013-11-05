namespace ExcelProfiler
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Office.Interop.Excel;
    using ExcelDna.Integration;
    using ExcelDna.Integration.CustomUI;

    /// <summary>
    /// CustomUI Ribbon class that uses ribbon XML included in the .dna file
    /// </summary>
    [ComVisible(true)]
    public class CustomRibbon : ExcelRibbon
    {
        public void OnProfileActiveWorkbook(IRibbonControl control)
        {
            Application app = (Application)ExcelDnaUtil.Application;
            Profiler profiler = new Profiler(app.ActiveWorkbook);

            // save calc settings
            XlCalculation calcSave = app.Calculation;
            bool iterSave = app.Iteration;

            if (app.Calculation != XlCalculation.xlCalculationManual)
            {
                app.Calculation = XlCalculation.xlCalculationManual;
            }

            WorkbookTimerResult result = profiler.GetWorkbookTimerResults();
            Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet sheet = (Worksheet)wb.Sheets[1];

            sheet.Name = "Summary";

            sheet.Cells[1, 1].Value2 = "Calculation Time Summary (in ms)";

            sheet.Cells[2, 1].Value2 = profiler.WorkbookName;

            sheet.Cells[4, 2].Value2 = "FullCalc";
            sheet.Cells[4, 3].Value2 = "Recalc";
            sheet.Cells[4, 4].Value2 = "Volatility";

            sheet.Cells[5, 1].Value2 = "Workbook";
            sheet.Cells[5, 2].Value2 = result.FullCalcTimerResult;
            sheet.Cells[5, 3].Value2 = result.RecalcTimerResult;
            sheet.Cells[5, 4].Value2 = result.RecalcTimerResult / result.FullCalcTimerResult;

            sheet.Cells[7, 1].Value2 = "SheetName";
            sheet.Cells[7, 2].Value2 = "SheetCalc";
            sheet.Cells[7, 3].Value2 = "UsedRange";
            sheet.Cells[7, 4].Value2 = "Overhead";

            int row = 8;

            foreach (SheetTimerResult sheetResult in result.SheetTimerResults)
            {
                sheet.Cells[row, 1].Value2 = sheetResult.SheetName;
                sheet.Cells[row, 2].Value2 = sheetResult.SheetCalcTime;
                sheet.Cells[row, 3].Value2 = sheetResult.UsedRangeCalcTime;
                sheet.Cells[row, 4].Value2 = Math.Max(sheetResult.SheetCalcTime - sheetResult.UsedRangeCalcTime, 0.0);
                    
                row++;
            }

            //format results

            ((Range)sheet.Columns[1]).EntireColumn.AutoFit();
            ((Range)sheet.Columns[2]).EntireColumn.AutoFit();
            ((Range)sheet.Columns[3]).EntireColumn.AutoFit();
            ((Range)sheet.Columns[4]).EntireColumn.AutoFit();

            ListObject table = sheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, sheet.Range["$A$7:$D$" + (row - 1).ToString()], null, XlYesNoGuess.xlYes);
            table.Name = "SheetCalcResults";
            table.ShowTableStyleRowStripes = false;

            //sort sheet calc results
            table.Sort.SortFields.Clear();
            table.Sort.SortFields.Add(sheet.Range["B8"], XlSortOn.xlSortOnValues, XlSortOrder.xlDescending, XlSortDataOption.xlSortNormal);
            table.Sort.Header = XlYesNoGuess.xlYes;
            table.Sort.MatchCase = false;
            table.Sort.SortMethod = XlSortMethod.xlPinYin;
            table.Sort.Apply();

            sheet.Range["A1"].Select();

            //restore calc settings
            if (app.Calculation != calcSave) { app.Calculation = calcSave; }
            if (app.Iteration != iterSave) { app.Iteration = iterSave; }
        }
    }
}