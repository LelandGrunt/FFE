using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Serilog;
using System;
using System.Collections.Generic;
using static ExcelDna.Integration.XlCall;

namespace FFE
{
    public static class FfeExcel
    {
        private static readonly Application excelApp;

        private static readonly ILogger log;

        static FfeExcel()
        {
            excelApp = (Application)ExcelDnaUtil.Application;

            log = Log.ForContext("UDF", "FFE");
        }

        public static bool ScreenUpdating
        {
            get => excelApp.ScreenUpdating;
            set => excelApp.ScreenUpdating = value;
        }

        public static void AddWorkbook()
        {
            excelApp.Workbooks.Add();
        }

        public static dynamic ReferenceToRange(ExcelReference xlref)
        {
            return excelApp.Range[Excel(xlfReftext, xlref, true)];
        }

        //https://www.add-in-express.com/creating-addins-blog/2011/03/23/excel-check-user-edit-cell/
        public static bool IsInEditingMode()
        {
            if (excelApp.Interactive) { return false; }

            try
            {
                excelApp.Interactive = false;
                excelApp.Interactive = true;
            }
            catch (Exception)
            {
                return true;
            }
            return false;
        }

        public static List<Range> FindCellsByValue(string value, Worksheet worksheet = null)
        {
            return FindCells(value, new ExcelFindOptions() { LookIn = XlFindLookIn.xlValues }, worksheet);
        }

        public static List<Range> FindCellsByFormula(string formula, Worksheet worksheet = null)
        {
            return FindCells(formula, new ExcelFindOptions() { LookIn = XlFindLookIn.xlFormulas }, worksheet);
        }

        public static List<Range> FindCells(string what, ExcelFindOptions excelFindOptions = null, Worksheet worksheet = null)
        {
            List<Range> foundedRanges = new List<Range>();

            if (excelFindOptions == null)
                excelFindOptions = new ExcelFindOptions();

            List<Worksheet> worksheets = new List<Worksheet>();

            if (worksheet != null)
            {
                worksheets.Add(worksheet);
            }
            else
            {
                foreach (Worksheet sheet in excelApp.Worksheets)
                {
                    worksheets.Add(sheet);
                }
            }

            Range rangeFirstFind = null;
            Range rangeCurrentFind = null;

            foreach (Worksheet sheet in worksheets)
            {
                rangeCurrentFind = sheet.Cells.Find(what,
                                                    LookIn: excelFindOptions.LookIn,
                                                    LookAt: excelFindOptions.LookAt,
                                                    SearchOrder: excelFindOptions.SearchOrder,
                                                    SearchDirection: excelFindOptions.SearchDirection,
                                                    MatchCase: excelFindOptions.MatchCase,
                                                    SearchFormat: excelFindOptions.SearchFormat);

                while (rangeCurrentFind != null)
                {
                    if (rangeFirstFind == null)
                    {
                        rangeFirstFind = rangeCurrentFind;
                    }

                    else if (rangeCurrentFind.get_Address(XlReferenceStyle.xlA1)
                             == rangeFirstFind.get_Address(XlReferenceStyle.xlA1))
                    {
                        break;
                    }

                    foundedRanges.Add(rangeCurrentFind);

                    rangeCurrentFind = sheet.Cells.FindNext(rangeCurrentFind);
                }

                rangeFirstFind = null;
                rangeCurrentFind = null;
            }

            return foundedRanges;
        }

        public static List<Range> FindCellsInRangeByValue(string value, Range range)
        {
            return FindCellsInRange(value, range, new ExcelFindOptions() { LookIn = XlFindLookIn.xlValues });
        }

        public static List<Range> FindCellsInRangeByFormula(string formula, Range range)
        {
            return FindCellsInRange(formula, range, new ExcelFindOptions() { LookIn = XlFindLookIn.xlFormulas });
        }

        public static List<Range> FindCellsInRange(string what, Range range, ExcelFindOptions excelFindOptions = null)
        {
            List<Range> foundedRanges = new List<Range>();

            if (range.Cells.Count == 1)
            {
                string formula = range.Formula;
                if (formula.Contains(what))
                {
                    foundedRanges.Add(range);
                }

                return foundedRanges;
            }

            if (excelFindOptions == null)
                excelFindOptions = new ExcelFindOptions();

            Range rangeFirstFind = null;
            Range rangeCurrentFind = null;

            rangeCurrentFind = range.Find(what,
                                          LookIn: excelFindOptions.LookIn,
                                          LookAt: excelFindOptions.LookAt,
                                          SearchOrder: excelFindOptions.SearchOrder,
                                          SearchDirection: excelFindOptions.SearchDirection,
                                          MatchCase: excelFindOptions.MatchCase,
                                          SearchFormat: excelFindOptions.SearchFormat);

            while (rangeCurrentFind != null)
            {
                if (rangeFirstFind == null)
                {
                    rangeFirstFind = rangeCurrentFind;
                }

                else if (rangeCurrentFind.get_Address(XlReferenceStyle.xlA1)
                         == rangeFirstFind.get_Address(XlReferenceStyle.xlA1))
                {
                    break;
                }

                foundedRanges.Add(rangeCurrentFind);

                rangeCurrentFind = range.FindNext(rangeCurrentFind);
            }

            rangeFirstFind = null;
            rangeCurrentFind = null;

            return foundedRanges;
        }

        public static void Refresh(List<Range> ranges, string functionName = null)
        {
            Recalculate(ranges, functionName);
        }

        public static void Refresh(this IEnumerable<Range> ranges, string functionName = null)
        {
            Recalculate(ranges, functionName);
        }

        private static void Recalculate(IEnumerable<Range> ranges, string functionName = null)
        {
            foreach (Range range in ranges)
            {
                if ((functionName != null && functionName.StartsWith("QAV"))
                    && AvqSetting.Default.StopRefreshAtFirstCallLimitReachedError
                    && Avq.CallLimitReachedError)
                {
                    log.Debug("Recalculation of AVQ functions were stopped due to Stop Refresh At First #AV_CALL_LIMIT_REACHED error setting.");

                    // Stop recalulation for current AVQ function.
                    continue;
                }

                string rangeAdress = range.Address[false, false, XlReferenceStyle.xlA1, true].Split(new[] { "]" }, StringSplitOptions.None)[1];

                log.Debug("Recalculate range {@Range} with function name {@FunctionName}", rangeAdress, functionName);

                excelApp.StatusBar = $"Recalculate {functionName} in {rangeAdress}...";

                if (excelApp.Calculation == XlCalculation.xlCalculationAutomatic)
                {
                    /* The replace method calculates all ranges with the same formula on the sheet, not only the selection. Why?
                    string formula = functionName + "(";
                    //range.Replace(formula, formula, XlLookAt.xlPart, XlSearchOrder.xlByColumns, false, false, false, false);
                    range.Replace(formula, formula);*/
                    // Or
                    //range.FormulaR1C1 = range.FormulaR1C1;
                    // Or
                    range.Dirty();
                }
                else
                {
                    // If range is not on active worksheet, then Select() method fail.
                    range.Worksheet.Select();
                    range.Select();
                    range.Calculate();
                    // Or (In manual calculation mode, the range is calculated after switch to automatic mode).
                    // IDEA: New AVQ option?
                    //range.Dirty();
                }
            }

            excelApp.StatusBar = null;
        }
    }

    public class ExcelFindOptions
    {
        public object LookIn { get; set; } = XlFindLookIn.xlFormulas;
        public object LookAt { get; set; } = XlLookAt.xlPart;
        public object SearchOrder { get; set; } = XlSearchOrder.xlByColumns;
        public XlSearchDirection SearchDirection { get; set; } = XlSearchDirection.xlNext;
        public object MatchCase { get; set; } = false;
        public object SearchFormat { get; set; } = false;
    }
}