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

        static FfeExcel()
        {
            Log.Logger = FfeLogger.ConfigureLogging(CbqSetting.Default.LogLevel)
                                  .ForContext("UDF", typeof(FfeExcel));

            excelApp = (Application)ExcelDnaUtil.Application;
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

        public static void Recalculate(List<Range> ranges, string functionName = null)
        {
            foreach (Range range in ranges)
            {
                string rangeAdress = range.Address[false, false, XlReferenceStyle.xlA1, true].Split(new[] { "]" }, StringSplitOptions.None)[1];

                Log.Debug("Recalculate range {@Range} with functionName {@FunctionName}", rangeAdress, functionName);

                excelApp.StatusBar = $"Recalculate {functionName} in {rangeAdress}...";

                //range.CalculateRowMajorOrder();
                string formula = functionName + "(";
                range.Replace(formula, formula);
            }

            excelApp.StatusBar = null;
        }

        public static void Recalculate(this IEnumerable<Range> ranges, string functionName = null)
        {
            foreach (Range range in ranges)
            {
                string rangeAdress = range.Address[false, false, XlReferenceStyle.xlA1, true].Split(new[] { "]" }, StringSplitOptions.None)[1];

                Log.Debug("Recalculate range {@Range} with functionName {@FunctionName}", rangeAdress, functionName);

                excelApp.StatusBar = $"Recalculate {functionName} in {rangeAdress}...";

                range.CalculateRowMajorOrder();
                /*string formula = functionName + "(";
                range.Replace(formula, formula);*/
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