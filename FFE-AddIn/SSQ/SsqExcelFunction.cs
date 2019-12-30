using ExcelDna.Integration;
using System;
using System.Collections.Generic;

namespace FFE
{
    public class SsqExcelFunction
    {
        public SsqExcelFunction(string name, Delegate func, ExcelFunctionAttribute excelFunctionAttribute, ExcelArgumentAttribute excelArgumentAttribute)
        {
            Name = name;
            Delegate = func;
            ExcelFunctionAttribute = excelFunctionAttribute;
            ExcelArgumentAttributes = new List<object>() { excelArgumentAttribute };
        }

        public string Name { get; }

        public Delegate Delegate { get; }

        public object ExcelFunctionAttribute { get; }

        public List<object> ExcelArgumentAttributes { get; }
    }
}