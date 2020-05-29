using ExcelDna.Integration;
using Serilog;
using System;
using System.Collections.Generic;

namespace FFE
{
    public static class SsqRegistration
    {
        static SsqRegistration()
        {
            FfeLogger.CreateDefaultLogger("SSQ");
        }

        public static void RegisterFunctions(this IEnumerable<SsqExcelFunction> ssqExcelFunctions)
        {
            List<Delegate> delegates = new List<Delegate>();
            List<object> excelFunctionAttributes = new List<object>();
            List<List<object>> excelArgumentAttributes = new List<List<object>>();

            foreach (SsqExcelFunction ssqExcelFunction in ssqExcelFunctions)
            {
                try
                {
                    delegates.Add(ssqExcelFunction.Delegate);
                    excelFunctionAttributes.Add(ssqExcelFunction.ExcelFunctionAttribute);
                    excelArgumentAttributes.Add(ssqExcelFunction.ExcelArgumentAttributes);

                    Log.Debug("Registered SSQ Excel function: {@SsqExcelFunction}", ((ExcelFunctionAttribute)ssqExcelFunction.ExcelFunctionAttribute).Name);
                }
                catch (Exception ex)
                {
                    Log.Error("Exception while registering SSQ Excel function {@SsqExcelFunction}. Message: {@ExceptionMessage}", ssqExcelFunction.Delegate.Method.Name, ex.Message);
                }
            }

            ExcelIntegration.RegisterDelegates(delegates, excelFunctionAttributes, excelArgumentAttributes);
        }
    }
}