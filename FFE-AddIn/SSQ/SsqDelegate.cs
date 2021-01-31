using ExcelDna.Integration;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FFE
{
    public static class SsqDelegate
    {
        private static readonly ILogger log;

        static SsqDelegate()
        {
            log = Log.ForContext("UDF", "SSQ");
        }

        public static IEnumerable<SsqExcelFunction> GetSsqExcelFunctions(SsqJson ssqJson)
        {
            List<SsqExcelFunction> ssqExcelFunctions = null;

            try
            {
                ssqExcelFunctions = new List<SsqExcelFunction>();

                foreach (KeyValuePair<string, UserDefinedFunction> udf in GetUserDefinedFunctions(ssqJson))
                {
                    UserDefinedFunction userDefinedFunction = udf.Value;
                    QueryInformation queryInformation = userDefinedFunction.QueryInformation;
                    QueryParameter queryParameter = userDefinedFunction.QueryParameter;

                    Delegate excelFunction = new Func<string, object>(wkn_isin_ticker =>
                    {
                        object value;

                        try
                        {
                            Ssq ssq = new Ssq(queryParameter,
                                              udf.Key);

                            value = ssq.TryQ(wkn_isin_ticker);
                        }
                        catch (Exception ex)
                        {
                            log.Error(ex.Message);
                            return ExcelError.ExcelErrorGettingData;
                        }

                        return value;
                    });

                    ExcelFunctionAttribute excelFunctionAttribute = new ExcelFunctionAttribute
                    {
                        Name = queryInformation.Name,
                        Description = queryInformation.Description + "\n" +
                                      "Author: " + queryInformation.Author + " (" + queryInformation.AuthorEmail + ")\n" +
                                      "Version " + queryInformation.Version + " of " + queryInformation.VersionDate + ", Provider: " + queryInformation.Provider,
                        Category = "FFE",
                        IsThreadSafe = true,
                        HelpTopic = !String.IsNullOrEmpty(queryInformation.HelpTopic) ? queryInformation.HelpTopic : null
                    };

                    ExcelArgumentAttribute excelArgumentAttribute = new ExcelArgumentAttribute { Name = queryInformation.ExcelArgNameStockIdentifier, Description = queryInformation.ExcelArgDescStockIdentifier };

                    ssqExcelFunctions.Add(new SsqExcelFunction(queryInformation.Name, excelFunction, excelFunctionAttribute, excelArgumentAttribute));

                    log.Debug("Created SSQ Excel function: {@SsqExcelFunction}", queryInformation.Name);
                }
            }
            catch (Exception ex)
            {
                log.Error("Exception while creating SSQ Excel functions. {@ExceptionMessage}", ex.Message);

                throw;
            }

            return ssqExcelFunctions;
        }

        public static IEnumerable<KeyValuePair<string, UserDefinedFunction>> GetUserDefinedFunctions(SsqJson ssqJson)
        {
            return ssqJson.UDF.Where(x => x.Value.QueryInformation.Enabled);
        }
    }
}