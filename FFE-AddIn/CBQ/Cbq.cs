using ExcelDna.Integration;
using Serilog;
using System;
using System.Globalization;

namespace FFE
{
    public static class Cbq
    {
        private static readonly ILogger log;

        static Cbq()
        {
            if (CbqSetting.Default.EnableLogging)
            {
                string udf = "CBQ";
                Log.Logger = FfeLogger.CreateSubLogger(udf, CbqSetting.Default.LogLevel);
                log = Log.ForContext("UDF", udf);
            }
        }

        /// <summary>
        /// Consorsbank Query for stock data.
        /// </summary>
        /// <param name="isin_wkn">The International Securities Identification Number or Wertpapierkennnummer (Germany).</param>
        /// <param name="exchange">The stock exchange.
        ///                        Example: GAT (=Tradegate)</param>
        /// <param name="info">The info to query. 
        ///                    Valid values: Price. 
        ///                    Default: Price</param>
        /// <returns>Returns stock data from the Consorsbank website.</returns>
        [FfeFunction(Provider = "consorsbank.de")]
        [ExcelFunction(Name = "QCB",
                       Description = "Function to get a stock value from the Consorsbank website",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QCB([ExcelArgument(Name = "ISIN/WKN", Description = "International Securities Identification Number or Wertpapierkennnummer (Germany)")]
                                 string isin_wkn,
                                 [ExcelArgument(Name = "Stock Exchange", Description = "Stock exchange.\nExample: GAT (=Tradegate)")]
                                 string exchange = "",
                                 [ExcelArgument(Name = "Info", Description = "Info to query.\nValid values: Price\nDefault: Price")]
                                 string info = "price")
        {
            return QueryConsorsbank(isin_wkn, exchange, info).Value;
        }

        /// <summary>
        /// Consorsbank Query for formatted stock data.
        /// </summary>
        /// <param name="isin_wkn">The International Securities Identification Number or Wertpapierkennnummer (Germany).</param>
        /// <param name="exchange">The stock exchange.
        ///                        Example: GAT (=Tradegate)</param>
        /// <param name="info">The info to query. 
        ///                    Valid values: Price. 
        ///                    Default: Price</param>
        /// <returns>Returns formatted stock data from the Consorsbank website.</returns>
        [ExcelFunction(Name = "QCBF",
                       Description = "Function to get a stock value from the Consorsbank website.\nFormats the value in the unit (e.g. currency) of stock value, if available.",
                       Category = "FFE",
                       IsThreadSafe = true,
                       IsMacroType = true)]
        public static object QCBF([ExcelArgument(Name = "ISIN/WKN", Description = "International Securities Identification Number or Wertpapierkennnummer (Germany)")]
                                  string isin_wkn,
                                  [ExcelArgument(Name = "Stock Exchange", Description = "Stock exchange.\nExample: GAT (=Tradegate)")]
                                  string exchange = "",
                                  [ExcelArgument(Name = "Info", Description = "Info to query.\nValid values: Price\nDefault: Price")]
                                  string info = "price")
        {
            var value = QueryConsorsbank(isin_wkn, exchange, info);

            try
            {
                // Price
                if (value.CbqInfo == CbqInfo.price)
                {
                    string xPath = CbqSetting.Default.QCB_Price_XPath_Currency;

                    string currency = value.WebParser.SelectByXPath(xPath);

                    if (!String.IsNullOrWhiteSpace(currency))
                    {
                        //TODO: Test if method is invoked by Excel formula.
                        ExcelReference reference = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
                        Microsoft.Office.Interop.Excel.Range refRange = FfeExcel.ReferenceToRange(reference);
                        refRange.NumberFormat = CbqSetting.Default.QCB_Price_NumberFormat.Replace("{CURRENCY}", currency);
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex.Source.Contains("Excel")
                    && ex.Message.Contains("NumberFormat"))
                {
                    // Ignore exception "Unable to set the NumberFormat property of the Range class".
                }
                else
                {
                    log.Error(ex.Message);
                    return CbqExcelError("FORMAT_ERROR");
                }
            }

            return value.Value;
        }

        private static (object Value, CbqInfo CbqInfo, IFfeWebParser WebParser) QueryConsorsbank(string isin_wkn,
                                                                                                 string exchange,
                                                                                                 string info)
        {
            object value = ExcelError.ExcelErrorNA;
            CbqInfo cbqInfo = CbqInfo.price; // Default
            IFfeWebParser webParser = null;

            try
            {
                if (!String.IsNullOrEmpty(info))
                {
                    info = info.ToLower();
                    cbqInfo = (CbqInfo)Enum.Parse(typeof(CbqInfo), info);
                    if (!(Enum.IsDefined(typeof(CbqInfo), cbqInfo)))
                        throw new NotSupportedException("Not supported Info parameter value");
                }

                // Price
                if (cbqInfo == CbqInfo.price)
                {
                    string url = CbqSetting.Default.QCB_URL;
                    url = url.Replace("{ISIN_WKN}", isin_wkn);
                    if (!String.IsNullOrEmpty(exchange))
                    {
                        url += "?exchange={EXCHANGE}";
                        url = url.Replace("{EXCHANGE}", exchange);
                    }

                    log.Debug("Querying the stock info {@Info} for {@IsinWkn} from {@Provider}", cbqInfo.ToString(), isin_wkn, "consorsbank.de");

                    webParser = new FfeWebAngleSharp(new Uri(url));
                    string xPath = CbqSetting.Default.QCB_Price_XPath_Price;
                    string quote = webParser.SelectByXPath(xPath);

                    decimal price = decimal.Parse(quote, new CultureInfo("de-DE"));

                    value = price;
                }
            }
            catch (FormatException fex)
            {
                log.Error(fex.Message);
                log.Debug("Could not parse {@QuoteValue} to decimal type", value);
                return (Value: ExcelError.ExcelErrorGettingData,
                        CbqInfo: cbqInfo,
                        WebParser: null);
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                return (Value: ExcelError.ExcelErrorGettingData,
                        CbqInfo: cbqInfo,
                        WebParser: null);
            }

            return (value, cbqInfo, webParser);
        }

        private enum CbqInfo
        {
            price
        }

        #region CBQ Excel Errors
        public static object CbqExcelError(string errorIdentifier = null)
        {
            return "#CB" + (!String.IsNullOrEmpty(errorIdentifier) ? "_" + errorIdentifier : "");
        }
        #endregion
    }
}