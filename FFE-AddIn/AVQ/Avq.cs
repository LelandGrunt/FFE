using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Newtonsoft.Json.Linq;
using Serilog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;

namespace FFE
{
    public static class Avq
    {
        private const string ADD_IN_NAME = "AVQ";
        private const string URL_ALPHA_VANTAGE_QUERY = "https://www.alphavantage.co/query";

        private static bool SendMissingApiKeyMessage { get; set; } = true;

        public static string UrlContenResult { private get; set; } = null;

        static Avq()
        {
            if (AvqSetting.Default.EnableLogging)
            {
                Log.Logger = FfeLogger.ConfigureLogging(AvqSetting.Default.LogLevel, ADD_IN_NAME)
                                      .ForContext("UDF", typeof(Avq));
            }
        }

        /// <summary>
        /// Alpha Vantage Query for daily equity data.
        /// </summary>
        /// <param name="symbol">The symbol of the stock.</param>
        /// <param name="info">The stock data info to return. 
        ///                    Valid values: Open, High, Low, Close, Volume. Values are case-insensitive. 
        ///                    Default: Close = Closing Price</param>
        /// <param name="tradingDay">The trading day interpreted as an index based value (if positive) or as a trading date subtraction (if negative).
        ///                          Default: 0 = Recent Trading Day</param>
        /// <param name="tradingDate">The trading date. 
        ///                           Default: None</param>
        /// <remarks>Default parameter values are ignored from Excel.</remarks>
        /// <returns>Returns stock data from the Alpha Vantage API (www.alphavantage.co).</returns>
        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVD",
                       Description = "Returns stock data from the Alpha Vantage API (www.alphavantage.co).\nOnly the latest 100 trading dates are supported.",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVD([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                  string symbol,
                                  [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Close, Volume\nDefault: Close = Closing Price")]
                                  string info,
                                  [ExcelArgument(Name = "Trading Day", Description = "positive value => index based trading day\nnegative value => trading date substraction\nDefault: 0 = Recent Trading Day")]
                                  int tradingDay,
                                  [ExcelArgument(Name = "Trading Date", Description = "is the trading date.\nDefault: None")]
                                  DateTime tradingDate)
        {
            object quote = ExcelError.ExcelErrorNA;

            /* "This API returns daily time series (date, daily open, daily high, daily low, daily close, daily volume)
             *  of the equity specified, covering up to 20 years of historical data." */
            const string API_FUNCTION = "TIME_SERIES_DAILY";

            JObject json = null;

            try
            {
                // Check if Alpha Vantage API Key is set, if not exit function.
                if (!CheckApiKey())
                {
                    return ExcelError.ExcelErrorNA;
                }

                // Provided Alpha Vantage Time Series Data(Default is "4. close").
                // For ease of use, the preceding numbering is not necessary.
                switch (info.ToLower())
                {
                    case "open":
                        info = "1. open";
                        break;
                    case "high":
                        info = "2. high";
                        break;
                    case "low":
                        info = "3. low";
                        break;
                    case "close":
                        info = "4. close";
                        break;
                    case "volume":
                        info = "5. volume";
                        break;
                    default:
                        info = "4. close";
                        break;
                }

                // If optional quoteDate parameter is set, then select the data point of given quoteDate.
                string strTradingDate = "";
                if (!tradingDate.Equals(new DateTime(1899, 12, 30, 00, 00, 00)) // Excel default date and time.
                    && !tradingDate.Equals(default))
                {
                    strTradingDate = tradingDate.ToString("yyyy-MM-dd");
                }
                else
                {
                    // Select the most recent data point if 0 (default) was given.
                    if (tradingDay != 0)
                    {
                        // If day is negative, then select data point "current date minus <tradingDay>".
                        if (tradingDay < 0)
                        {
                            strTradingDate = DateTime.Today.AddDays(tradingDay).ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            // Else select the data point at position <tradingDay> (zero - based index).
                            tradingDay = tradingDay - 1;
                        }
                    }
                }

                // API documentation: https://www.alphavantage.co/documentation/#daily
                string url = URL_ALPHA_VANTAGE_QUERY + "?" + "function=" + API_FUNCTION + "&symbol=" + symbol + "&apikey=" + ApiKey;
                Log.Debug("Alpha Vantage query URL: {@AvQueryUrl}.", url);

                string urlContent = null;
                if (UrlContenResult == null)
                {
                    urlContent = FfeWeb.GetHttpResponseContent(url);
                    json = JObject.Parse(urlContent);
                }
                else // else running AVQ tests.
                {
                    json = JObject.Parse(UrlContenResult);
                }

                if (json.ContainsKey("Time Series (Daily)"))
                {
                    /* https://www.alphavantage.co/support/#support
                     * However, we ask that your language-specific library/wrapper preserves the content of our JSON/CSV responses in both success and error cases. */
                    Log.Debug("Alpha Vantage information message: {@AvInformation}", (string)json["Meta Data"]["1. Information"]);

                    // Get JSON object "Time Series (Daily)" with time series data points.
                    JToken timeSeriesDaily = json["Time Series (Daily)"];

                    JToken quoteDate = null;
                    if (!String.IsNullOrEmpty(strTradingDate)) // Get JSON object "Date" by given <tradingDate>.
                    {
                        Log.Debug("Get quote by trading date {@TradingDate}", strTradingDate);
                        quoteDate = timeSeriesDaily[strTradingDate];
                    }
                    else // Get JSON object "Date" by given index <tradingDay>.
                    {
                        Log.Debug("Get quote by trading date index {@TradingDateIndex}", tradingDay);
                        quoteDate = timeSeriesDaily.ElementAt(tradingDay).First;
                    }

                    if (quoteDate != null)
                    {
                        JToken jItem = quoteDate[info];
                        if (jItem != null)
                        {
                            quote = decimal.Parse((string)jItem, CultureInfo.GetCultureInfo("en-US"));
                        }
                        else
                        {
                            Log.Debug("Info {@Info} does not exists.", info);
                        }
                    }
                    else
                    {
                        Log.Debug("Data point {@DataPoint} not found.", (!String.IsNullOrEmpty(strTradingDate) ? strTradingDate : tradingDay.ToString()));
                    }
                }
                else
                {
                    string avNote = (string)json["Note"];
                    if (!String.IsNullOrEmpty(avNote))
                    {
                        Log.Debug("Alpha Vantage note: {@AvNote}", avNote);
                        if (avNote.Contains("API call frequency"))
                        {
                            Log.Debug("Alpha Vantage API call frequency limit has been reached. Please wait and try again later.");
                            return AvqExcelErrorCallLimitReached();
                        }
                    }

                    return ExcelError.ExcelErrorGettingData;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "QAVD");
                if (json != null) { Log.Error("Alpha Vantage JSON response: {@AvResponse}", json.ToString()); }
                return ExcelError.ExcelErrorGettingData;
            }

            return quote;
        }

        public static void QueryAlphaVantageAsBatch(AvqInfo info = AvqInfo.price, bool nonContiguous = false, string namedRange = null)
        {
            Application excelApp = (Application)ExcelDnaUtil.Application;

            try
            {
                excelApp.StatusBar = null;
                excelApp.ScreenUpdating = false;

                Range selection = excelApp.Selection;
                List<Range> ranges = new List<Range>();

                switch (AvqSetting.Default.BatchQueryMode)
                {
                    case "ddcAvqBatchQueryModeItemContiguousRange":
                        nonContiguous = false;
                        break;
                    case "ddcAvqBatchQueryModeItemNonContiguousRange":
                        nonContiguous = true;
                        break;
                    case "ddcAvqBatchQueryModeItemNamedRange":
                        namedRange = AvqSetting.Default.BatchQueryNamedRange;
                        break;
                    default:
                        break;
                }

                if (namedRange != null)
                {
                    foreach (Name name in excelApp.ActiveWorkbook.Names)
                    {
                        if (name.Name.Equals(namedRange)) { ranges.Add(name.RefersToRange); break; }
                        if (namedRange.LastIndexOf("*") > 0
                            && name.Name.Contains(namedRange.Replace("*", "")))
                        {
                            ranges.Add(name.RefersToRange);
                        }
                    }
                }
                else if (nonContiguous)
                {
                    if (selection.Areas.Count % 2 == 1) { throw new AvqException("Selection must contain even number of columns."); }
                    foreach (Range area in selection.Areas)
                    {
                        if (area.Columns.Count != 1) { throw new AvqException("Combination of Contiguous and Non-Contiguous mode is not supported."); }
                    }
                    for (int i = 1; i <= selection.Areas.Count; i += 2)
                    {
                        ranges.Add(selection.Areas[i]);
                    }
                }
                else
                {
                    foreach (Range area in selection.Areas)
                    {
                        if (area.Columns.Count != 2) { throw new AvqException("Combination of Non-Contiguous and Contiguous mode is not supported."); }
                    }
                    ranges.Add(selection);
                }

                int a = 0;
                foreach (Range range in ranges)
                {
                    a += 2;

                    Areas areas = range.Areas;

                    foreach (Range area in areas)
                    {
                        if (area.Cells.Columns.Count != 2 && !nonContiguous)
                        {
                            throw new AvqException("Selection/Named Range must have two columns (first column with symbol, in the second the stock quote is inserted).");
                        }

                        Dictionary<string, dynamic> quotes = new Dictionary<string, dynamic>();

                        string symbol = null;
                        Range columnCells = area.Columns[1].Cells;
                        foreach (Range cell in columnCells)
                        {
                            symbol = cell.Text;
                            if (!String.IsNullOrWhiteSpace(symbol)
                                && !quotes.ContainsKey(symbol.Trim()))
                            {
                                quotes.Add(symbol.Trim(), new { Value = AvqExcelError("N/A"), Timestamp = DateTimeOffset.Now });
                            }
                        }

                        string symbols = quotes.Keys.Aggregate((i, j) => i + "," + j);

                        string url = URL_ALPHA_VANTAGE_QUERY + "?" + "function=" + "BATCH_STOCK_QUOTES" + "&symbols=" + symbols + "&apikey=" + ApiKey;
                        Log.Debug("Alpha Vantage query URL: {@AvQueryUrl}.", url);

                        string urlContent = FfeWeb.GetHttpResponseContent(url);
                        AvqBatchStockQuotes json = AvqBatchStockQuotes.FromJson(urlContent);

                        if (json.Note != null && json.Note.Contains("API call frequency"))
                        {
                            throw new AvqApiCallFrequencyLimitException(json.Note);
                        }
                        if (json.ErrorMessage != null)
                        {
                            throw new AvqException(json.ErrorMessage);
                        }
                        if (json.Information != null)
                        {
                            throw new AvqException(json.Information);
                        }

                        if (json != null)
                        {
                            json.StockQuotes.ForEach(s =>
                            {
                                string infoValue = info == AvqInfo.price ? s.Price : s.Volume;
                                double value = double.Parse(infoValue, CultureInfo.GetCultureInfo("en-us"));
                                quotes[s.Symbol] = new { Value = value, s.Timestamp };
                            });
                        }

                        for (int i = 1; i <= columnCells.Count; i++)
                        {
                            Range cellStock;
                            if (!nonContiguous)
                            {
                                cellStock = columnCells[i, 1];
                            }
                            else
                            {
                                cellStock = selection.Areas[a - 1].Columns[1].Cells[i, 1];
                            }

                            symbol = cellStock.Text;
                            if (!String.IsNullOrWhiteSpace(symbol))
                            {
                                Range cellQuote;
                                if (!nonContiguous)
                                {
                                    cellQuote = columnCells[i, 2];
                                }
                                else
                                {
                                    cellQuote = selection.Areas[a].Columns[1].Cells[i, 1];
                                }

                                var quote = quotes[symbol.Trim()];
                                cellQuote.Value = quote.Value;

                                const string author = "FFE.AVQ:";
                                Comment comment = cellQuote.Comment;
                                if (AvqSetting.Default.BatchComment)
                                {
                                    if (comment == null)
                                    {
                                        comment = cellQuote.AddComment(author);
                                    }
                                    else
                                    {
                                        comment.Text(author);
                                    }
                                    DateTimeOffset timestamp = quote.Timestamp;
                                    string commentText = "\n" + "Timestamp: " + timestamp.LocalDateTime + "\nProvider: alphavantage.co";
                                    comment.Text(commentText, 9);
                                    comment.Shape.TextFrame.Characters(author.Length + 1).Font.Bold = false;
                                    comment.Shape.TextFrame.AutoSize = true;
                                }
                                else
                                {
                                    if (comment != null
                                        && comment.Text().StartsWith(author))
                                    {
                                        comment.Delete();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                excelApp.StatusBar = ex.Message;
                Log.Error(ex.Message);
                Microsoft.VisualBasic.Interaction.MsgBox(ex.Message, MsgBoxStyle.Exclamation, ADD_IN_NAME);
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }
        }

        private static bool CheckApiKey()
        {
            if (String.IsNullOrEmpty(ApiKey)
#if !DEBUG && !TEST
                || ApiKey.Equals("demo")
#endif           
                )
            {
                // Send notification only once in a session.
                if (SendMissingApiKeyMessage)
                {
                    Microsoft.VisualBasic.Interaction.MsgBox("Please set Alpha Vantage API Key." + "\n" + "Go to Ribbon FFE | Group AVQ -> Set Api Key." + "\n" + "Claim your free API Key here: https://www.alphavantage.co/support/#api-key", MsgBoxStyle.OkOnly, ADD_IN_NAME);
                    SendMissingApiKeyMessage = false;
                }
                return false;
            }
            else if (ApiKey.Equals("demo"))
            {
                Log.Warning("Alpha Vantage API Key 'demo' works only with limited symbols. Please request your own Alpha Vantage API Key.");
                return true;
            }
            else
            {
                return true;
            }
        }

        public static string ApiKey
        {
            get
            {
                return AvqSetting.Default.ApiKey;
            }
            set
            {
                AvqSetting.Default.ApiKey = value;
                AvqSetting.Default.Save();
            }
        }

        public enum AvqInfo
        {
            price,
            volume
        }

        #region AVQ Excel Errors
        public static object AvqExcelError(string errorIdentifier = null)
        {
            return "#AV" + (!String.IsNullOrEmpty(errorIdentifier) ? "_" + errorIdentifier : "");
        }
        public static object AvqExcelErrorCallLimitReached()
        {
            return AvqExcelError("CALL_LIMIT_REACHED");
        }
        #endregion

        #region Exceptions
        [Serializable]
        public class AvqException : Exception
        {
            public AvqException() { }

            public AvqException(string message) : base(message) { }

            public AvqException(string message, Exception innerException) : base(message, innerException) { }

            protected AvqException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        }

        [Serializable]
        public class AvqApiCallFrequencyLimitException : Exception
        {
            public AvqApiCallFrequencyLimitException() { }

            public AvqApiCallFrequencyLimitException(string message) : base(message) { }

            public AvqApiCallFrequencyLimitException(string message, Exception innerException) : base(message, innerException) { }

            protected AvqApiCallFrequencyLimitException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        }
        #endregion
    }
}