using ExcelDna.Integration;
using Microsoft.VisualBasic;
using Newtonsoft.Json.Linq;
using Serilog;
using System;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;

namespace FFE
{
    public static partial class Avq
    {
        private const string URL_ALPHA_VANTAGE_QUERY = "https://www.alphavantage.co/query";

        public static bool CallLimitReachedError = false;

        public static string UrlContenResult { private get; set; } = null;

        private static bool SendMissingApiKeyMessage { get; set; } = true;

        private static readonly ILogger log;

        static Avq()
        {
            if (AvqSetting.Default.EnableLogging)
            {
                string udf = "AVQ";
                Log.Logger = FfeLogger.CreateSubLogger(udf, AvqSetting.Default.LogLevel);
                log = Log.ForContext("UDF", udf);
            }
        }

        #region Excel Functions
        /// <remarks>Default parameter values are ignored from Excel.</remarks>
        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVD",
                       Description = "Returns stock data from the Alpha Vantage TIME_SERIES_DAILY API (www.alphavantage.co).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVD([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                  string symbol,
                                  [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Close, Volume (case insensitive)\nDefault: Close = Closing Price")]
                                  string info,
                                  [ExcelArgument(Name = "Trading Day", Description = "positive value => index based trading day\nnegative value => trading date substraction\nDefault: 0 = Recent Trading Day")]
                                  int tradingDay,
                                  [ExcelArgument(Name = "Trading Date", Description = "is the trading date.\nDefault: None")]
                                  DateTime tradingDate,
                                  [ExcelArgument(Name = "Best Match", Description = "Finds best trading data point, if no trading (day/date) argument matches. Valid values: FALSE (= 0), TRUE (<> 0)\nDefault: FALSE")]
                                  bool bestMatch = false,
                                  [ExcelArgument(Name = "Output Size", Description = "The size of returned data points.\nValid values: compact, full (case insensitive)\nDefault: compact")]
                                  string outputSize = "compact")
        {
            // Set default stock data info, if nothing was given.
            info = string.IsNullOrEmpty(info) ? "4. close" : info;

            // Set default output size, if nothing was given.
            AvStockTimeSeriesOutputSize? avStockTimeSeriesOutputSize = AvStockTimeSeriesOutputSize.compact;
            if (!string.IsNullOrEmpty(outputSize))
            {
                avStockTimeSeriesOutputSize = (AvStockTimeSeriesOutputSize)Enum.Parse(typeof(AvStockTimeSeriesOutputSize), outputSize.ToLower());
            }

            return QAV("TIME_SERIES_DAILY", symbol: symbol, info: info, dataPointIndex: tradingDay, dataPointDate: tradingDate, outputSize: avStockTimeSeriesOutputSize, bestMatch: bestMatch);
        }

        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVDA",
                       Description = "Returns stock data from the Alpha Vantage TIME_SERIES_DAILY_ADJUSTED API (www.alphavantage.co).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVDA([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                   string symbol,
                                   [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Close, Adjusted Close, Volume, Dividend Amount, Split Coefficient\nDefault: Close = Adjusted Closing Price")]
                                   string info,
                                   [ExcelArgument(Name = "Trading Day", Description = "positive value => index based trading day\nnegative value => trading date substraction\nDefault: 0 = Recent Trading Day")]
                                   int tradingDay,
                                   [ExcelArgument(Name = "Trading Date", Description = "is the trading date.\nDefault: None")]
                                   DateTime tradingDate,
                                   [ExcelArgument(Name = "Best Match", Description = "Finds best trading data point, if no trading (day/date) argument matches.\nDefault: FALSE")]
                                   bool bestMatch = false,
                                   [ExcelArgument(Name = "Output Size", Description = "The size of returned data points.\nValid values: compact, full\nDefault: compact")]
                                   string outputSize = "compact")
        {
            // Set default stock data info, if nothing was given.
            info = string.IsNullOrEmpty(info) ? "5. adjusted close" : info;

            // Set default output size, if nothing was given.
            AvStockTimeSeriesOutputSize? avStockTimeSeriesOutputSize = AvStockTimeSeriesOutputSize.compact;
            if (!string.IsNullOrEmpty(outputSize))
            {
                avStockTimeSeriesOutputSize = (AvStockTimeSeriesOutputSize)Enum.Parse(typeof(AvStockTimeSeriesOutputSize), outputSize.ToLower());
            }

            return QAV("TIME_SERIES_DAILY_ADJUSTED", symbol: symbol, info: info, dataPointIndex: tradingDay, dataPointDate: tradingDate, outputSize: avStockTimeSeriesOutputSize, bestMatch: bestMatch);
        }

        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVW",
                       Description = "Returns stock data from the Alpha Vantage TIME_SERIES_WEEKLY API (www.alphavantage.co).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVW([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                  string symbol,
                                  [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Close, Volume\nDefault: Close = Closing Price")]
                                  string info,
                                  [ExcelArgument(Name = "Trading Week", Description = "positive value => index based trading week\nnegative value => trading week substraction\nDefault: 0 = Recent Trading Week")]
                                  int tradingWeek,
                                  [ExcelArgument(Name = "Trading Date", Description = "is the trading date.\nDefault: None")]
                                  DateTime tradingDate,
                                  [ExcelArgument(Name = "Best Match", Description = "Finds best trading data point, if no trading (week/date) argument matches.\nDefault: FALSE")]
                                  bool bestMatch = false)
        {
            // Set default stock data info, if nothing was given.
            info = string.IsNullOrEmpty(info) ? "4. close" : info;

            // Subtraction of trading week.
            tradingDate = SubstractTradingWeek(tradingDate, tradingWeek);

            return QAV("TIME_SERIES_WEEKLY", symbol: symbol, info: info, dataPointIndex: tradingWeek, dataPointDate: tradingDate, bestMatch: bestMatch);
        }

        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVWA",
                       Description = "Returns stock data from the Alpha Vantage TIME_SERIES_WEEKLY_ADJUSTED API (www.alphavantage.co).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVWA([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                   string symbol,
                                   [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Close, Adjusted Close, Volume, Dividend Amount\nDefault: Adjusted Close = Adjusted Closing Price")]
                                   string info,
                                   [ExcelArgument(Name = "Trading Week", Description = "positive value => index based trading week\nnegative value => trading week substraction\nDefault: 0 = Recent Trading Week")]
                                   int tradingWeek,
                                   [ExcelArgument(Name = "Trading Date", Description = "is the trading date.\nDefault: None")]
                                   DateTime tradingDate,
                                   [ExcelArgument(Name = "Best Match", Description = "Finds best trading data point, if no trading (week/date) argument matches.\nDefault: FALSE")]
                                   bool bestMatch = false)
        {
            // Set default stock data info, if nothing was given.
            info = string.IsNullOrEmpty(info) ? "5. adjusted close" : info;

            // Subtraction of trading week.
            tradingDate = SubstractTradingWeek(tradingDate, tradingWeek);

            return QAV("TIME_SERIES_WEEKLY_ADJUSTED", symbol: symbol, info: info, dataPointIndex: tradingWeek, dataPointDate: tradingDate, bestMatch: bestMatch);
        }

        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVM",
                       Description = "Returns stock data from the Alpha Vantage TIME_SERIES_MONTHLY API (www.alphavantage.co).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVM([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                  string symbol,
                                  [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Close, Volume\nDefault: Close = Closing Price")]
                                  string info,
                                  [ExcelArgument(Name = "Trading Month", Description = "positive value => index based trading month\nnegative value => trading month substraction\nDefault: 0 = Recent Trading Month")]
                                  int tradingMonth,
                                  [ExcelArgument(Name = "Trading Date", Description = "is the trading date.\nDefault: None")]
                                  DateTime tradingDate,
                                  [ExcelArgument(Name = "Best Match", Description = "Finds best trading data point, if no trading (month/date) argument matches.\nDefault: FALSE")]
                                  bool bestMatch = false)
        {
            // Set default stock data info, if nothing was given.
            info = string.IsNullOrEmpty(info) ? "4. close" : info;

            // Subtraction of trading month.
            tradingDate = SubstractTradingMonth(tradingDate, tradingMonth);

            return QAV("TIME_SERIES_MONTHLY", symbol: symbol, info: info, dataPointIndex: tradingMonth, dataPointDate: tradingDate, bestMatch: bestMatch);
        }

        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVMA",
                       Description = "Returns stock data from the Alpha Vantage TIME_SERIES_MONTHLY_ADJUSTED API (www.alphavantage.co).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVMA([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                   string symbol,
                                   [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Close, Adjusted Close, Volume, Dividend Amount\nDefault: Adjusted Close = Adjusted Closing Price")]
                                   string info,
                                   [ExcelArgument(Name = "Trading Month", Description = "positive value => index based trading month\nnegative value => trading month substraction\nDefault: 0 = Recent Trading Month")]
                                   int tradingMonth,
                                   [ExcelArgument(Name = "Trading Date", Description = "is the trading date.\nDefault: None")]
                                   DateTime tradingDate,
                                   [ExcelArgument(Name = "Best Match", Description = "Finds best trading data point, if no trading (month/date) argument matches.\nDefault: FALSE")]
                                   bool bestMatch = false)
        {
            // Set default stock data info, if nothing was given.
            info = string.IsNullOrEmpty(info) ? "5. adjusted close" : info;

            // Subtraction of trading month.
            tradingDate = SubstractTradingMonth(tradingDate, tradingMonth);

            return QAV("TIME_SERIES_MONTHLY_ADJUSTED", symbol: symbol, info: info, dataPointIndex: tradingMonth, dataPointDate: tradingDate, bestMatch: bestMatch);
        }

        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVQ",
                       Description = "Returns stock data from the Alpha Vantage GLOBAL_QUOTE API (www.alphavantage.co).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVQ([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                  string symbol,
                                  [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Price, Volume, Latest Trading Day, Previous Close, Change, Change Percent\nDefault: Price")]
                                  string info)
        {
            // Symbol is mandatory.
            if (string.IsNullOrEmpty(symbol))
            {
                return ExcelError.ExcelErrorNA;
            }

            // Set default stock data info, if nothing was given.
            info = string.IsNullOrEmpty(info) ? "05. price" : info;

            object stockInfo = ExcelError.ExcelErrorNA;

            JObject json = null;

            try
            {
                // Check if Alpha Vantage API Key is set, if not exit function.
                if (!CheckApiKey())
                {
                    return ExcelError.ExcelErrorNA;
                }

                log.Debug("Querying the stock info {@Info} for {@Symbol} from {@Provider}", info, symbol, "alphavantage.co");

                // API documentation: https://www.alphavantage.co/documentation/#latestprice
                string url = AvUrlBuilder("GLOBAL_QUOTE", symbol: symbol);
                log.Debug("Alpha Vantage query URL: {@AvQueryUrl}.", url);

                json = GetAvJson(url);

                // The "Global Quote" element/token is expected, otherwise something went wrong.
                if (((JProperty)json.Children().ElementAtOrDefault(0))?.Name.Equals("Global Quote") == true)
                {
                    JToken globalQuote = ((JProperty)json.Children().ElementAt(0)).Value;

                    stockInfo = GetStockInfo(globalQuote, info);
                }
                else
                {
                    return ReturnAvError(json);
                }
            }
            catch (Exception ex)
            {
                return CatchAvException("QAVQ", ex, json);
            }

            return stockInfo;
        }

        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVID",
                       Description = "Returns stock data from the Alpha Vantage TIME_SERIES_INTRADAY API (www.alphavantage.co).\nThe function always returns the recent data point.",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVID([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                   string symbol,
                                   [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Close, Volume\nDefault: Open")]
                                   string info,
                                   [ExcelArgument(Name = "Data Point Index", Description = "is the index (zero-based) of the time series data point.\nDefault: 0 = First data point")]
                                   int dataPointIndex = 0,
                                   [ExcelArgument(Name = "Interval", Description = "is the time interval between two data points.\nValid values: [1|5|15|30|60]min\nDefault: 5min")]
                                   string interval = "5min",
                                   [ExcelArgument(Name = "Output Size", Description = "The size of returned data points.\nValid values: compact, full\nDefault: compact")]
                                   string outputSize = "compact")
        {
            // Symbol is mandatory.
            if (string.IsNullOrEmpty(symbol))
            {
                return ExcelError.ExcelErrorNA;
            }

            if (!string.IsNullOrEmpty(interval)
                && !interval.Equals("1min")
                && !interval.Equals("5min")
                && !interval.Equals("15min")
                && !interval.Equals("30min")
                && !interval.Equals("60min"))
            {
                log.Error("Interval {@Interval} is not supported", interval);
                return ExcelError.ExcelErrorNA;
            }
            else
            {
                // Set default interval, if nothing was given.
                interval = "5min";
            }

            // Set default stock data info, if nothing was given.
            info = string.IsNullOrEmpty(info) ? "1. open" : info;

            // Set default output size, if nothing was given.
            AvStockTimeSeriesOutputSize? avStockTimeSeriesOutputSize = AvStockTimeSeriesOutputSize.compact;
            if (!string.IsNullOrEmpty(outputSize))
            {
                avStockTimeSeriesOutputSize = (AvStockTimeSeriesOutputSize)Enum.Parse(typeof(AvStockTimeSeriesOutputSize), outputSize.ToLower());
            }

            object stockInfo;

            JObject json = null;

            try
            {
                // Check if Alpha Vantage API Key is set, if not exit function.
                if (!CheckApiKey())
                {
                    return ExcelError.ExcelErrorNA;
                }

                log.Debug("Querying the stock info {@Info} for {@Symbol} from {@Provider}", info, symbol, "alphavantage.co");

                // API documentation: https://www.alphavantage.co/documentation/#intraday
                string url = AvUrlBuilder("TIME_SERIES_INTRADAY", symbol: symbol, outputSize: avStockTimeSeriesOutputSize, interval: interval);
                log.Debug("Alpha Vantage query URL: {@AvQueryUrl}.", url);

                json = GetAvJson(url);

                // The second element/token is always a time series object with the data points, if not, something went wrong.
                if (((JProperty)json.Children().ElementAtOrDefault(1))?.Name.Contains("Time Series") == true)
                {
                    JToken firstTimeSerie = ((JProperty)((JProperty)json.Children().ElementAt(1)).Value.ElementAt(dataPointIndex)).Value;

                    stockInfo = GetStockInfo(firstTimeSerie, info);
                }
                else
                {
                    return ReturnAvError(json);
                }
            }
            catch (Exception ex)
            {
                return CatchAvException("QAVID", ex, json);
            }

            return stockInfo;
        }

        [FfeFunction(Provider = "alphavantage.co")]
        [ExcelFunction(Name = "QAVTS",
                       Description = "Returns stock data from the Alpha Vantage TIME_SERIES_DAILY API (www.alphavantage.co).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QAVTS([ExcelArgument(Name = "Symbol", Description = "is the symbol of the stock.")]
                                   string symbol,
                                   [ExcelArgument(Name = "Info", Description = "is the stock data info to return.\nValid values: Open, High, Low, Close, Adjusted Close, Volume, Dividend Amount (case insensitive)\nDefault: Close = Closing Price")]
                                   string info,
                                   [ExcelArgument(Name = "Interval", Description = "Defines the AlphaVantage API to use.\nValid values: Daily, Weekly, Monthly, 1min, 5min, 15min, 30min, 60min\nDefault: Daily")]
                                   string interval,
                                   [ExcelArgument(Name = "Trading Day", Description = "positive value => index based trading day\nnegative value => trading date substraction\nDefault: 0 = Recent Trading Day")]
                                   int tradingDay,
                                   [ExcelArgument(Name = "Trading Date", Description = "is the trading date.\nDefault: None")]
                                   DateTime tradingDate,
                                   [ExcelArgument(Name = "Adjusted", Description = "is a flag for querying adjusted values or not.\nValid values: FALSE (= 0), TRUE (<> 0)\nDefault: FALSE")]
                                   bool adjusted = false,
                                   [ExcelArgument(Name = "Output Size", Description = "The size of returned data points.\nValid values: compact, full (case insensitive)\nDefault: compact")]
                                   string outputSize = "compact",
                                   [ExcelArgument(Name = "Best Match", Description = "Finds best trading data point, if no trading (day/date) argument matches. Valid values: FALSE (= 0), TRUE (<> 0)\nDefault: FALSE")]
                                   bool bestMatch = false)
        {
            string api = null;

            // Set default interval, if nothing was given.
            interval = string.IsNullOrEmpty(interval) ? "daily" : interval.ToLower();
            if (interval.Equals("1min")
                || interval.Equals("5min")
                || interval.Equals("15min")
                || interval.Equals("30min")
                || interval.Equals("60min"))
            {
                return QAVID(symbol: symbol, info: info, interval: interval, outputSize: outputSize);
            }
            else if (interval.Equals("daily")
                     || interval.Equals("weekly")
                     || interval.Equals("monthly"))
            {
                // Set default output size, if nothing was given.
                AvStockTimeSeriesOutputSize? avStockTimeSeriesOutputSize = AvStockTimeSeriesOutputSize.compact;
                if (!string.IsNullOrEmpty(outputSize))
                {
                    avStockTimeSeriesOutputSize = (AvStockTimeSeriesOutputSize)Enum.Parse(typeof(AvStockTimeSeriesOutputSize), outputSize.ToLower());
                }

                api = $"TIME_SERIES_{interval.ToUpper()}" + (adjusted ? "_ADJUSTED" : "");
                return QAV(api, symbol: symbol, info: info, dataPointIndex: tradingDay, dataPointDate: tradingDate, outputSize: avStockTimeSeriesOutputSize, bestMatch: bestMatch);
            }
            else
            {
                log.Error("Interval {@Interval} is not supported", interval);
                return ExcelError.ExcelErrorNA;
            }
        }
        #endregion

        private static object QAV(string alphaVantageApi,
                                  string symbol,
                                  string info = "4. close",
                                  int dataPointIndex = 0,
                                  DateTime dataPointDate = default,
                                  AvStockTimeSeriesOutputSize? outputSize = null,
                                  bool bestMatch = false)
        {
            // Symbol is mandatory.
            if (string.IsNullOrEmpty(symbol))
            {
                return ExcelError.ExcelErrorNA;
            }

            object stockInfo = ExcelError.ExcelErrorNA;

            JObject json = null;

            try
            {
                // Check if Alpha Vantage API Key is set, if not exit function.
                if (!CheckApiKey())
                {
                    return ExcelError.ExcelErrorNA;
                }

                // If optional quoteDate parameter is set, then select the data point of given quoteDate.
                string strTradingDate = "";
                if (!dataPointDate.Equals(new DateTime(1899, 12, 30, 00, 00, 00)) // Excel default date and time.
                    && !dataPointDate.Equals(default))
                {
                    strTradingDate = dataPointDate.ToString("yyyy-MM-dd");
                }
                else
                {
                    // Select the most recent data point if 0 (default) was given.
                    if (dataPointIndex != 0)
                    {
                        // If day is negative, then select data point "current date minus <tradingDay>".
                        if (dataPointIndex < 0)
                        {
                            strTradingDate = DateTime.Today.AddDays(dataPointIndex).ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            // Else select the data point at position <tradingDay> (zero - based index).
                            dataPointIndex = dataPointIndex - 1;
                        }
                    }
                }

                log.Debug("Querying the stock info {@Info} for {@Symbol} from {@Provider}", info, symbol, "alphavantage.co");

                // API documentation: https://www.alphavantage.co/documentation/#time-series-data
                string url = AvUrlBuilder(alphaVantageApi, symbol: symbol, outputSize: outputSize);
                log.Debug("Alpha Vantage query URL: {@AvQueryUrl}.", url);

                json = GetAvJson(url);

                // The second element/token is always a time series object with the data points, if not, something went wrong.
                if (((JProperty)json.Children().ElementAtOrDefault(1))?.Name.Contains("Time Series") == true)
                {
                    /* https://www.alphavantage.co/support/#support
                     * However, we ask that your language-specific library/wrapper preserves the content of our JSON/CSV responses in both success and error cases. */
                    log.Debug("Alpha Vantage information message: {@AvInformation}", (string)json["Meta Data"]["1. Information"]);

                    // Get "Time Series" JSON object with time series data points.
                    JToken timeSeries = ((JProperty)json.Children().ElementAt(1)).Value;

                    JToken quoteDate = null;
                    if (!String.IsNullOrEmpty(strTradingDate)) // Get JSON object "Date" by given <tradingDate>.
                    {
                        log.Debug("Get quote by trading date {@TradingDate}", strTradingDate);
                        quoteDate = timeSeries[strTradingDate];

                        // Find the first trading date, that matches best.
                        if (quoteDate == null
                            && bestMatch)
                        {
                            log.Debug("Trading date {@TradingDate} not found. Select nearest one (Best Match enabled).", strTradingDate);
                            quoteDate = ((JProperty)timeSeries.First(td => DateTime.Parse(td.Value<JProperty>().Name) < DateTime.Parse(strTradingDate))).Value;
                        }
                    }
                    else // Get JSON object "Date" by given index <tradingDay>.
                    {
                        log.Debug("Get quote by trading date index {@TradingDateIndex}", dataPointIndex + 1);
                        quoteDate = timeSeries.ElementAtOrDefault(dataPointIndex);
                        quoteDate = quoteDate?.First;

                        // Get last trading date/data point if index is out of range.
                        if (quoteDate == null
                            && bestMatch)
                        {
                            log.Debug("Data point at index {@TradingDateIndex} not found. Select last one (Best Match enabled).", dataPointIndex + 1);
                            quoteDate = ((JProperty)timeSeries.Last).Value;
                        }
                    }

                    if (quoteDate != null)
                    {
                        stockInfo = GetStockInfo(quoteDate, info);

                    }
                    else
                    {
                        log.Debug("Data point {@DataPoint} not found.", (!String.IsNullOrEmpty(strTradingDate) ? strTradingDate : (dataPointIndex + 1).ToString()));
                    }
                }
                else
                {
                    return ReturnAvError(json);
                }
            }
            catch (Exception ex)
            {
                return CatchAvException("QAV", ex, json);
            }

            return stockInfo;
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
                    Interaction.MsgBox("Please set Alpha Vantage API Key." + "\n" + "Go to Ribbon FFE | Group AVQ -> Set Api Key." + "\n" + "Claim your free API Key here: https://www.alphavantage.co/support/#api-key", MsgBoxStyle.OkOnly, typeof(Avq).Name.ToUpper());
                    SendMissingApiKeyMessage = false;
                }
                return false;
            }
            else if (ApiKey.Equals("demo"))
            {
                log.Warning("Alpha Vantage API Key 'demo' works only with limited symbols. Please request your own Alpha Vantage API Key.");
                return true;
            }
            else
            {
                return true;
            }
        }

        public static string AvUrlBuilder(string api, string apiKey = null, string symbol = null, AvStockTimeSeriesOutputSize? outputSize = null, string interval = null)
        {
            return URL_ALPHA_VANTAGE_QUERY + "?"
                   + $"function={api}"
                   + (!string.IsNullOrEmpty(symbol) ? $"&symbol={symbol}" : null)
                   + (outputSize.HasValue ? $"&outputsize={outputSize.Value}" : null)
                   + (!string.IsNullOrEmpty(interval) ? $"&interval={interval}" : null)
                   + $"&apikey={apiKey ?? ApiKey}";
        }

        // Avoids duplicate codes.
        private static DateTime SubstractTradingWeek(DateTime tradingDate, int tradingWeek)
        {
            // Subtraction of trading week (only if no trading date is given = overruling).
            if ((tradingDate.Equals(new DateTime(1899, 12, 30, 00, 00, 00)) // Excel default date and time.
                || tradingDate.Equals(default))
                && tradingWeek < 0)
            {
                tradingDate = DateTime.Today;
                tradingDate = tradingDate.AddDays(tradingWeek * 7);
            }
            return tradingDate;
        }

        // Avoids duplicate codes.
        private static DateTime SubstractTradingMonth(DateTime tradingDate, int tradingMonth)
        {
            // Subtraction of trading month (only if no trading date is given = overruling).
            if ((tradingDate.Equals(new DateTime(1899, 12, 30, 00, 00, 00)) // Excel default date and time.
                || tradingDate.Equals(default))
                && tradingMonth < 0)
            {
                tradingDate = DateTime.Today;
                tradingDate = tradingDate.AddMonths(tradingMonth);
            }
            return tradingDate;
        }

        // Avoids duplicate codes.
        private static JObject GetAvJson(string url)
        {
            if (UrlContenResult == null)
            {
                return JObject.Parse(FfeWeb.GetHttpResponseContent(url));
            }
            else // else running AVQ tests.
            {
                return JObject.Parse(UrlContenResult);
            }
        }

        // Avoids duplicate codes.
        private static object GetStockInfo(JToken timeSerie, string info)
        {
            object stockInfo;

            JToken jStockInfo = ((JProperty)timeSerie.Value<JToken>().FirstOrDefault(i => ((JProperty)i).Name.ToLower().Contains(info.ToLower()))).Value;
            if (jStockInfo != null)
            {
                if (info.Contains("latest trading day"))
                {
                    stockInfo = DateTime.Parse((string)jStockInfo, CultureInfo.GetCultureInfo("en-US"));
                }
                else if (info.Contains("change percent"))
                {
                    stockInfo = jStockInfo.ToString();
                    // TOOD: Divide by 100?
                    //quote = decimal.Parse(jStockInfo.ToString().TrimEnd('%'), CultureInfo.GetCultureInfo("en-US"));
                }
                else
                {
                    stockInfo = decimal.Parse((string)jStockInfo, CultureInfo.GetCultureInfo("en-US"));
                }
            }
            else
            {
                log.Debug("Stock data info {@Info} not found.", info);
                stockInfo = ExcelError.ExcelErrorNA;
            }

            return stockInfo;
        }

        // Avoids duplicate codes.
        private static object ReturnAvError(JObject json)
        {
            string avNote = (string)json["Note"];
            if (!String.IsNullOrEmpty(avNote))
            {
                log.Debug("Alpha Vantage note: {@AvNote}", avNote);
                if (avNote.Contains("API call frequency"))
                {
                    CallLimitReachedError = true;

                    log.Debug("Alpha Vantage API call frequency limit has been reached. Please wait and try again later.");
                    return AvqExcelErrorCallLimitReached();
                }
            }

            return ExcelError.ExcelErrorGettingData;
        }

        // Avoids duplicate codes.
        private static object CatchAvException(string function, Exception ex, JObject json)
        {
            log.Error(ex, function);

            if (json != null)
            {
                log.Error("Alpha Vantage JSON response: {@AvResponse}", json.ToString());
            }

            return ExcelError.ExcelErrorGettingData;
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

        #region Enumerations
        public enum AvStockTimeSeriesOutputSize
        {
            compact,
            full
        }
        #endregion

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