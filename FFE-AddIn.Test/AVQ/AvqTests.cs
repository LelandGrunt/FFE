using ExcelDna.Integration;
using Nager.Date;
using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace FFE
{
    public class QAVTest
    {
        public readonly string ApiKey = "demo";
        public readonly string Symbol = "MSFT";
        public object Expected { get; set; }
        public string Info { get; set; } = "close";
        // Day, Week or Month index or substraction.
        public int DatePart { get; set; } = 0;
        public DateTime QuoteDate { get; set; } = default(DateTime);
        public string OutputSize { get; set; } = "compact";
        public bool BestMatch { get; set; } = true;
        public string Interval { get; set; } = "5min";
        public bool Adjusted { get; set; } = false;
    }

    public class AvqTests
    {
        private readonly string AvTimeSeriesGlobalQuoteJsonTestFile = @"AVQ\GLOBAL_QUOTE.json";
        private readonly string AvTimeSeriesDailyJsonTestFile = @"AVQ\TIME_SERIES_DAILY.json";
        private readonly string AvTimeSeriesDailyAdjustedJsonTestFile = @"AVQ\TIME_SERIES_DAILY_ADJUSTED.json";
        private readonly string AvTimeSeriesIntradayJsonTestFile = @"AVQ\TIME_SERIES_INTRADAY.json";
        private readonly string AvTimeSeriesMonthlyJsonTestFile = @"AVQ\TIME_SERIES_MONTHLY.json";
        private readonly string AvTimeSeriesMonthlyAdjustedJsonTestFile = @"AVQ\TIME_SERIES_MONTHLY_ADJUSTED.json";
        private readonly string AvTimeSeriesWeeklyJsonTestFile = @"AVQ\TIME_SERIES_WEEKLY.json";
        private readonly string AvTimeSeriesWeeklyAdjustedJsonTestFile = @"AVQ\TIME_SERIES_WEEKLY_ADJUSTED.json";

        public AvqTests() { }

        #region Offline tests
        // Test data for GLOBAL_QUOTE API.
        public static IEnumerable<object[]> GetQAVQTestData()
        {
            yield return new object[] { new QAVTest { Expected = 166.1900M, Info = "open" } };
            yield return new object[] { new QAVTest { Expected = 166.5900M, Info = "high" } };
            yield return new object[] { new QAVTest { Expected = 165.2700M, Info = "low" } };
            yield return new object[] { new QAVTest { Expected = 166.2100M, Info = "price" } };
            yield return new object[] { new QAVTest { Expected = 2942132M, Info = "volume" } };
            yield return new object[] { new QAVTest { Expected = new DateTime(2020, 1, 23), Info = "latest trading day" } };
            yield return new object[] { new QAVTest { Expected = 165.7000M, Info = "previous close" } };
            yield return new object[] { new QAVTest { Expected = 0.5100M, Info = "change" } };
            yield return new object[] { new QAVTest { Expected = "0.3078%", Info = "change percent" } };
        }

        // Test data for TIME_SERIES_DAILY API.
        public static IEnumerable<object[]> GetQAVDTestData()
        {
            yield return new object[] { new QAVTest { Expected = 166.1900M, Info = "open", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.5900M, Info = "high", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 165.2700M, Info = "low", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.4550M, Info = "close", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 3443378M, Info = "volume", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 165.7000M, Info = "close", DatePart = 2 } };
            // tradingDay with negative value is not testable because current date (Today) is used for subtraction, which is not compatible with offline file.
            //yield return new object[] { new QAVTest { Expected = 166.5000M, Info = "close", Day = -3 } };
            yield return new object[] { new QAVTest { Expected = 157.5900M, Info = "close", DatePart = 2, QuoteDate = new DateTime(2019, 12, 30) } };
            yield return new object[] { new QAVTest { Expected = 158.9600M, Info = "close", QuoteDate = new DateTime(2019, 12, 27) } };
        }

        // Test data for TIME_SERIES_DAILY_ADJUSTED API.
        public static IEnumerable<object[]> GetQAVDATestData()
        {
            yield return new object[] { new QAVTest { Expected = 166.1900M, Info = "open", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.5900M, Info = "high", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 165.2700M, Info = "low", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.3200M, Info = "close", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.3200M, Info = "adjusted close", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 2667463M, Info = "volume", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 0.0000M, Info = "dividend amount", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 1.0000M, Info = "split coefficient", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 165.7000M, Info = "adjusted close", DatePart = 2 } };
            // tradingDay with negative value is not testable because current date (Today) is used for subtraction, which is not compatible with offline file.
            //yield return new object[] { new QAVTest { Expected = 166.5000M, Info = "adjusted close", Day = -3 } };
            yield return new object[] { new QAVTest { Expected = 157.5900M, Info = "adjusted close", DatePart = 2, QuoteDate = new DateTime(2019, 12, 30) } };
            yield return new object[] { new QAVTest { Expected = 158.9600M, Info = "adjusted close", QuoteDate = new DateTime(2019, 12, 27) } };
        }

        // Test data for TIME_SERIES_INTRADAY API.
        public static IEnumerable<object[]> GetQAVIDTestData()
        {
            yield return new object[] { new QAVTest { Expected = 166.5000M, Info = "open" } };
            yield return new object[] { new QAVTest { Expected = 166.5568M, Info = "high" } };
            yield return new object[] { new QAVTest { Expected = 166.2900M, Info = "low", } };
            yield return new object[] { new QAVTest { Expected = 166.3500M, Info = "close" } };
            yield return new object[] { new QAVTest { Expected = 172371M, Info = "volume" } };
            yield return new object[] { new QAVTest { Expected = 166.3500M, Info = "close", DatePart = 3 } };
        }

        // Test data for TIME_SERIES_WEEKLY API.
        public static IEnumerable<object[]> GetQAVWTestData()
        {
            yield return new object[] { new QAVTest { Expected = 166.6800M, Info = "open", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 168.1900M, Info = "high", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 165.2700M, Info = "low", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.3300M, Info = "close", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 55998384M, Info = "volume", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 167.1000M, Info = "close", DatePart = 2 } };
            // tradingWeek with negative value is not testable because current date (Today) is used for subtraction, which is not compatible with offline file.
            //yield return new object[] { new QAVTest { Expected = 161.3400M, Info = "close", Day = -3 } };
            yield return new object[] { new QAVTest { Expected = 158.9600M, Info = "close", DatePart = 2, QuoteDate = new DateTime(2019, 12, 27) } };
            yield return new object[] { new QAVTest { Expected = 157.4100M, Info = "close", QuoteDate = new DateTime(2019, 12, 20) } };
        }

        // Test data for TIME_SERIES_WEEKLY_ADJUSTED API.
        public static IEnumerable<object[]> GetQAVWATestData()
        {
            yield return new object[] { new QAVTest { Expected = 166.6800M, Info = "open", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 168.1900M, Info = "high", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 165.2700M, Info = "low", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.3600M, Info = "close", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.3600M, Info = "adjusted close", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 56039452M, Info = "volume", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 0.0000M, Info = "dividend amount", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 167.1000M, Info = "adjusted close", DatePart = 2 } };
            // tradingWeek with negative value is not testable because current date (Today) is used for subtraction, which is not compatible with offline file.
            //yield return new object[] { new QAVTest { Expected = 161.3400M, Info = "adjusted close", Day = -3 } };
            yield return new object[] { new QAVTest { Expected = 158.9600M, Info = "adjusted close", DatePart = 2, QuoteDate = new DateTime(2019, 12, 27) } };
            yield return new object[] { new QAVTest { Expected = 157.4100M, Info = "adjusted close", QuoteDate = new DateTime(2019, 12, 20) } };
        }

        // Test data for TIME_SERIES_MONTHLY API.
        public static IEnumerable<object[]> GetQAVMTestData()
        {
            yield return new object[] { new QAVTest { Expected = 158.7800M, Info = "open", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 168.1900M, Info = "high", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 156.5100M, Info = "low", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.4000M, Info = "close", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 334269131M, Info = "volume", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 157.7000M, Info = "close", DatePart = 2 } };
            // tradingMonth with negative value is not testable because current date (Today) is used for subtraction, which is not compatible with offline file.
            //yield return new object[] { new QAVTest { Expected = 151.3800M, Info = "close", Day = -3 } };
            yield return new object[] { new QAVTest { Expected = 157.7000M, Info = "close", DatePart = 2, QuoteDate = new DateTime(2019, 12, 31) } };
            yield return new object[] { new QAVTest { Expected = 151.3800M, Info = "close", QuoteDate = new DateTime(2019, 11, 29) } };
        }

        // Test data for TIME_SERIES_MONTHLY_ADJUSTED API.
        public static IEnumerable<object[]> GetQAVMATestData()
        {
            yield return new object[] { new QAVTest { Expected = 158.7800M, Info = "open", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 168.1900M, Info = "high", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 156.5100M, Info = "low", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.2800M, Info = "close", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 166.2800M, Info = "adjusted close", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 334329837M, Info = "volume", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 0.0000M, Info = "dividend amount", DatePart = 1 } };
            yield return new object[] { new QAVTest { Expected = 157.7000M, Info = "adjusted close", DatePart = 2 } };
            // tradingMonth with negative value is not testable because current date (Today) is used for subtraction, which is not compatible with offline file.
            //yield return new object[] { new QAVTest { Expected = 151.3800M, Info = "adjusted close", Day = -3 } };
            yield return new object[] { new QAVTest { Expected = 157.7000M, Info = "adjusted close", DatePart = 2, QuoteDate = new DateTime(2019, 12, 31) } };
            yield return new object[] { new QAVTest { Expected = 151.3800M, Info = "adjusted close", QuoteDate = new DateTime(2019, 11, 29) } };
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "GLOBAL_QUOTE")]
        [Theory]
        [MemberData(nameof(GetQAVQTestData))]
        public void QAVQTest(QAVTest test)
        {
            Avq.UrlContenResult = File.ReadAllText(AvTimeSeriesGlobalQuoteJsonTestFile);
            object actual = Avq.QAVQ(test.Symbol, info: test.Info);
            Assert.Equal(test.Expected, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_DAILY")]
        [Theory]
        [MemberData(nameof(GetQAVDTestData))]
        public void QAVDTest(QAVTest test)
        {
            Avq.UrlContenResult = File.ReadAllText(AvTimeSeriesDailyJsonTestFile);
            object actual = Avq.QAVD(test.Symbol, info: test.Info, tradingDay: test.DatePart, tradingDate: test.QuoteDate);
            Assert.Equal(test.Expected, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_DAILY_ADJUSTED")]
        [Theory]
        [MemberData(nameof(GetQAVDATestData))]
        public void QAVDATest(QAVTest test)
        {
            Avq.UrlContenResult = File.ReadAllText(AvTimeSeriesDailyAdjustedJsonTestFile);
            object actual = Avq.QAVDA(test.Symbol, info: test.Info, tradingDay: test.DatePart, tradingDate: test.QuoteDate);
            Assert.Equal(test.Expected, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_INTRADAY")]
        [Theory]
        [MemberData(nameof(GetQAVIDTestData))]
        public void QAVIDTest(QAVTest test)
        {
            Avq.UrlContenResult = File.ReadAllText(AvTimeSeriesIntradayJsonTestFile);
            object actual = Avq.QAVID(test.Symbol, info: test.Info);
            Assert.Equal(test.Expected, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_WEEKLY")]
        [Theory]
        [MemberData(nameof(GetQAVWTestData))]
        public void QAVWTest(QAVTest test)
        {
            Avq.UrlContenResult = File.ReadAllText(AvTimeSeriesWeeklyJsonTestFile);
            object actual = Avq.QAVW(test.Symbol, info: test.Info, tradingWeek: test.DatePart, tradingDate: test.QuoteDate);
            Assert.Equal(test.Expected, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_WEEKLY_ADJUSTED")]
        [Theory]
        [MemberData(nameof(GetQAVWATestData))]
        public void QAVWATest(QAVTest test)
        {
            Avq.UrlContenResult = File.ReadAllText(AvTimeSeriesWeeklyAdjustedJsonTestFile);
            object actual = Avq.QAVWA(test.Symbol, info: test.Info, tradingWeek: test.DatePart, tradingDate: test.QuoteDate);
            Assert.Equal(test.Expected, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_MONTHLY")]
        [Theory]
        [MemberData(nameof(GetQAVMTestData))]
        public void QAVMTest(QAVTest test)
        {
            Avq.UrlContenResult = File.ReadAllText(AvTimeSeriesMonthlyJsonTestFile);
            object actual = Avq.QAVM(test.Symbol, info: test.Info, tradingMonth: test.DatePart, tradingDate: test.QuoteDate);
            Assert.Equal(test.Expected, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_MONTHLY_ADJUSTED")]
        [Theory]
        [MemberData(nameof(GetQAVMATestData))]
        public void QAVMATest(QAVTest test)
        {
            Avq.UrlContenResult = File.ReadAllText(AvTimeSeriesMonthlyAdjustedJsonTestFile);
            object actual = Avq.QAVMA(test.Symbol, info: test.Info, tradingMonth: test.DatePart, tradingDate: test.QuoteDate);
            Assert.Equal(test.Expected, actual);
        }
        #endregion

        #region Online/Web tests
        public static IEnumerable<object[]> GetQAVQWebTestData()
        {
            yield return new object[] { new QAVTest { } };
            yield return new object[] { new QAVTest { Info = "price" } };
            yield return new object[] { new QAVTest { Info = "volume" } };
            yield return new object[] { new QAVTest { Info = "latest trading day" } };
            yield return new object[] { new QAVTest { Info = "change percent" } };
        }

        public static IEnumerable<object[]> GetQAVWebTestData()
        {
            yield return new object[] { new QAVTest { } };
            yield return new object[] { new QAVTest { Info = "close", DatePart = 0 } };
            yield return new object[] { new QAVTest { Info = "close", DatePart = 101, OutputSize = "full" } };
            yield return new object[] { new QAVTest { Info = "close", DatePart = 101, OutputSize = "full", BestMatch = false } };
            yield return new object[] { new QAVTest { Info = "close", DatePart = -3, BestMatch = false } };
            yield return new object[] { new QAVTest { Info = "close", QuoteDate = DateTime.Today.AddDays(-3), BestMatch = false } };
            yield return new object[] { new QAVTest { Info = "close", QuoteDate = DateTime.Today.LastDayOfWeek(DayOfWeek.Friday), BestMatch = false } };
            yield return new object[] { new QAVTest { Info = "volume" } };
        }

        public static IEnumerable<object[]> GetQAVIDWebTestData()
        {
            yield return new object[] { new QAVTest { } };
            yield return new object[] { new QAVTest { Info = "close", Interval = "1min" } };
            yield return new object[] { new QAVTest { Info = "close", Interval = "5min" } };
            /* Not available with demo API key.
            yield return new object[] { new QAVTest { Info = "close", Interval = "15min" } };
            yield return new object[] { new QAVTest { Info = "close", Interval = "30min" } };
            yield return new object[] { new QAVTest { Info = "close", Interval = "60min" } };*/
            yield return new object[] { new QAVTest { Info = "close", DatePart = 3 } };
            yield return new object[] { new QAVTest { Info = "close", DatePart = 101, OutputSize = "full" } };
            yield return new object[] { new QAVTest { Info = "volume" } };
        }

        public static IEnumerable<object[]> GetQAVDWebNegativeTestData()
        {
            yield return new object[] { new QAVTest { Expected = ExcelError.ExcelErrorGettingData, Info = "n/a" } };
            yield return new object[] { new QAVTest { Expected = ExcelError.ExcelErrorNA, Info = "close", DatePart = 101, BestMatch = false, OutputSize = "compact" } };
            yield return new object[] { new QAVTest { Expected = ExcelError.ExcelErrorNA, Info = "close", QuoteDate = DateTime.Today.AddDays(+1), BestMatch = false } };
            yield return new object[] { new QAVTest { Expected = ExcelError.ExcelErrorNA, Info = "volume", OutputSize = "n/a" } };
        }

        public static IEnumerable<object[]> GetQAVTSWebTestData()
        {
            yield return new object[] { new QAVTest { Interval = "Daily" } };
            yield return new object[] { new QAVTest { Interval = "Daily", Adjusted = true } };
            yield return new object[] { new QAVTest { Interval = "Daily", Adjusted = false, DatePart = 101, OutputSize = "full" } };
            yield return new object[] { new QAVTest { Interval = "Daily", Adjusted = false, QuoteDate = DateTime.Today.AddDays(-3) } };
            yield return new object[] { new QAVTest { Interval = "Daily", Adjusted = false, QuoteDate = DateTime.Today.LastDayOfWeek(DayOfWeek.Sunday), BestMatch = true } };
            yield return new object[] { new QAVTest { Interval = "Weekly" } };
            yield return new object[] { new QAVTest { Interval = "Weekly", Adjusted = true } };
            yield return new object[] { new QAVTest { Interval = "Monthly" } };
            yield return new object[] { new QAVTest { Interval = "Monthly", Adjusted = true } };
            yield return new object[] { new QAVTest { Interval = "1min" } };
            yield return new object[] { new QAVTest { Interval = "5min" } };
            /* Not available with demo API key.
            yield return new object[] { new QAVTest { Interval = "15min" } };
            yield return new object[] { new QAVTest { Interval = "30min" } };
            yield return new object[] { new QAVTest { Interval = "60min" } };*/
        }

        public static IEnumerable<object[]> GetQAVMWebNegativeTestData()
        {
            yield return new object[] { new QAVTest { Expected = ExcelError.ExcelErrorNA, Info = "close", QuoteDate = new DateTime(DateTime.Today.Year - 1, 01, 30), BestMatch = false } };
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "GLOBAL_QUOTE")]
        [Theory]
        [MemberData(nameof(GetQAVQWebTestData))]
        public void QAVQWebTest(QAVTest test)
        {
            Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={test.Symbol}&apikey={test.ApiKey}");
            object actual = Avq.QAVQ(test.Symbol, info: test.Info);

            // If a date specific argument was given, then stock info may not available when trading date is on weekend or a public holiday.
            DateTime tradingDate = DateTime.Today.AddDays(test.DatePart);
            if (tradingDate.DayOfWeek == DayOfWeek.Saturday
                || tradingDate.DayOfWeek == DayOfWeek.Sunday
                || DateSystem.IsPublicHoliday(tradingDate, CountryCode.US))
            {
                // No quotes are available on the weekend or on public holidays.
                Assert.True(ExcelError.ExcelErrorNA.Equals(actual)
                            || actual is DateTime
                            || actual is string
                            || actual is decimal);
            }
            else
            {
                switch (test.Info)
                {
                    case "latest trading day":
                        Assert.IsType<DateTime>(actual);
                        break;
                    case "change percent":
                        Assert.IsType<string>(actual);
                        break;
                    default:
                        Assert.IsType<decimal>(actual);
                        break;
                }
            }
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_DAILY")]
        [Theory]
        [MemberData(nameof(GetQAVWebTestData))]
        public void QAVDWebTest(QAVTest test)
        {
            if (test.OutputSize.Equals("full"))
            {
                Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={test.Symbol}&outputsize={test.OutputSize}&apikey={test.ApiKey}");
            }
            else
            {
                Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={test.Symbol}&apikey={test.ApiKey}");
            }
            object actual = Avq.QAVD(test.Symbol, info: test.Info, tradingDay: test.DatePart, tradingDate: test.QuoteDate, bestMatch: test.BestMatch, outputSize: test.OutputSize);

            QAVAssert(test, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_DAILY")]
        [Theory]
        [MemberData(nameof(GetQAVDWebNegativeTestData))]
        public void QAVDWebNegativeTest(QAVTest test)
        {
            Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={test.Symbol}&apikey={test.ApiKey}");

            if (test.OutputSize.Equals("n/a"))
            {
                Assert.Throws<ArgumentException>(() => Avq.QAVD(test.Symbol, info: test.Info, tradingDay: test.DatePart, tradingDate: test.QuoteDate, bestMatch: test.BestMatch, outputSize: test.OutputSize));
            }
            else
            {
                object actual = Avq.QAVD(test.Symbol, info: test.Info, tradingDay: test.DatePart, tradingDate: test.QuoteDate, bestMatch: test.BestMatch, outputSize: test.OutputSize);
                Assert.Equal(test.Expected, actual);
            }
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_DAILY_ADJUSTED")]
        [Theory]
        [MemberData(nameof(GetQAVWebTestData))]
        public void QAVDAWebTest(QAVTest test)
        {
            if (test.OutputSize.Equals("full"))
            {
                Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&symbol={test.Symbol}&outputsize={test.OutputSize}&apikey={test.ApiKey}");
            }
            else
            {
                Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY_ADJUSTED&symbol={test.Symbol}&apikey={test.ApiKey}");
            }
            object actual = Avq.QAVDA(test.Symbol, info: test.Info, tradingDay: test.DatePart, tradingDate: test.QuoteDate, bestMatch: test.BestMatch, outputSize: test.OutputSize);

            QAVAssert(test, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_INTRADAY")]
        [Theory]
        [MemberData(nameof(GetQAVIDWebTestData))]
        public void QAVIDWebTest(QAVTest test)
        {
            if (test.OutputSize.Equals("full"))
            {
                Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={test.Symbol}&interval={test.Interval}&outputsize={test.OutputSize}&apikey={test.ApiKey}");
            }
            else
            {
                Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={test.Symbol}&interval={test.Interval}&apikey={test.ApiKey}");
            }
            object actual = Avq.QAVID(test.Symbol, info: test.Info, dataPointIndex: test.DatePart, interval: test.Interval, outputSize: test.OutputSize);

            QAVAssert(test, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_WEEKLY")]
        [Theory]
        [MemberData(nameof(GetQAVWebTestData))]
        public void QAVWWebTest(QAVTest test)
        {
            Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_WEEKLY&symbol={test.Symbol}&apikey={test.ApiKey}");
            object actual = Avq.QAVW(test.Symbol, info: test.Info, tradingWeek: test.DatePart, tradingDate: test.QuoteDate, bestMatch: test.BestMatch);

            QAVAssert(test, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_WEEKLY_ADJUSTED")]
        [Theory]
        [MemberData(nameof(GetQAVWebTestData))]
        public void QAVWAWebTest(QAVTest test)
        {
            Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_WEEKLY_ADJUSTED&symbol={test.Symbol}&apikey={test.ApiKey}");
            object actual = Avq.QAVWA(test.Symbol, info: test.Info, tradingWeek: test.DatePart, tradingDate: test.QuoteDate, bestMatch: test.BestMatch);

            QAVAssert(test, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_MONTHLY")]
        [Theory]
        [MemberData(nameof(GetQAVWebTestData))]
        public void QAVMWebTest(QAVTest test)
        {
            Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_MONTHLY&symbol={test.Symbol}&apikey={test.ApiKey}");
            object actual = Avq.QAVM(test.Symbol, info: test.Info, tradingMonth: test.DatePart, tradingDate: test.QuoteDate, bestMatch: test.BestMatch);

            QAVAssert(test, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_MONTHLY")]
        [Theory]
        [MemberData(nameof(GetQAVMWebNegativeTestData))]
        public void QAVMWebNegativeTest(QAVTest test)
        {
            Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_MONTHLY&symbol={test.Symbol}&apikey={test.ApiKey}");
            object actual = Avq.QAVM(test.Symbol, info: test.Info, tradingMonth: test.DatePart, tradingDate: test.QuoteDate, bestMatch: test.BestMatch);
            Assert.Equal(test.Expected, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_MONTHLY_ADJUSTED")]
        [Theory]
        [MemberData(nameof(GetQAVWebTestData))]
        public void QAVMAWebTest(QAVTest test)
        {
            Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_MONTHLY_ADJUSTED&symbol={test.Symbol}&apikey={test.ApiKey}");
            object actual = Avq.QAVMA(test.Symbol, info: test.Info, tradingMonth: test.DatePart, tradingDate: test.QuoteDate, bestMatch: test.BestMatch);

            QAVAssert(test, actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "AVQ")]
        [Trait("API", "TIME_SERIES_API")]
        [Theory]
        [MemberData(nameof(GetQAVTSWebTestData))]
        public void QAVTSWebTest(QAVTest test)
        {
            string api;
            string interval = test.Interval.ToLower();
            if (interval.Equals("daily")
                || interval.Equals("weekly")
                || interval.Equals("monthly"))
            {
                api = $"TIME_SERIES_{interval.ToUpper()}" + (test.Adjusted ? "_ADJUSTED" : "");

                // Remove not valid API parameters (for demo API key).
                if (!interval.Equals("daily")
                    || test.OutputSize.Equals("compact"))
                {
                    test.OutputSize = null;
                }
                test.Interval = null;
            }
            else
            {
                api = "TIME_SERIES_INTRADAY";

                // Remove not valid TIME_SERIES_INTRADAY API parameter (for demo API key).
                test.OutputSize = null;
            }

            Avq.AvStockTimeSeriesOutputSize? outputSize = null;
            if (test.OutputSize != null)
            {
                outputSize = (Avq.AvStockTimeSeriesOutputSize)Enum.Parse(typeof(Avq.AvStockTimeSeriesOutputSize), test.OutputSize);
            }
            string avUri = Avq.AvUrlBuilder(api, apiKey: test.ApiKey, symbol: test.Symbol, outputSize: outputSize, interval: test.Interval);
            Avq.UrlContenResult = FfeWeb.GetHttpResponseContent(avUri);
            object actual = Avq.QAVTS(test.Symbol, info: test.Info, interval: test.Interval, tradingDay: test.DatePart, tradingDate: test.QuoteDate, adjusted: test.Adjusted, outputSize: test.OutputSize, bestMatch: test.BestMatch);

            QAVAssert(test, actual);
        }

        private void QAVAssert(QAVTest test, object actual)
        {
            /* If a date specific argument was given and the BestMatch option was set to false, 
               then stock info may not available when trading date is on weekend or a public holiday.*/
            if ((test.DatePart < 0
                || test.QuoteDate != default(DateTime))
                && !test.BestMatch)
            {
                DateTime tradingDate = DateTime.Today.AddDays(test.DatePart);
                if (tradingDate.DayOfWeek == DayOfWeek.Saturday
                    || tradingDate.DayOfWeek == DayOfWeek.Sunday
                    || DateSystem.IsPublicHoliday(tradingDate, CountryCode.US))
                {
                    // No quotes are available on the weekend or on public holidays.
                    Assert.True(ExcelError.ExcelErrorNA.Equals(actual) || actual is decimal);
                }
            }
            else
            {
                // For current trading dates only test on type is possible.
                switch (test.Info)
                {
                    case "volume":
                        Assert.True(int.TryParse(actual.ToString(), out int value));
                        break;
                    default:
                        Assert.IsType<decimal>(actual);
                        break;
                }
            }
        }
        #endregion
    }

    #region Helper Extensions
    public static class DateTimeExtension
    {
        // Returns the last given day of week based on this DateTime object.
        public static DateTime LastDayOfWeek(this DateTime dateTime, DayOfWeek dayOfWeek)
        {
            while (dateTime.DayOfWeek != dayOfWeek)
            {
                dateTime = dateTime.AddDays(-1);
            }
            return dateTime;
        }
    }
    #endregion
}