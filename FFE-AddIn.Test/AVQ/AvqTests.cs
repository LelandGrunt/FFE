using ExcelDna.Integration;
using Nager.Date;
using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace FFE
{
    public class QAvDTest
    {
        public string Symbol { get; set; }
        public decimal Expected { get; set; }
        public string Info { get; set; } = "close";
        public int Day { get; set; } = 0;
        public DateTime QuoteDate { get; set; } = default(DateTime);
    }

    public class AvqTests
    {
        private readonly string AlphaVantageTestJsonFile = @"AVQ\AlphaVantageTest.json";

        public AvqTests() { }

        public static IEnumerable<object[]> GetAvqTestData()
        {
            yield return new object[] { new QAvDTest { Symbol = "MSFT", Expected = 20180919.1000M, Info = "open", Day = 1 } };
            yield return new object[] { new QAvDTest { Symbol = "MSFT", Expected = 20180919.2000M, Info = "high", Day = 1 } };
            yield return new object[] { new QAvDTest { Symbol = "MSFT", Expected = 20180919.3000M, Info = "low", Day = 1 } };
            yield return new object[] { new QAvDTest { Symbol = "MSFT", Expected = 20180919.4000M, Info = "close", Day = 1 } };
            yield return new object[] { new QAvDTest { Symbol = "MSFT", Expected = 20180919, Info = "volume", Day = 1 } };
            yield return new object[] { new QAvDTest { Symbol = "MSFT", Expected = 20180917.4000M, Info = "close", Day = 3 } };
            yield return new object[] { new QAvDTest { Symbol = "MSFT", Expected = 20180919.4000M, Info = "close", Day = 2, QuoteDate = new DateTime(2018, 9, 19) } };
            yield return new object[] { new QAvDTest { Symbol = "MSFT", Expected = 20180919.4000M, Info = "close", QuoteDate = new DateTime(2018, 9, 19) } };
            yield return new object[] { new QAvDTest { Symbol = "MSFT", Expected = 20180917.4000M, Info = "close", QuoteDate = new DateTime(2018, 9, 17) } };
        }

        [Trait("Query", "AVQ")]
        [Theory]
        [MemberData(nameof(GetAvqTestData))]
        public void QAvDTest(QAvDTest qAvDTest)
        {
            Avq.UrlContenResult = File.ReadAllText(AlphaVantageTestJsonFile);
            object actual = Avq.QAVD(qAvDTest.Symbol, info: qAvDTest.Info, tradingDay: qAvDTest.Day, tradingDate: qAvDTest.QuoteDate);
            Assert.Equal(qAvDTest.Expected, actual);
        }

        public static IEnumerable<object[]> GetAvqTestDateData()
        {
            yield return new object[] { new QAvDTest { } };
            yield return new object[] { new QAvDTest { Info = "close", Day = 0 } };
            yield return new object[] { new QAvDTest { Info = "close", Day = -3 } };
            yield return new object[] { new QAvDTest { Info = "volume" } };
        }

        [Trait("Query", "AVQ")]
        [Theory]
        [MemberData(nameof(GetAvqTestDateData))]
        public void QAvDTestDate(QAvDTest qAvDTest)
        {
            const string symbol = "MSFT";
            Avq.UrlContenResult = FfeWeb.GetHttpResponseContent($"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={symbol}&apikey=demo");
            object actual = Avq.QAVD(symbol, info: qAvDTest.Info, tradingDay: qAvDTest.Day, tradingDate: qAvDTest.QuoteDate);

            DateTime tradingDate = DateTime.Today.AddDays(qAvDTest.Day);
            if (tradingDate.DayOfWeek == DayOfWeek.Saturday
                || tradingDate.DayOfWeek == DayOfWeek.Sunday
                || DateSystem.IsPublicHoliday(tradingDate, CountryCode.US))
            {
                // No quotes are available on the weekend or on public holidays.
                Assert.True(ExcelError.ExcelErrorNA.Equals(actual) || actual is decimal);
            }
            else
            {
                // For current trading dates only test on type is possible.
                if (qAvDTest.Info.Equals("volume"))
                {
                    Assert.True(Int32.TryParse(actual.ToString(), out int value));
                }
                else
                {
                    Assert.IsType<decimal>(actual);
                }
            }
        }

        /* Not testable due to missing Excel instance.
         * TODO: Create excel instance?
        [Theory]
        [InlineData(Avq.AvqInfos.price)]
        [InlineData(Avq.AvqInfos.volume)]
        public void QueryAlphaVantageAsBatch(Avq.AvqInfos info)
        {
            Avq.QueryAlphaVantageAsBatch(info);
            Avq.QueryAlphaVantageAsBatch(info, namedRange: "AVQB");
            Avq.QueryAlphaVantageAsBatch(info, namedRange: "AVQB_*");
        }*/
    }
}