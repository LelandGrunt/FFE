using Xunit;

namespace FFE
{
    public class SsqHttpClientTests
    {
        private const WebParser WEB_PARSER = WebParser.HttpClient;

        [Theory]
        [InlineData("IE00B4L5Y983")]
        [SsqFunction("https://www.justetf.com/de/etf-profile.html?isin={ISIN}",
                     StockIdentifierPlaceholder = "{ISIN}",
                     XPath = "/html/body/div[1]/div[3]/div[7]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/span[2]",
                     Locale = "de-DE",
                     WebParser = WEB_PARSER)]
        public void SSQ_JEQ_Test_Attribute_XPath(string isin)
        {
            object actual = new Ssq(typeof(SsqAngleSharpTests)).TryQ(isin);
            Assert.IsType<decimal>(actual);
        }

        [Theory]
        [InlineData("IE00B4L5Y983")]
        [SsqFunction("https://www.justetf.com/de/etf-profile.html?isin={ISIN}",
                     xPath: "/html/body/div[1]/div[3]/div[7]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/span[2]",
                     StockIdentifierPlaceholder = "{ISIN}",
                     Locale = "de-DE",
                     WebParser = WEB_PARSER)]
        public void SSQ_JEQ_Test_Attribute_XPath_Without_RegEx(string isin)
        {
            object actual = new Ssq(typeof(SsqAngleSharpTests)).TryQ(isin);
            Assert.IsType<decimal>(actual);
        }

        [Theory]
        [InlineData("IE00B4L5Y983")]
        [SsqFunction("https://www.justetf.com/de/etf-profile.html?isin={ISIN}",
                     StockIdentifierPlaceholder = "{ISIN}",
                     RegExPattern = "<span>(?<Currency>\\w+)</span>.*?<span>(?<quote>\\d+,\\d+)</span>",
                     RegExGroupName = "quote",
                     RegExMatchIndex = 0,
                     Locale = "de-DE",
                     WebParser = WEB_PARSER)]
        public void SSQ_JEQ_Test_Attribute_RegEx(string isin)
        {
            object actual = new Ssq(typeof(SsqAngleSharpTests)).TryQ(isin);
            Assert.IsType<decimal>(actual);
        }

        [Theory]
        [InlineData("IE00B4L5Y983")]
        [SsqFunction("https://www.justetf.com/de/etf-profile.html?isin={ISIN}",
                     StockIdentifierPlaceholder = "{ISIN}",
                     XPath = "/html/body/div[1]/div[3]/div[7]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/span[2]",
                     RegExPattern = @"(?<quote>\d+,\d+)",
                     RegExGroupName = "quote",
                     RegExMatchIndex = 0,
                     Locale = "de-DE",
                     WebParser = WEB_PARSER)]
        public void SSQ_JEQ_Test_Attribute_XPath_RegEx(string isin)
        {
            object actual = new Ssq(typeof(SsqAngleSharpTests)).TryQ(isin);
            Assert.IsType<decimal>(actual);
        }
    }
}