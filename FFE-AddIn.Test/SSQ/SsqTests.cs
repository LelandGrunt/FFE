using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace FFE
{
    public class SsqTests : IDisposable
    {
        // INFO: Newtonsoft parser excluded because of incomplete implementation!
        private static readonly IEnumerable<Parser> Parsers = Enum.GetValues(typeof(Parser)).Cast<Parser>().Where(p => !p.Equals(Parser.Newtonsoft));

        #region Test Data
        public static IEnumerable<object[]> ParserTestData()
        {
            foreach (var parser in Parsers)
            {
                yield return new object[] { parser };
            }
        }

        public static IEnumerable<object[]> SsqTestData()
        {
            foreach (var parser in Parsers)
            {
                foreach (Provider provider in Provider.ProviderTestData())
                {
                    //HACK: Overwrite provider values for special cases.
                    if (parser.Equals(Parser.AngleSharp)
                        && provider.Name.Equals("Yahoo_Finance"))
                    {
                        provider.XPath = "/html/body/div[1]/div/div/div/div/div/div[2]/div/div/div/div[3]/div/div/div/div[3]/div/span";
                    }
                    if (parser.Equals(Parser.HttpClient)
                        && provider.Name.Equals("Consorsbank"))
                    {
                        provider.XPath = "//strong[starts-with(@class,'price price-')]";
                    }

                    switch (provider.StockIdentifierPlaceholder)
                    {
                        case "{WKN}":
                            foreach (string wkn in Provider.WknTestData())
                            {
                                provider.IsinTickerWkn = wkn;
                                yield return new object[] { parser, provider };
                            }
                            break;
                        case "{ISIN}":
                            foreach (string isin in Provider.IsinTestData())
                            {
                                provider.IsinTickerWkn = isin;
                                yield return new object[] { parser, provider };
                            }
                            break;
                        case "{TICKER}":
                            foreach (string ticker in Provider.TickerTestData())
                            {
                                provider.IsinTickerWkn = ticker;
                                yield return new object[] { parser, provider };
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        public static IEnumerable<object[]> SsqMethodTestData()
        {
            yield return new object[] { Parser.Auto, Provider.ProviderTestData().First() };
        }

        public static IEnumerable<object[]> SsqMethodMinTestData()
        {
            yield return new object[] { Parser.Auto, Provider.ProviderMinTestData().First() };
        }
        #endregion

        #region SsqTests
        [Trait("Query", "SSQ")]
        [Trait("Method", "XPath")]
        [Theory]
        [MemberData(nameof(SsqTestData))]
        public void SsqXPathTest(Parser parser, Provider provider)
        {
            Ssq ssq = new Ssq(provider.Url.OriginalString,
                              xPath: provider.XPath,
                              parser: parser,
                              stockIdentifierPlaceholder: provider.StockIdentifierPlaceholder,
                              locale: provider.Locale,
                              udfName: provider.UdfName);
            object actual = ssq.Q(provider.IsinTickerWkn);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "CssSelector")]
        [Theory]
        [MemberData(nameof(SsqTestData))]
        public void SsqCssSelectorTest(Parser parser, Provider provider)
        {
            Ssq ssq = new Ssq(provider.Url.OriginalString,
                              cssSelector: provider.CssSelector,
                              parser: parser,
                              stockIdentifierPlaceholder: provider.StockIdentifierPlaceholder,
                              locale: provider.Locale,
                              udfName: provider.UdfName);
            object actual = ssq.Q(provider.IsinTickerWkn);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "RegEx")]
        [Theory]
        [MemberData(nameof(SsqTestData))]
        public void SsqRegExTest(Parser parser, Provider provider)
        {
            Ssq ssq = new Ssq(provider.Url.OriginalString,
                              regExPattern: provider.RegExPattern,
                              regExGroupName: provider.RegExGroupName,
                              regExMatchIndex: provider.RegExMatchIndex,
                              parser: parser,
                              stockIdentifierPlaceholder: provider.StockIdentifierPlaceholder,
                              locale: provider.Locale,
                              udfName: provider.UdfName);
            object actual = ssq.Q(provider.IsinTickerWkn);
            Assert.IsType<decimal>(actual);
        }
        #endregion

        #region SsqMethodTests
        [Trait("Query", "SSQ")]
        [Trait("Method", "XPath")]
        [Trait("SsqUse", "Constructor")]
        [Theory]
        [MemberData(nameof(SsqMethodMinTestData))]
        public void SsqConstructorXPathMinTest(Parser parser, Provider provider)
        {
            object actual = new SsqTestMethod(parser).SsqConstructorXPathMinTest(provider);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "XPath")]
        [Trait("SsqUse", "Constructor")]
        [Theory]
        [MemberData(nameof(SsqMethodTestData))]
        public void SsqConstructorXPathTest(Parser parser, Provider provider)
        {
            object actual = new SsqTestMethod(parser).SsqConstructorXPathTest(provider);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "RegEx")]
        [Trait("SsqUse", "Constructor")]
        [Theory]
        [MemberData(nameof(SsqMethodMinTestData))]
        public void SsqConstructorRegExMinTest(Parser parser, Provider provider)
        {
            object actual = new SsqTestMethod(parser).SsqConstructorRegExMinTest(provider);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "RegEx")]
        [Trait("SsqUse", "Constructor")]
        [Theory]
        [MemberData(nameof(SsqMethodTestData))]
        public void SsqConstructorRegExTest(Parser parser, Provider provider)
        {
            object actual = new SsqTestMethod(parser).SsqConstructorRegExTest(provider);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "XPath")]
        [Trait("Method", "RegEx")]
        [Trait("SsqUse", "Constructor")]
        [Theory]
        [MemberData(nameof(SsqMethodTestData))]
        public void SsqConstructorXPathRegExTest(Parser parser, Provider provider)
        {
            object actual = new SsqTestMethod(parser).SsqConstructorXPathRegExTest(provider);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "XPath")]
        [Trait("SsqUse", "Property")]
        [Theory]
        [MemberData(nameof(SsqMethodTestData))]
        public void SsqPropertyXPathTest(Parser parser, Provider provider)
        {
            object actual = new SsqTestMethod(parser).SsqPropertyXPathTest(provider);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "RegEx")]
        [Trait("SsqUse", "Property")]
        [Theory]
        [MemberData(nameof(SsqMethodTestData))]
        public void SsqPropertyRegExTest(Parser parser, Provider provider)
        {
            object actual = new SsqTestMethod(parser).SsqPropertyRegExTest(provider);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "XPath")]
        [Trait("Method", "RegEx")]
        [Trait("SsqUse", "Property")]
        [Theory]
        [MemberData(nameof(SsqMethodTestData))]
        public void SsqPropertyXPathRegExTest(Parser parser, Provider provider)
        {
            object actual = new SsqTestMethod(parser).SsqPropertyXPathRegExTest(provider);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "XPath")]
        [Trait("SsqUse", "Attribute")]
        [Theory]
        [InlineData("IE00B4L5Y983")]
        [SsqFunction("https://www.justetf.com/de/etf-profile.html?isin={ISIN}",
                     StockIdentifierPlaceholder = "{ISIN}",
                     XPath = "/html/body/div[1]/div[3]/div[7]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/span[2]",
                     Locale = "de-DE",
                     Parser = Parser.Auto)]
        public void SsqAttributeXPathTest(string isin)
        {
            object actual = new Ssq(typeof(SsqTests)).Q(isin);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "XPath")]
        [Trait("SsqUse", "Attribute")]
        [Theory]
        [InlineData("IE00B4L5Y983")]
        [SsqFunction("https://www.justetf.com/de/etf-profile.html?isin={ISIN}",
                     xPath: "/html/body/div[1]/div[3]/div[7]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/span[2]",
                     StockIdentifierPlaceholder = "{ISIN}",
                     Locale = "de-DE",
                     Parser = Parser.Auto)]
        public void SsqAttributeXPathWithoutRegExTest(string isin)
        {
            object actual = new Ssq(typeof(SsqTests)).Q(isin);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "RegEx")]
        [Trait("SsqUse", "Attribute")]
        [Theory]
        [InlineData("IE00B4L5Y983")]
        [SsqFunction("https://www.justetf.com/de/etf-profile.html?isin={ISIN}",
                     StockIdentifierPlaceholder = "{ISIN}",
                     RegExPattern = "<span>(?<Currency>\\w+)</span>.*?<span>(?<quote>\\d+,\\d+)</span>",
                     RegExGroupName = "quote",
                     RegExMatchIndex = 0,
                     Locale = "de-DE",
                     Parser = Parser.Auto)]
        public void SsqAttributeRegExTest(string isin)
        {
            object actual = new Ssq(typeof(SsqTests)).Q(isin);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Trait("Method", "XPath")]
        [Trait("Method", "RegEx")]
        [Trait("SsqUse", "Attribute")]
        [Theory]
        [InlineData("IE00B4L5Y983")]
        [SsqFunction("https://www.justetf.com/de/etf-profile.html?isin={ISIN}",
                     StockIdentifierPlaceholder = "{ISIN}",
                     XPath = "/html/body/div[1]/div[3]/div[7]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/span[2]",
                     RegExPattern = @"(?<quote>\d+,\d+)",
                     RegExGroupName = "quote",
                     RegExMatchIndex = 0,
                     Locale = "de-DE",
                     Parser = Parser.Auto)]
        public void SsqAttributeXPathRegExTest(string isin)
        {
            object actual = new Ssq(typeof(SsqTests)).Q(isin);
            Assert.IsType<decimal>(actual);
        }
        #endregion

        [Trait("Query", "SSQ")]
        [Theory]
        [InlineData("QJE", "IE00B4L5Y983")]
        //[InlineData("QBB", "MSFT:US")]
        [InlineData("QYF", "MSFT")]
        public void SsqExcelFunctionTest(string ssqFunction, string wkn_isin_ticker)
        {
            IEnumerable<SsqExcelFunction> ssqExcelFunctions = SsqLoader.GetExcelFunctions()
                                                                       .Where(f => ((ExcelFunctionAttribute)f.ExcelFunctionAttribute).Name.Equals(ssqFunction));
            var ssqExcelFuntion = ssqExcelFunctions.Single();
            object actual = ssqExcelFuntion.Delegate.DynamicInvoke(wkn_isin_ticker);
            Assert.IsType<decimal>(actual);
        }

        [Trait("Query", "SSQ")]
        [Fact]
        public void SsqLoaderTest()
        {
            Assert.NotEmpty(SsqLoader.GetExcelFunctions());
        }

        [Trait("Query", "SSQ")]
        [Fact]
        public void SsqJsonVersionDateTest()
        {
            Assert.IsType<DateTime>(SsqLoader.SsqJson.VersionDate);
        }

        public void Dispose()
        {
            // INFO: Avoids bad HTTP requests?
            System.Threading.Thread.Sleep(3000);
        }
    }
}