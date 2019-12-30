using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Mathematics;
using BenchmarkDotNet.Order;

namespace FFE
{
    [BenchmarkCategory("WebParserMethodExpressionPerformanceTests")]
    //[StopOnFirstError]
    [Orderer(SummaryOrderPolicy.FastestToSlowest, MethodOrderPolicy.Alphabetical)]
    [MinColumn]
    [MaxColumn]
    [RankColumn(NumeralSystem.Arabic)]
    [RankColumn(NumeralSystem.Stars)]
    public class WebParserMethodExpressionPerformanceTests : WebParserMethodExpressionPerformanceParameter
    {
        [Benchmark]
        public decimal? SelectByProviderParserMethodExpression()
        {
            decimal? quote = null;
            switch (WebParserMethod)
            {
                case WebParserMethod.XPath:
                    Provider.XPath = Expression.XPath;
                    quote = new WebParserPerformanceTests(WebParser).SelectByXPath(Provider);
                    break;
                case WebParserMethod.CssSelector:
                    Provider.CssSelector = Expression.CssSelector;
                    quote = new WebParserPerformanceTests(WebParser).SelectByCssSelector(Provider);
                    break;
                case WebParserMethod.RegEx:
                    Provider.RegExPattern = Expression.RegExPattern;
                    Provider.RegExGroupName = Expression.RegExGroupName ?? "quote";
                    Provider.RegExMatchIndex = Expression.RegExMatchIndex ?? 0;
                    quote = new WebParserPerformanceTests(WebParser).SelectByRegEx(Provider);
                    break;
                default:
                    break;
            }
            return quote;
        }
    }
}