using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using BenchmarkDotNet.Mathematics;
using BenchmarkDotNet.Order;
using System.Threading;

namespace FFE
{
    [BenchmarkCategory("WebParserMethodExpressionPerformanceWebTests")]
    [SimpleJob(RunStrategy.ColdStart, launchCount: 1, targetCount: 100, id: "WebParserMethodExpressionPerformanceWebTestsJob")]
    //[SimpleJob(BenchmarkDotNet.Jobs.RuntimeMoniker.Net462, launchCount: 1, targetCount: 100, id: "WebParserMethodExpressionPerformanceWebTestsJob")]
    //[DryJob(BenchmarkDotNet.Jobs.RuntimeMoniker.Net462)]
    //[StopOnFirstError]
    [Orderer(SummaryOrderPolicy.FastestToSlowest, MethodOrderPolicy.Alphabetical)]
    [MinColumn]
    [MaxColumn]
    [RankColumn(NumeralSystem.Arabic)]
    [RankColumn(NumeralSystem.Stars)]
    public class WebParserMethodExpressionPerformanceWebTests : WebParserMethodExpressionPerformanceParameter
    {
        #region Setup
        [GlobalSetup]
        public void GlobalSetup() => Thread.Sleep(5000);

        [IterationSetup]
        public void IterationSetup() => Thread.Sleep(3000);
        #endregion

        [Benchmark]
        public decimal? WebSelectByProviderParserMethodExpression()
        {
            decimal? quote = null;
            switch (WebParserMethod)
            {
                case WebParserMethod.XPath:
                    Provider.XPath = Expression.XPath;
                    quote = new WebParserPerformanceWebTests(WebParser).WebSelectByXPath(Provider);
                    break;
                case WebParserMethod.CssSelector:
                    Provider.CssSelector = Expression.CssSelector;
                    quote = new WebParserPerformanceWebTests(WebParser).WebSelectByCssSelector(Provider);
                    break;
                case WebParserMethod.RegEx:
                    Provider.RegExPattern = Expression.RegExPattern;
                    Provider.RegExGroupName = Expression.RegExGroupName ?? "quote";
                    Provider.RegExMatchIndex = Expression.RegExMatchIndex ?? 0;
                    quote = new WebParserPerformanceWebTests(WebParser).WebSelectByRegEx(Provider);
                    break;
                default:
                    break;
            }
            return quote;
        }
    }
}