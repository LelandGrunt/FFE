using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Mathematics;
using BenchmarkDotNet.Order;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace FFE
{
    [BenchmarkCategory("WebTest")]
    //[SimpleJob(RunStrategy.ColdStart, launchCount: 1, targetCount: 100, id: "WebTestsJob")]
    //[StopOnFirstError]
    [Orderer(SummaryOrderPolicy.FastestToSlowest, MethodOrderPolicy.Alphabetical)]
    [MinColumn]
    [MaxColumn]
    [RankColumn(NumeralSystem.Arabic)]
    [RankColumn(NumeralSystem.Stars)]
    [Config(typeof(WebConfig))]
    public class WebParserPerformanceWebTests
    {
        public WebParserPerformanceWebTests() { }

        public WebParserPerformanceWebTests(Parser webParser)
        {
            WebParser = webParser;
        }

        [GlobalSetup]
        public void GlobalSetup() => Thread.Sleep(5000);

        [IterationSetup]
        public void IterationSetup() => Thread.Sleep(3000);

        //[ParamsAllValues]
        [Params(Parser.Auto, Parser.HAP, Parser.AngleSharp, Parser.HttpClient, Parser.WebClient)]
        public Parser WebParser { get; set; }

        public IEnumerable<Provider> Providers() => ProviderJson.FromJson(Resource.Provider).Provider.Values.AsEnumerable();

        [Benchmark]
        [BenchmarkCategory("XPath")]
        [ArgumentsSource(nameof(Providers))]
        public decimal WebSelectByXPath(Provider Provider)
        {
            IFfeWebParser webParser = null;
            switch (WebParser)
            {
                case Parser.Auto:
                    webParser = FfeWebController.AutoWebParserSelection(Provider.Url, xPath: Provider.XPath, regExPattern: Provider.RegExPattern);
                    break;
                case Parser.HAP:
                    webParser = new FfeWebHap(Provider.Url);
                    break;
                case Parser.AngleSharp:
                    webParser = new FfeWebAngleSharp(Provider.Url);
                    break;
                case Parser.HttpClient:
                    webParser = new FfeWebHttpClient(Provider.Url);
                    break;
                case Parser.WebClient:
                    webParser = new FfeWebClient(Provider.Url);
                    break;
                default:
                    break;
            }
            string value = webParser.SelectByXPath(Provider.XPath);

            decimal quote = decimal.Parse(value);
            return quote;
        }

        [Benchmark]
        [BenchmarkCategory("CssSelector")]
        [ArgumentsSource(nameof(Providers))]
        public decimal WebSelectByCssSelector(Provider Provider)
        {
            IFfeWebParser webParser = null;
            switch (WebParser)
            {
                case Parser.Auto:
                    webParser = FfeWebController.AutoWebParserSelection(Provider.Url, cssSelector: Provider.CssSelector, regExPattern: Provider.RegExPattern);
                    break;
                case Parser.HAP:
                    webParser = new FfeWebHap(Provider.Url);
                    break;
                case Parser.AngleSharp:
                    webParser = new FfeWebAngleSharp(Provider.Url);
                    break;
                case Parser.HttpClient:
                    webParser = new FfeWebHttpClient(Provider.Url);
                    break;
                case Parser.WebClient:
                    webParser = new FfeWebClient(Provider.Url);
                    break;
                default:
                    break;
            }
            string value = webParser.SelectByCssSelector(Provider.CssSelector);

            decimal quote = decimal.Parse(value);
            return quote;
        }

        [Benchmark]
        [BenchmarkCategory("RegEx")]
        [ArgumentsSource(nameof(Providers))]
        public decimal WebSelectByRegEx(Provider Provider)
        {
            IFfeWebParser webParser = null;
            switch (WebParser)
            {
                case Parser.Auto:
                    webParser = FfeWebController.AutoWebParserSelection(Provider.Url);
                    break;
                case Parser.HAP:
                    webParser = new FfeWebHap(Provider.Url);
                    break;
                case Parser.AngleSharp:
                    webParser = new FfeWebAngleSharp(Provider.Url);
                    break;
                case Parser.HttpClient:
                    webParser = new FfeWebHttpClient(Provider.Url);
                    break;
                case Parser.WebClient:
                    webParser = new FfeWebClient(Provider.Url);
                    break;
                default:
                    break;
            }
            string value = FfeRegEx.RegExByIndexAndGroup(webParser.GetHtml(), Provider.RegExPattern, Provider.RegExMatchIndex.Value, Provider.RegExGroupName);

            decimal quote = decimal.Parse(value);
            return quote;
        }
    }
}