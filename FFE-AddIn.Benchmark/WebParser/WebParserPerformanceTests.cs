using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Mathematics;
using BenchmarkDotNet.Order;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FFE
{
    [BenchmarkCategory("DryRun")]
    [MediumRunJob]
    [StopOnFirstError]
    [Orderer(SummaryOrderPolicy.FastestToSlowest, MethodOrderPolicy.Alphabetical)]
    [MinColumn]
    [MaxColumn]
    [RankColumn(NumeralSystem.Arabic)]
    [RankColumn(NumeralSystem.Stars)]
    [Config(typeof(Config))]
    public class WebParserPerformanceTests
    {
        public WebParserPerformanceTests() { }

        public WebParserPerformanceTests(Parser webParser)
        {
            if (webParser == Parser.HttpClient
                || webParser == Parser.WebClient
                || webParser == Parser.Auto)
            {
                throw new Exception("No parse method available.");
            }

            WebParser = webParser;
        }

        [ParamsSource(nameof(WebParsers))]
        public Parser WebParser { get; set; }

        public IEnumerable<Parser> WebParsers()
        {
            yield return Parser.HAP;
            yield return Parser.AngleSharp;
        }

        public IEnumerable<Provider> Providers() => ProviderJson.FromJson(Resource.Provider).Provider.Select(x => x.Value);

        [Benchmark]
        [BenchmarkCategory("XPath")]
        [ArgumentsSource(nameof(Providers))]
        public decimal SelectByXPath(Provider Provider)
        {
            string html = Resource.ResourceManager.GetString(Provider.Name);
            IFfeWebParser webParser = null;
            switch (WebParser)
            {
                case Parser.Auto:
                    throw new Exception("No parse method available.");
                case Parser.HAP:
                    webParser = new FfeWebHap(html);
                    break;
                case Parser.AngleSharp:
                    webParser = new FfeWebAngleSharp(html);
                    break;
                case Parser.HttpClient:
                    throw new Exception("No parse method available.");
                case Parser.WebClient:
                    throw new Exception("No parse method available.");
                default:
                    break;
            }
            string value = webParser.SelectByXPath(Provider.XPath);
            decimal price = decimal.Parse(value);
            return price;
        }

        [Benchmark]
        [BenchmarkCategory("CssSelector")]
        [ArgumentsSource(nameof(Providers))]
        public decimal SelectByCssSelector(Provider Provider)
        {
            string html = Resource.ResourceManager.GetString(Provider.Name);
            IFfeWebParser webParser = null;
            switch (WebParser)
            {
                case Parser.Auto:
                    throw new Exception("No parse method available.");
                case Parser.HAP:
                    webParser = new FfeWebHap(html);
                    break;
                case Parser.AngleSharp:
                    webParser = new FfeWebAngleSharp(html);
                    break;
                case Parser.HttpClient:
                    throw new Exception("No parse method available.");
                case Parser.WebClient:
                    throw new Exception("No parse method available.");
                default:
                    break;
            }
            string value = webParser.SelectByCssSelector(Provider.CssSelector);
            decimal price = decimal.Parse(value);
            return price;
        }

        [Benchmark]
        [BenchmarkCategory("RegEx")]
        [ArgumentsSource(nameof(Providers))]
        public decimal SelectByRegEx(Provider Provider)
        {
            string html = Resource.ResourceManager.GetString(Provider.Name);
            IFfeWebParser webParser = null;
            switch (WebParser)
            {
                case Parser.Auto:
                    throw new Exception("No parse method available.");
                case Parser.HAP:
                    webParser = new FfeWebHap(html);
                    break;
                case Parser.AngleSharp:
                    webParser = new FfeWebAngleSharp(html);
                    break;
                case Parser.HttpClient:
                    throw new Exception("No parse method available.");
                case Parser.WebClient:
                    throw new Exception("No parse method available.");
                default:
                    break;
            }
            string value = FfeRegEx.RegExByIndexAndGroup(webParser.GetHtml(), Provider.RegExPattern, Provider.RegExMatchIndex.Value, Provider.RegExGroupName);
            decimal price = decimal.Parse(value);
            return price;
        }
    }
}