using BenchmarkDotNet.Attributes;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FFE
{
    public class WebParserMethodExpressionPerformanceParameter
    {
        private readonly string Json;

        public WebParserMethodExpressionPerformanceParameter()
        {
            Json = File.ReadAllText(@".\ParserMethodExpression.json");
        }

        [ParamsSource(nameof(SingleProvider))]
        public Provider Provider;

        public IEnumerable<Provider> SingleProvider()
        {
            yield return ParserExpressionJson.FromJson(Json).Provider;
        }

        [ParamsSource(nameof(WebParsers))]
        public Parser WebParser;

        public IEnumerable<Parser> WebParsers => ParserExpressionJson.FromJson(Json).WebParsers.AsEnumerable();

        [ParamsSource(nameof(Expressions))]
        public Expression Expression;

        public IEnumerable<Expression> Expressions => ParserExpressionJson.FromJson(Json).Expression.Values.AsEnumerable();

        [ParamsSource(nameof(WebParserMethods))]
        public WebParserMethod WebParserMethod { get; set; }

        public IEnumerable<WebParserMethod> WebParserMethods => ParserExpressionJson.FromJson(Json).WebParserMethods.AsEnumerable();
    }
}