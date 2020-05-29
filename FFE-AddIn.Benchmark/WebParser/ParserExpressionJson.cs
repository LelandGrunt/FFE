using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.Collections.Generic;
using System.Globalization;

namespace FFE
{
    public partial class ParserExpressionJson
    {
        [JsonProperty("Provider")]
        public Provider Provider { get; set; }

        [JsonProperty("WebParsers", ItemConverterType = typeof(StringEnumConverter))]
        public Parser[] WebParsers { get; set; }

        [JsonProperty("WebParserMethods", ItemConverterType = typeof(StringEnumConverter))]
        public WebParserMethod[] WebParserMethods { get; set; }

        [JsonProperty("Expression")]
        public Dictionary<string, Expression> Expression { get; set; }
    }

    public partial class ParserExpressionJson
    {
        public static ParserExpressionJson FromJson(string json) => JsonConvert.DeserializeObject<ParserExpressionJson>(json, ParserExpressionJsonConverter.Settings);
    }

    internal static class ParserExpressionJsonConverter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }
}