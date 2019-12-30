using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Serilog;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace FFE
{
    public partial class SsqJson
    {
        [JsonProperty("Version")]
        public string Version { get; set; }

        [JsonProperty("Version Date")]
        public DateTime VersionDate { get; set; }

        [JsonProperty("UDF")]
        public Dictionary<string, UserDefinedFunction> UDF { get; set; }
    }

    public class UserDefinedFunction
    {
        [JsonProperty("QueryInformation")]
        public QueryInformation QueryInformation { get; set; }

        [JsonProperty("QueryParameter")]
        public QueryParameter QueryParameter { get; set; }
    }

    public class QueryInformation
    {
        [JsonProperty("Name")]
        public string Name { get; set; }

        [JsonProperty("Description")]
        public string Description { get; set; }

        [JsonProperty("Author")]
        public string Author { get; set; }

        [JsonProperty("Author Email")]
        public string AuthorEmail { get; set; }

        [JsonProperty("Version")]
        public string Version { get; set; }

        [JsonProperty("Version Date")]
        public string VersionDate { get; set; }

        [JsonProperty("Help Topic", NullValueHandling = NullValueHandling.Ignore)]
        public string HelpTopic { get; set; }

        [JsonProperty("Help Link", NullValueHandling = NullValueHandling.Ignore)]
        public string HelpLink { get; set; }

        [JsonProperty("Provider")]
        public string Provider { get; set; }

        [JsonProperty("Enabled")]
        public bool Enabled { get; set; }

        [JsonProperty("ExcelArgNameStockIdentifier", NullValueHandling = NullValueHandling.Ignore)]
        public string ExcelArgNameStockIdentifier { get; set; }

        [JsonProperty("ExcelArgDescStockIdentifier", NullValueHandling = NullValueHandling.Ignore)]
        public string ExcelArgDescStockIdentifier { get; set; }
    }

    public class QueryParameter
    {
        [JsonProperty("Url")]
        public string Url { get; set; }

        [JsonProperty("StockIdentifierPlaceholder", NullValueHandling = NullValueHandling.Ignore)]
        public string StockIdentifierPlaceholder { get; set; }

        [JsonProperty("XPath", NullValueHandling = NullValueHandling.Ignore)]
        public string XPath { get; set; }

        [JsonProperty("CssSelector", NullValueHandling = NullValueHandling.Ignore)]
        public string CssSelector { get; set; }

        [JsonProperty("RegExPattern", NullValueHandling = NullValueHandling.Ignore)]
        public string RegExPattern { get; set; }

        [JsonProperty("RegExGroupName", NullValueHandling = NullValueHandling.Ignore)]
        public string RegExGroupName { get; set; }

        [JsonProperty("RegExMatchIndex", NullValueHandling = NullValueHandling.Ignore)]
        public int? RegExMatchIndex { get; set; }

        [JsonProperty("Locale", NullValueHandling = NullValueHandling.Ignore)]
        public string Locale { get; set; }

        [JsonProperty("Parser", NullValueHandling = NullValueHandling.Ignore)]
        [JsonConverter(typeof(StringEnumConverter))]
        public Parser Parser { get; set; }
    }

    public partial class SsqJson
    {
        public static SsqJson FromJson(string json)
        {
            try
            {
                return JsonConvert.DeserializeObject<SsqJson>(json, SsqJsonConverter.Settings);
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);

                throw;
            }
        }
    }

    internal static class SsqJsonConverter
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