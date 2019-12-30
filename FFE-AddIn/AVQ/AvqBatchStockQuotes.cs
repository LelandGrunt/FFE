using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace FFE
{
    public partial class AvqBatchStockQuotes
    {
        [JsonProperty("Meta Data")]
        public MetaData MetaData { get; set; }

        [JsonProperty("Stock Quotes")]
        public List<StockQuote> StockQuotes { get; set; }

        [JsonProperty("Information")]
        public string Information { get; set; }

        [JsonProperty("Error Message")]
        public string ErrorMessage { get; set; }

        [JsonProperty("Note")]
        public string Note { get; set; }
    }

    public partial class MetaData
    {
        [JsonProperty("1. Information")]
        public string Information { get; set; }

        [JsonProperty("2. Notes")]
        public string Notes { get; set; }

        [JsonProperty("3. Time Zone")]
        public string TimeZone { get; set; }
    }

    public partial class StockQuote
    {
        [JsonProperty("1. symbol")]
        public string Symbol { get; set; }

        [JsonProperty("2. price")]
        public string Price { get; set; }

        [JsonProperty("3. volume")]
        public string Volume { get; set; }

        [JsonProperty("4. timestamp")]
        public DateTimeOffset Timestamp { get; set; }
    }

    public partial class AvqBatchStockQuotes
    {
        public static AvqBatchStockQuotes FromJson(string json) => JsonConvert.DeserializeObject<AvqBatchStockQuotes>(json, AvqBatchStockQuotesConverter.Settings);
    }

    internal static class AvqBatchStockQuotesConverter
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