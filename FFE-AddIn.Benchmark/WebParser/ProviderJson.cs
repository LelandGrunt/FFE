using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.Collections.Generic;
using System.Globalization;

namespace FFE
{
    public partial class ProviderJson
    {
        [JsonProperty("Provider")]
        public Dictionary<string, Provider> Provider { get; set; }
    }

    public partial class ProviderJson
    {
        public static ProviderJson FromJson(string json) => JsonConvert.DeserializeObject<ProviderJson>(json, ProviderJsonConverter.Settings);
    }

    internal static class ProviderJsonConverter
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