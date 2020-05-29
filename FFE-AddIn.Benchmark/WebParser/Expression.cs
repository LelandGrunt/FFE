using Newtonsoft.Json;

namespace FFE
{
    public class Expression
    {
        [JsonProperty("Name")]
        public string Name { get; set; }

        [JsonProperty("XPath")]
        public string XPath { get; set; }

        [JsonProperty("CssSelector")]
        public string CssSelector { get; set; }

        [JsonProperty("RegExPattern")]
        public string RegExPattern { get; set; }

        [JsonProperty("RegExGroupName")]
        public string RegExGroupName { get; set; } = "quote";

        [JsonProperty("RegExMatchIndex")]
        public int? RegExMatchIndex { get; set; } = 0;

        public override string ToString()
        {
            return Name;
        }
    }
}