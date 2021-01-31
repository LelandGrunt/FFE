using Newtonsoft.Json.Linq;
using System;

namespace FFE
{
    class FfeJsonNewtonsoft : IFfeJsonParser
    {
        //private readonly JObject Document;
        private readonly JArray Document;

        public bool GenerateRandomUserAgent { get; set; } = true;

        public FfeJsonNewtonsoft(Uri uri)
        {
            Document = Load(uri.AbsoluteUri);
        }

        public dynamic Load(string url)
        {
            /*using (Stream stream = FfeWeb.GetHttpResponseContentAsStreamReader(url, GenerateRandomUserAgent))
            using (StreamReader streamReader = new StreamReader(stream))
            using (JsonReader reader = new JsonTextReader(streamReader))
            {
                //return JObject.Load(reader);
                return JArray.Load(reader);
            }*/

            string json = FfeWeb.GetHttpResponseContent(url, GenerateRandomUserAgent);
            //return JObject.Parse(json);
            return JArray.Parse(json);
        }

        public string SelectByJsonPath(string jsonPath)
        {
            return (string)Document.SelectToken(jsonPath);
        }

        public string GetJson()
        {
            return Document.ToString();
        }
    }
}