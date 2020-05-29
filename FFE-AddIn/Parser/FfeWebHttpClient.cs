using System;

namespace FFE
{
    public class FfeWebHttpClient : FfeWeb, IFfeWebParser
    {
        private readonly string Document;

        public new bool GenerateRandomUserAgent { get; set; } = true;

        public FfeWebHttpClient() : base() { }

        public FfeWebHttpClient(Uri uri)
        {
            Document = Load(uri.AbsoluteUri);
        }

        public dynamic Load(string url)
        {
            return GetHttpResponseContent(url, GenerateRandomUserAgent);
        }

        public string SelectByXPath(string xPath)
        {
            return FfeWebHap.SelectByXPath(xPath, GetHtml());
        }

        public string SelectByCssSelector(string cssSelector)
        {
            return FfeWebAngleSharp.SelectByCssSelector(cssSelector, GetHtml());
        }

        public string GetHtml() => Document;
    }
}