using System;
using System.Net;

namespace FFE
{
    public class FfeWebClient : IFfeWebParser
    {
        private readonly string Document;

        public FfeWebClient(Uri uri)
        {
            // https://docs.microsoft.com/en-us/dotnet/framework/whats-new/#net47
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                                                 | SecurityProtocolType.Tls11
                                                 | SecurityProtocolType.Tls
                                                 | SecurityProtocolType.Ssl3;

            Document = Load(uri.AbsoluteUri);
        }

        public dynamic Load(string url)
        {
            using (GZipWebClient client = new GZipWebClient())
            {
                return client.DownloadString(url);
            }
        }

        public string SelectByXPath(string xPath)
        {
            return FfeWebHap.SelectByXPath(xPath, GetHtml());
        }

        public string SelectByCssSelector(string cssSelector)
        {
            return FfeWebAngleSharp.SelectByCssSelector(cssSelector, GetHtml());
        }

        public string GetHtml()
        {
            string html = new FfeWebAngleSharp(Document).GetHtml();
            return html;
        }

        //https://stackoverflow.com/questions/8936089/htmlagilitypack-gzip-encryption-exception
        private class GZipWebClient : WebClient
        {
            protected override WebRequest GetWebRequest(Uri address)
            {
                HttpWebRequest request = (HttpWebRequest)base.GetWebRequest(address);
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.None;

                // Create and set random user agent.
                request.UserAgent = new Bogus.DataSets.Internet().UserAgent();

                // Try to avoid robot detection (BBQ).
                /*request.Accept = "text/html, application/xhtml+xml, image/jxr";
                request.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip, deflate, none");
                request.Headers.Add(HttpRequestHeader.AcceptCharset, "iso-8859-15");
                request.Headers.Add(HttpRequestHeader.AcceptLanguage, "en-us");
                request.KeepAlive = true;*/

                return request;
            }
        }
    }
}