using HtmlAgilityPack;
using System;
using static FFE.FfeWeb;

namespace FFE
{
    public class FfeWebHap : IFfeWebParser
    {
        private readonly HtmlDocument Document;

        public bool GenerateRandomUserAgent { get; set; } = true;

        public FfeWebHap(Uri uri)
        {
            Document = Load(uri.AbsoluteUri);
        }

        public FfeWebHap(string html)
        {
            Document = Parse(html);
        }

        public dynamic Load(string url)
        {
            HtmlWeb htmlWeb = new HtmlWeb();

            if (GenerateRandomUserAgent)
            {
                htmlWeb.UserAgent = GetValidUserAgent();
            }

            HtmlDocument htmlDocument = htmlWeb.Load(url);

            return htmlDocument;
        }

        public string SelectByCssSelector(string cssSelector)
        {
            //TODO: Integrate Fizzler / HtmlAgilityPack.CssSelectors.
            return FfeWebAngleSharp.SelectByCssSelector(cssSelector, GetHtml());
        }

        public string SelectByXPath(string xPath)
        {
            HtmlNode htmlNode = Document.DocumentNode.SelectSingleNode(xPath);

            if (htmlNode == null)
            {
                throw new XPathException() { Html = GetHtml(), XPath = xPath };
            }

            return htmlNode.InnerText;
        }

        public static string SelectByXPath(string xPath, string html)
        {
            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);

            HtmlNode htmlNode = htmlDocument.DocumentNode.SelectSingleNode(xPath);

            if (htmlNode == null)
            {
                throw new XPathException() { Html = htmlDocument.Text, XPath = xPath };
            }

            return htmlNode.InnerText;
        }

        public string GetHtml() => Document.Text;

        private HtmlDocument Parse(string html)
        {
            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);
            return htmlDocument;
        }
    }
}