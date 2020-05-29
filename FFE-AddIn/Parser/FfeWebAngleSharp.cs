using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using AngleSharp.Io;
using System;
using System.Threading.Tasks;

namespace FFE
{
    public class FfeWebAngleSharp : IFfeWebParser
    {
        private readonly IDocument Document;

        public bool GenerateRandomUserAgent { get; set; } = true;

        public FfeWebAngleSharp(Uri uri)
        {
            Document = Load(uri.OriginalString);
        }

        public FfeWebAngleSharp(string html)
        {
            Document = Parse(html);
        }

        public dynamic Load(string url)
        {
            return Task.Run(() => LoadAsync(url)).Result;
        }

        public string SelectByCssSelector(string cssSelector)
        {
            IElement element = Document.QuerySelector(cssSelector);

            if (element == null)
            {
                throw new CssSelectorException() { Html = GetHtml(), CssSelector = cssSelector };
            }

            return element.InnerHtml;
        }

        public static string SelectByCssSelector(string cssSelector, string html)
        {
            HtmlParser htmlParser = new HtmlParser();
            IHtmlDocument document = htmlParser.ParseDocument(html);
            IElement element = document.QuerySelector(cssSelector);

            if (element == null)
            {
                throw new CssSelectorException() { Html = document.ToHtml(), CssSelector = cssSelector };
            }

            return element.InnerHtml;
        }

        public string SelectByXPath(string xPath)
        {
            return FfeWebHap.SelectByXPath(xPath, GetHtml());

            // INFO: Add "using AngleSharp.XPath;"
            /*INode node = Document.Body.SelectSingleNode(xPath);
            return node.TextContent;*/
        }

        public string GetHtml() => Document.ToHtml();

        private async Task<IDocument> LoadAsync(string address)
        {
            IConfiguration configuration;
            if (GenerateRandomUserAgent)
            {
                Request request = new Request();
                request.Headers["User-Agent"] = FfeWeb.GetValidUserAgent();
                configuration = Configuration.Default.With(request).WithDefaultLoader();
            }
            else
            {
                configuration = Configuration.Default.WithDefaultLoader();
            }

            IBrowsingContext browsingContext = BrowsingContext.New(configuration);
            IDocument document = await browsingContext.OpenAsync(address);

            if (string.IsNullOrEmpty(document.Body.InnerHtml))
            {
                throw new WebException($"The URL {address} returned no HTML content.");
            }

            return document;
        }

        private IHtmlDocument Parse(string html)
        {
            HtmlParser htmlParser = new HtmlParser();
            IHtmlDocument document = htmlParser.ParseDocument(html);
            return document;
        }
    }
}