using System;
using System.Globalization;

namespace FFE
{
    [AttributeUsage(AttributeTargets.Method, Inherited = false)]
    public class SsqFunctionAttribute : Attribute
    {
        public SsqFunctionAttribute(string url)
        {
            Url = url;
        }

        public SsqFunctionAttribute(string url, string xPath = null, string cssSelector = null, string regExPattern = null)
        {
            Url = url;
            XPath = xPath;
            CssSelector = cssSelector;
            RegExPattern = regExPattern;
        }

        public string Url { get; }
        public string StockIdentifierPlaceholder { get; set; } = "{ISIN_TICKER_WKN}";
        public string XPath { get; set; }
        public string CssSelector { get; set; }
        public string RegExPattern { get; set; } = String.Format(@"(?<quote>\d+{0}\d+)", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
        public int RegExMatchIndex { get; set; } = 0;
        public string RegExGroupName { get; set; } = "quote";
        public string Locale { get; set; } = CultureInfo.CurrentCulture.Name;
        public Parser Parser { get; set; } = Parser.Auto;
    }
}