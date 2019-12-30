namespace FFE
{
    public class SsqTestMethod
    {
        public readonly Parser Parser;

        public SsqTestMethod(Parser parser)
        {
            Parser = parser;
        }

        public object SsqConstructorXPathMinTest(Provider provider)
        {
            string url = provider.Url.OriginalString;
            string xPath = provider.XPath;
            Ssq ssq = new Ssq(url, xPath, parser: Parser);
            return ssq.Q(provider.IsinTickerWkn);
        }

        public object SsqConstructorXPathTest(Provider provider)
        {
            string udfName = provider.UdfName;
            string url = provider.Url.OriginalString;
            string stockIdentifierPlaceholder = provider.StockIdentifierPlaceholder;
            string xPath = provider.XPath;
            string locale = provider.Locale;
            Ssq ssq = new Ssq(url, xPath, stockIdentifierPlaceholder: stockIdentifierPlaceholder, locale: locale, parser: Parser, udfName: udfName);
            return ssq.Q(provider.IsinTickerWkn);
        }

        public object SsqConstructorRegExMinTest(Provider provider)
        {
            string url = provider.Url.OriginalString;
            string regExPattern = provider.RegExPattern;
            Ssq ssq = new Ssq(url, regExPattern: regExPattern, parser: Parser);
            return ssq.Q(provider.IsinTickerWkn);
        }

        public object SsqConstructorRegExTest(Provider provider)
        {
            string udfName = provider.UdfName;
            string url = provider.Url.OriginalString;
            string stockIdentifierPlaceholder = provider.StockIdentifierPlaceholder;
            string regExPattern = provider.RegExPattern;
            string regExGroupName = provider.RegExGroupName;
            int regExMatchIndex = 0;
            string locale = provider.Locale;
            Ssq ssq = new Ssq(url, stockIdentifierPlaceholder: stockIdentifierPlaceholder, locale: locale, udfName: udfName,
                              regExPattern: regExPattern, regExGroupName: regExGroupName, regExMatchIndex: regExMatchIndex,
                              parser: Parser);
            return ssq.Q(provider.IsinTickerWkn);
        }

        public object SsqConstructorXPathRegExTest(Provider provider)
        {
            string udfName = provider.UdfName;
            string url = provider.Url.OriginalString;
            string stockIdentifierPlaceholder = provider.StockIdentifierPlaceholder;
            string xPath = provider.XPath;
            string regExPattern = @"(?<quote>\d+,\d+)";
            string regExGroupName = "quote";
            int regExMatchIndex = 0;
            string locale = provider.Locale;
            Ssq ssq = new Ssq(url, xPath, stockIdentifierPlaceholder: stockIdentifierPlaceholder, locale: locale, udfName: udfName,
                              regExPattern: regExPattern, regExGroupName: regExGroupName, regExMatchIndex: regExMatchIndex,
                              parser: Parser);
            return ssq.Q(provider.IsinTickerWkn);
        }

        public object SsqPropertyXPathTest(Provider provider)
        {
            Ssq ssq = new Ssq(provider.UdfName)
            {
                Url = provider.Url.OriginalString,
                StockIdentifierPlaceholder = provider.StockIdentifierPlaceholder,
                XPath = provider.XPath,
                Culture = System.Globalization.CultureInfo.CurrentCulture,
                Parser = Parser
            };
            return ssq.Q(provider.IsinTickerWkn);
        }

        public object SsqPropertyRegExTest(Provider provider)
        {

            Ssq ssq = new Ssq(provider.UdfName)
            {
                Url = provider.Url.OriginalString,
                StockIdentifierPlaceholder = provider.StockIdentifierPlaceholder,
                RegExPattern = provider.RegExPattern,
                RegExGroupName = provider.RegExGroupName,
                RegExMatchIndex = provider.RegExMatchIndex,
                Culture = System.Globalization.CultureInfo.CurrentCulture,
                Parser = Parser
            };
            return ssq.Q(provider.IsinTickerWkn);
        }

        public object SsqPropertyXPathRegExTest(Provider provider)
        {
            Ssq ssq = new Ssq(provider.UdfName)
            {
                Url = provider.Url.OriginalString,
                StockIdentifierPlaceholder = provider.StockIdentifierPlaceholder,
                XPath = provider.XPath,
                RegExPattern = @"(?<quote>\d+,\d+)",
                RegExGroupName = "quote",
                RegExMatchIndex = 0,
                Culture = System.Globalization.CultureInfo.CurrentCulture,
                Parser = Parser
            };
            return ssq.Q(provider.IsinTickerWkn);
        }
    }
}