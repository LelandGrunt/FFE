using System;
using System.Collections.Generic;

namespace FFE
{
    public class Provider
    {
        public string UdfName { get; set; }
        public string Name { get; set; }
        public Uri Url { get; set; }
        public string StockIdentifierPlaceholder { get; set; }
        public string IsinTickerWkn { get; set; }
        public string XPath { get; set; }
        public string CssSelector { get; set; }
        public string RegExPattern { get; set; }
        public string RegExGroupName { get; set; } = "quote";
        public int RegExMatchIndex { get; set; } = 0;
        public string Locale { get; set; } = System.Globalization.CultureInfo.CurrentCulture.Name;

        public static IEnumerable<Provider> ProviderTestData()
        {
            yield return new Provider
            {
                UdfName = "JEQ",
                Name = "justETF",
                Url = new Uri("https://www.justetf.com/de/etf-profile.html?isin={ISIN}"),
                StockIdentifierPlaceholder = "{ISIN}",
                IsinTickerWkn = "IE00B4L5Y983",
                XPath = "/html/body/div[1]/div[3]/div[7]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/span[2]",
                CssSelector = ".tab-container > div:nth-child(1) > div:nth-child(3) > div:nth-child(2) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > span:nth-child(2)",
                RegExPattern = "<div class=\"col-xs-6\">.*?<div class=\"val\">.*?<span>(?<Currency>.*)</span>.*?<span>(?<quote>\\d*,\\d*)</span>.*?</div>",
                RegExGroupName = "quote",
                RegExMatchIndex = 0,
                Locale = "de-DE"
            };
            yield return new Provider
            {
                UdfName = "QCB",
                Name = "Consorsbank",
                Url = new Uri("https://www.consorsbank.de/Wertpapierhandel/Fonds/Kurs-Snapshot/snapshotoverview/{WKN}"),
                StockIdentifierPlaceholder = "{WKN}",
                IsinTickerWkn = "US5949181045",
                XPath = "//strong[starts-with(@class,'price price-')]",
                CssSelector = ".price",
                RegExPattern = @"<strong class=.price price-.*.>.*?(?<price>\d*,\d*)</strong>",
                RegExGroupName = "price",
                RegExMatchIndex = 0,
                Locale = "de-DE"
            };
            yield return new Provider
            {
                UdfName = "YFQ",
                Name = "Yahoo_Finance",
                Url = new Uri("https://finance.yahoo.com/quote/{TICKER}"),
                StockIdentifierPlaceholder = "{TICKER}",
                IsinTickerWkn = "MSFT",
                XPath = "/html/body/div[1]/div/div/div[1]/div/div[2]/div/div/div[4]/div/div/div/div[3]/div[1]/div/span[1]",
                CssSelector = @".Fz\(36px\)",
                RegExPattern = "<span class=\"Trsdu\\(0.3s\\) Fw\\(b\\) Fz\\(36px\\) Mb\\(-4px\\) D\\(ib\\)\" data-reactid=\"\\d*\">(?<quote>\\d*.\\d*)</span>",
                RegExGroupName = "quote",
                RegExMatchIndex = 0,
                Locale = "en-US"
            };
        }

        public static IEnumerable<string> WknTestData()
        {
            return new string[] { "870747", "865985", "A1JWVX" };
        }

        public static IEnumerable<string> IsinTestData()
        {
            return new string[] { "IE00B4L5Y983", "IE00B14X4T88", "IE00BKM4GZ66" };
        }

        public static IEnumerable<string> TickerTestData()
        {
            return new string[] { "MSFT", "AAPL", "FB" };
        }

        public static IEnumerable<Provider> ProviderMinTestData()
        {
            yield return new Provider
            {
                UdfName = "JEQ",
                Name = "justETF",
                Url = new Uri("https://www.justetf.com/de/etf-profile.html?isin={ISIN_TICKER_WKN}"),
                IsinTickerWkn = "IE00B4L5Y983",
                XPath = "/html/body/div[1]/div[3]/div[7]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/span[2]",
                CssSelector = ".tab-container > div:nth-child(1) > div:nth-child(3) > div:nth-child(2) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > span:nth-child(2)",
                RegExPattern = "<div class=\"col-xs-6\">.*?<span>(?<quote>\\d*,\\d*)</span>.*?</div>"
            };
        }
    }
}