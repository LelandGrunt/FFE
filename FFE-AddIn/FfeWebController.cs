﻿using Serilog;
using System;

namespace FFE
{
    public static class FfeWebController
    {
        public static (string value, string rawSource) GetValueFromWeb(string url,
                                                                       string xPath = null,
                                                                       string cssSelector = null,
                                                                       string regExPattern = null, string regExGroup = null, int regExMatchIndex = 0,
                                                                       Parser parser = Parser.Auto)
        {
            if (xPath == null
                && cssSelector == null
                && regExPattern == null)
            {
                throw new ArgumentException("No selection criteria were provided. Provide a least one of the following selection criteria: xPath, cssSelector or regExPattern.");
            }

            string value = null;
            string rawSource = null;

            Uri uri = new Uri(url);

            IFfeWebParser ffeWebParser = null;
            switch (parser)
            {
                case Parser.Auto:
                    ffeWebParser = AutoWebParserSelection(uri, xPath, cssSelector, regExPattern);
                    break;
                case Parser.HAP:
                    ffeWebParser = new FfeWebHap(uri);
                    break;
                case Parser.AngleSharp:
                    ffeWebParser = new FfeWebAngleSharp(uri);
                    break;
                case Parser.HttpClient:
                    ffeWebParser = new FfeWebHttpClient(uri);
                    break;
                case Parser.WebClient:
                    ffeWebParser = new FfeWebClient(uri);
                    break;
                default:
                    ffeWebParser = new FfeWebHap(uri);
                    break;
            }

            rawSource = ffeWebParser.GetHtml();

            if (Log.IsEnabled(Serilog.Events.LogEventLevel.Debug))
            {
                rawSource.WriteToFile($"PageSource_{uri.Host}.html", parser.ToString());
            }

            // Input for RegEx (if set).
            string input = null;
            if (!String.IsNullOrEmpty(xPath))
            {
                value = ffeWebParser.SelectByXPath(xPath);
                input = value;
            }
            else if (!String.IsNullOrEmpty(cssSelector))
            {
                value = ffeWebParser.SelectByCssSelector(cssSelector);
                input = value;
            }
            else // Select by RegEx (HTML source code = RegEx input).
            {
                //HACK: AngelSharp does not provide full HTML source code.
                if (parser == Parser.AngleSharp)
                {
                    Log.Warning("Regular Expression with AngleSharp WebParser does not work well. HttpClient is used instead.");
                    rawSource = FfeWeb.GetHttpResponseContent(uri);
                }
                input = rawSource;
            }

            if (!String.IsNullOrEmpty(regExPattern))
            {
                value = FfeRegEx.RegExByIndexAndGroup(input, regExPattern, regExMatchIndex, regExGroup);

                if (String.IsNullOrEmpty(value))
                {
                    if (Log.IsEnabled(Serilog.Events.LogEventLevel.Debug))
                    {
                        input.WriteToFile("RegExInput.html", "RegEx");
                    }
                    throw new RegExException() { Input = input, Pattern = regExPattern };
                }
            }

            return (value, rawSource);
        }

        public static IFfeWebParser AutoWebParserSelection(Uri uri,
                                                           string xPath = null,
                                                           string cssSelector = null,
                                                           string regExPattern = null)
        {
            if (!String.IsNullOrEmpty(xPath))
            {
                //TODO: Based on benchmark results, the AngleSharp parser is faster than the HtmlAgilityPack one.
                return new FfeWebHap(uri);
            }
            else if (!String.IsNullOrEmpty(cssSelector))
            {
                return new FfeWebAngleSharp(uri);
            }
            else if (!String.IsNullOrEmpty(regExPattern))
            {
                //TODO: Based on benchmark results, the AngleSharp parser is faster than the HttpClient one.
                return new FfeWebHttpClient(uri);
            }
            else
            {
                return new FfeWebHap(uri);
            }
        }
    }
}