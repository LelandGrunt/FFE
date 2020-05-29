using ExcelDna.Integration;
using Serilog;
using System;
using System.Globalization;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace FFE
{
    public class Ssq : FfeBase
    {
        public readonly string UDF_NAME;

        #region Constructors
        public Ssq([System.Runtime.CompilerServices.CallerMemberName] string udfName = "") : base(udfName, typeof(Ssq))
        {
            UDF_NAME = udfName;
        }

        public Ssq(Type type, [System.Runtime.CompilerServices.CallerMemberName] string udfName = "") : base(udfName, typeof(Ssq))
        {
            SsqFunctionAttribute ssqAttribute = (SsqFunctionAttribute)Attribute.GetCustomAttribute(type.GetMethod(udfName), typeof(SsqFunctionAttribute));

            if (ssqAttribute != null)
            {
                UDF_NAME = udfName;

                Url = ssqAttribute.Url;
                StockIdentifierPlaceholder = ssqAttribute.StockIdentifierPlaceholder;

                XPath = ssqAttribute.XPath;

                CssSelector = ssqAttribute.CssSelector;

                RegExPattern = ssqAttribute.RegExPattern;
                RegExMatchIndex = ssqAttribute.RegExMatchIndex;
                if (ssqAttribute.RegExGroupName != null) { RegExGroupName = ssqAttribute.RegExGroupName; }

                Culture = new CultureInfo(ssqAttribute.Locale);

                Parser = ssqAttribute.Parser;
            }
        }

        public Ssq(string url,
                   string xPath = null,
                   string cssSelector = null,
                   string regExPattern = null, int regExMatchIndex = 0, string regExGroupName = "quote",
                   string stockIdentifierPlaceholder = "{ISIN_TICKER_WKN}",
                   string locale = null,
                   Parser parser = Parser.Auto,
                   [System.Runtime.CompilerServices.CallerMemberName] string udfName = "") : base(udfName, typeof(Ssq))
        {
            UDF_NAME = udfName;

            Url = url;
            StockIdentifierPlaceholder = stockIdentifierPlaceholder;

            XPath = xPath;

            CssSelector = cssSelector;

            RegExPattern = regExPattern;
            RegExMatchIndex = regExMatchIndex;
            RegExGroupName = regExGroupName;

            Culture = new CultureInfo(locale ?? CultureInfo.CurrentCulture.Name);

            Parser = parser;
        }

        public Ssq(QueryParameter queryParameter,
                   [System.Runtime.CompilerServices.CallerMemberName] string udfName = "") : base(udfName, typeof(Ssq))
        {
            UDF_NAME = udfName;

            Url = queryParameter.Url;
            StockIdentifierPlaceholder = queryParameter.StockIdentifierPlaceholder;

            XPath = queryParameter.XPath;

            CssSelector = queryParameter.CssSelector;

            RegExPattern = queryParameter.RegExPattern;
            RegExMatchIndex = queryParameter.RegExMatchIndex.GetValueOrDefault();
            RegExGroupName = queryParameter.RegExGroupName;

            Culture = new CultureInfo(queryParameter.Locale ?? CultureInfo.CurrentCulture.Name);

            Parser = queryParameter.Parser;
        }
        #endregion

        #region Properties
        [RequiredAttribute(ErrorMessage = "URL is required.")]
        public string Url { get; set; }

        public string StockIdentifierPlaceholder { get; set; } = "{ISIN_TICKER_WKN}";

        public string XPath { get; set; }

        public string CssSelector { get; set; }

        public string RegExPattern { get; set; } = String.Format(@"(?<quote>\d+{0}\d+)", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);

        public int RegExMatchIndex { get; set; } = 0;

        public string RegExGroupName { get; set; } = "quote";

        public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;

        public Parser Parser { get; set; } = Parser.Auto;
        #endregion

        #region Methods
        public object TryQ(string wkn_isin_ticker, string info = null) { return Q(wkn_isin_ticker, info, tryQ: true); }

        public object Q(string wkn_isin_ticker, string info = null) { return Q(wkn_isin_ticker, info, tryQ: false); }

        private object Q(string stockIdentifier, string regExGroupName, bool tryQ = false)
        {
            object value;

            try
            {
                Log.Debug($"UDF={UDF_NAME}" +
                          $"; URL={Url}" +
                          $"; XPath={XPath};" +
                          $"; CssSelector={CssSelector};" +
                          $"; RegExPattern={RegExPattern}; RegExMatchIndex={RegExMatchIndex}; RegExGroupName={RegExGroupName}" +
                          $"; Culture={Culture.Name}; StockIdentifierPlaceholder={StockIdentifierPlaceholder}" +
                          $"; Parser={Parser.ToString()}");

                RequiredAttribute.Check(this);

                if (String.IsNullOrEmpty(XPath)
                    && String.IsNullOrEmpty(CssSelector)
                    && String.IsNullOrEmpty(RegExPattern))
                {
                    throw new SsqException("XPath, CSS Selector or Regular Expression is required.");
                }

                if (!String.IsNullOrEmpty(RegExPattern))
                {
                    regExGroupName = String.IsNullOrEmpty(regExGroupName) ? RegExGroupName : regExGroupName;
                    if (!Regex.IsMatch(RegExPattern, $@"\(\?\<{regExGroupName}\>.*\)"))
                    {
                        throw new SsqException("RegEx pattern must contain the RegEx group name.");
                    }
                }

                if (!String.IsNullOrEmpty(StockIdentifierPlaceholder)
                    && Url.Split(new[] { StockIdentifierPlaceholder }, StringSplitOptions.None).Length == 1) //1 = No split.
                {
                    throw new SsqException($"URL must contain stock identifier placeholder {StockIdentifierPlaceholder}");
                }

                string url = Url.Replace(StockIdentifierPlaceholder, stockIdentifier);
                var quote = FfeWebController.GetValueFromWeb(url, xPath: XPath,
                                                                  cssSelector: CssSelector,
                                                                  regExPattern: RegExPattern, regExGroup: regExGroupName, regExMatchIndex: RegExMatchIndex,
                                                                  parser: Parser);

                //TODO: Implement type converter?
                /* TypeCode tc = (TypeCode)Enum.Parse(typeof(TypeCode), "Decimal");
                value = Convert.ChangeType(quote.value, tc, Culture); */
                var parseSucceeded = decimal.TryParse(quote.value, NumberStyles.Any, Culture, out decimal price);
                if (parseSucceeded)
                {
                    value = price;
                }
                else
                {
                    value = quote.value;
                }
            }
            catch (XPathException ex)
            {
                Log.Error(ex.Message);
                Log.Error("XPath=" + ex.XPath);
                Log.Error("HTML=" + ex.Html);

                if (tryQ)
                {
                    return SsqExcelError("XPATH_ERROR");
                }
                else
                {
                    throw;
                }
            }
            catch (CssSelectorException ex)
            {
                Log.Error(ex.Message);
                Log.Error("CssSelector=" + ex.CssSelector);
                Log.Error("HTML=" + ex.Html);

                if (tryQ)
                {
                    return SsqExcelError("CSS_SELECTOR_ERROR");
                }
                else
                {
                    throw;
                }
            }
            catch (RegExException ex)
            {
                Log.Error(ex.Message);
                Log.Error("RegExPattern=" + ex.Pattern);
                Log.Error("Input=" + ex.Input);

                if (tryQ)
                {
                    return SsqExcelError("REGEX_ERROR");
                }
                else
                {
                    throw;
                }
            }
            catch (Exception ex)
            {
                Log.Error(string.Concat("{@ExceptionMessage}", $" {ex.InnerException.Message}"), ex.Message);

                if (tryQ)
                {
                    if (ex is FormatException || ex is InvalidCastException) { return SsqExcelError("PARSE_ERROR"); }
                    if (ex is FFE.WebException) { return SsqExcelError("WEB_ERROR"); }
                    if (ex is System.Net.WebException) { return SsqExcelError("WEB_ERROR"); }
                    return ExcelError.ExcelErrorGettingData;
                }
                else
                {
                    throw;
                }
            }

            return value;
        }
        #endregion

        #region Excel Errors
        public object SsqExcelError(string errorIdentifier = null)
        {
            return $"#{UDF_NAME}" + (!String.IsNullOrEmpty(errorIdentifier) ? "_" + errorIdentifier : "");
        }
        #endregion

        #region Exceptions
        [Serializable]
        public class SsqException : Exception
        {
            public SsqException() { }

            public SsqException(string message) : base(message) { }

            public SsqException(string message, Exception innerException) : base(message, innerException) { }

            protected SsqException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        }
        #endregion
    }
}