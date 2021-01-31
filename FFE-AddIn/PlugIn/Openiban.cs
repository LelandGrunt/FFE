using ExcelDna.Integration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace FFE
{
    public static partial class Openiban
    {
        private const string URL_OPEN_IBAN = "https://openiban.com";

        public static string UrlContenResult { private get; set; } = null;

        private static readonly ILogger log;

        static Openiban()
        {
            if (PlugInSetting.Default.EnableLogging)
            {
                string udf = "OPENIBAN";
                Log.Logger = FfeLogger.CreateSubLogger(udf, PlugInSetting.Default.LogLevel);
                log = Log.ForContext("UDF", udf);
            }
        }

        #region Excel Functions
        [FfeFunction(Provider = "openiban.com")]
        [ExcelFunction(Name = "FFE.OPENIBAN.QBIC",
                       Description = "Returns the BIC (Business Identifier Code) for the given IBAN (openiban.com).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QBIC([ExcelArgument(Name = "IBAN", Description = "the IBAN for which the BIC should be returned.")]
                                  string iban)
        {
            const string FUNCTON_NAME = "QBIC";

            // IBAN is mandatory.
            if (string.IsNullOrEmpty(iban))
            {
                return OpenibanExcelError(FUNCTON_NAME, "NULL!");
            }
            // Check on valid IBAN.
            // https://stackoverflow.com/questions/44656264/iban-regex-design
            if (!System.Text.RegularExpressions.Regex.IsMatch(iban, @"^([A-Z]{2}[ \-]?[0-9]{2})(?=(?:[ \-]?[A-Z0-9]){9,30}$)((?:[ \-]?[A-Z0-9]{3,5}){2,7})([ \-]?[A-Z0-9]{1,3})?$"))
            {
                log.Error("Invalid {@IBAN}", iban);
                return OpenibanExcelError(FUNCTON_NAME, "INVALID_IBAN");
            }

            JObject json = null;

            try
            {
                //var ffeAttribute = (FfeFunctionAttribute)Attribute.GetCustomAttribute(typeof(Openiban).GetMethod(System.Reflection.MethodBase.GetCurrentMethod().Name), typeof(FfeFunctionAttribute));
                log.Debug("Querying the BIC for IBAN {@IBAN} from {@Provider}", iban, "openiban.com");

                var endpoint = $"{URL_OPEN_IBAN}/validate/{iban}?getBIC=true&validateBankCode=true";
                log.Debug("Openiban endpoint URL: {@OpenibanEndpointUrl}.", endpoint);

                // TODO: Newtonsoft.Json.JsonConvert.DeserializeObject<List<MyType>(jsonData); //faster with typed object
                if (UrlContenResult == null)
                {
                    json = JObject.Parse(FfeWeb.GetHttpResponseContent(endpoint));
                }
                else // else running tests.
                {
                    json = JObject.Parse(UrlContenResult);
                }

                if ((bool)json["valid"])
                {
                    return (string)json["bankData"]["bic"];
                }
                else
                {
                    log.Error((string)json["messages"][0]);
                    return OpenibanExcelError(FUNCTON_NAME, "INVALID_IBAN");
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, FUNCTON_NAME);
                if (json != null) { log.Error("Openiban JSON response: {@OpenibanResponse}", json.ToString()); }
                return ExcelError.ExcelErrorGettingData;
            }
        }

        [FfeFunction(Provider = "openiban.com")]
        [ExcelFunction(Name = "FFE.OPENIBAN.QCOUNTRIES",
                       Description = "Returns an array of country names and codes supported by the function QIBAN (openiban.com).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object[,] QCOUNTRIES()
        {
            const string FUNCTON_NAME = "QCOUNTRIES";

            string json = null;

            try
            {
                log.Debug("Querying the supported countries from {@Provider}", "openiban.com");

                var endpoint = $"{URL_OPEN_IBAN}/countries";
                log.Debug("Openiban endpoint URL: {@OpenibanEndpointUrl}.", endpoint);

                if (UrlContenResult == null)
                {
                    json = FfeWeb.GetHttpResponseContent(endpoint);
                }
                else // else running tests.
                {
                    json = UrlContenResult;
                }

                Dictionary<string, string> countries = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

                int numberOfCountries = countries.Count;
                object[,] countryArray = new object[numberOfCountries, 2];

                for (int i = 0; i < numberOfCountries; i++)
                {
                    KeyValuePair<string, string> country = countries.ElementAt(i);
                    countryArray[i, 0] = country.Key;
                    countryArray[i, 1] = country.Value;
                }

                return countryArray;
            }
            catch (Exception ex)
            {
                log.Error(ex, FUNCTON_NAME);
                if (json != null) { log.Error("Openiban JSON response: {@OpenibanResponse}", json.ToString()); }
                return new object[0, 0];
            }
        }

        [FfeFunction(Provider = "openiban.com")]
        [ExcelFunction(Name = "FFE.OPENIBAN.QIBAN",
                       Description = "Returns the IBAN (International Bank Account Number) for given bank code and account number for the given country (openiban.com) (Beta).",
                       Category = "FFE",
                       IsThreadSafe = true)]
        public static object QIBAN([ExcelArgument(Name = "Country Code", Description = "the country code (ISO 3166-1 alpha-2 code) for which the IBAN should be calculated.\nUse =QCOUNTRIES() to get an array of supported countries.")]
                                   string countryCode,
                                   [ExcelArgument(Name = "Bank Code", Description = "the bank code for which the IBAN should be calculated.")]
                                   string bankCode,
                                   [ExcelArgument(Name = "Account Number", Description = "the account number for which the IBAN should be calculated.")]
                                   string accountNumber)
        {
            const string FUNCTON_NAME = "QIBAN";

            // All arguments are mandatory (Prevalidation).
            if (string.IsNullOrEmpty(countryCode)
                //|| !System.Text.RegularExpressions.Regex.IsMatch(countryCode, @"^[a-zA-Z]{2}$")
                || string.IsNullOrEmpty(bankCode)
                || string.IsNullOrEmpty(accountNumber))
            {
                return OpenibanExcelError(FUNCTON_NAME, "NULL!");
            }

            JObject json = null;

            try
            {
                log.Debug("Querying the IBAN from {@Provider}", "openiban.com");

                var endpoint = $"{URL_OPEN_IBAN}/calculate/{countryCode}/{bankCode}/{accountNumber}";
                log.Debug("Openiban endpoint URL: {@OpenibanEndpointUrl}.", endpoint);

                if (UrlContenResult == null)
                {
                    json = JObject.Parse(FfeWeb.GetHttpResponseContent(endpoint));
                }
                else // else running tests.
                {
                    json = JObject.Parse(UrlContenResult);
                }

                if ((bool)json["valid"])
                {
                    return (string)json["iban"];
                }
                else
                {
                    log.Error((string)json["message"]);
                    return OpenibanExcelError(FUNCTON_NAME, "INVALID");
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, FUNCTON_NAME);
                if (json != null) { log.Error("Openiban JSON response: {@OpenibanResponse}", json.ToString()); }
                return ExcelError.ExcelErrorGettingData;
            }
        }
        #endregion

        #region Openiban Excel Errors
        public static object OpenibanExcelError(string errorPrefix, string errorIdentifier = null)
        {
            return "#" + errorPrefix + (!String.IsNullOrEmpty(errorIdentifier) ? "_" + errorIdentifier : "");
        }
        #endregion

        #region Exceptions
        [Serializable]
        public class OpenibanException : Exception
        {
            public OpenibanException() { }

            public OpenibanException(string message) : base(message) { }

            public OpenibanException(string message, Exception innerException) : base(message, innerException) { }

            protected OpenibanException(SerializationInfo info, StreamingContext context) : base(info, context) { }
        }
        #endregion
    }
}