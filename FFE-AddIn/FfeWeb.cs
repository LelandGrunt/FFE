using Serilog;
using System;
using System.IO;
using System.Net;
using System.Net.Http;

namespace FFE
{
    public class FfeWeb
    {
        public bool GenerateRandomUserAgent { get; set; } = true;

        private static readonly HttpClient httpClient;

        private static readonly ILogger log;

        static FfeWeb()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                                                 | SecurityProtocolType.Tls11
                                                 | SecurityProtocolType.Tls
                                                 | SecurityProtocolType.Ssl3;

            HttpClientHandler handler = new HttpClientHandler()
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.None
            };

            httpClient = new HttpClient(handler);

            log = Log.ForContext("UDF", "FFE");
        }

        public static string GetHttpResponseContent(Uri uri, bool generateUserAgent = false)
        {
            return GetHttpResponseContent(uri.AbsoluteUri, generateUserAgent);
        }

        public static string GetHttpResponseContent(string uri, bool generateUserAgent = false)
        {
            if (generateUserAgent)
            {
                SetValidGeneratedUserAgent();
            }

            using (HttpResponseMessage response = httpClient.GetAsync(uri).Result)
            {
                response.EnsureSuccessStatusCode();

                using (HttpContent content = response.Content)
                {
                    string result = content.ReadAsStringAsync().Result;

                    return result;
                }
            }
        }

        public static byte[] GetHttpResponseContentAsByteArray(string uri, bool generateUserAgent = false)
        {
            if (generateUserAgent)
            {
                SetValidGeneratedUserAgent();
            }

            using (HttpResponseMessage response = httpClient.GetAsync(uri).Result)
            {
                response.EnsureSuccessStatusCode();

                using (HttpContent content = response.Content)
                {
                    return content.ReadAsByteArrayAsync().Result;
                }
            }
        }

        public static Stream GetHttpResponseContentAsStreamReader(string uri, bool generateUserAgent = false)
        {
            if (generateUserAgent)
            {
                SetValidGeneratedUserAgent();
            }

            /*Stream stream = httpClient.GetStreamAsync(uri).Result;
            return new StreamReader(stream);*/
            return httpClient.GetStreamAsync(uri).Result;
        }

        public static string GetValidUserAgent(int maxTries = 5)
        {
            // If a invalid user agent string is generated, then try again (max. <maxTries> times).
            string userAgent = null;

            int tries = 0;
            bool validUserAgent = false;
            do
            {
                userAgent = new Bogus.DataSets.Internet().UserAgent();
                HttpRequestMessage httpRequestMessage = new HttpRequestMessage();
                validUserAgent = httpRequestMessage.Headers.UserAgent.TryParseAdd(userAgent);
                tries++;
            } while (!validUserAgent
                     && tries <= maxTries);

            return userAgent;
        }

        #region Private
        private static void SetValidGeneratedUserAgent(int maxTries = 5)
        {
            try
            {
                // If a invalid user agent string is generated, then try again (max. <maxTries> times).
                string userAgent = null;

                int tries = 0;
                bool validUserAgent = false;
                do
                {
                    userAgent = new Bogus.DataSets.Internet().UserAgent();
                    httpClient.DefaultRequestHeaders.UserAgent.Clear();
                    validUserAgent = httpClient.DefaultRequestHeaders.UserAgent.TryParseAdd(userAgent);
                    tries++;
                } while (!validUserAgent
                         && tries <= maxTries);
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
            }
        }
        #endregion
    }
}