using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FFE
{
    public static class SsqLoader
    {
        private static readonly ILogger log;

        static SsqLoader()
        {
            if (SsqSetting.Default.EnableLogging)
            {
                string udf = "SSQ";
                Log.Logger = FfeLogger.CreateSubLogger(udf, SsqSetting.Default.LogLevel);
                log = Log.ForContext("UDF", udf);
            }

            Load();
        }

        public static Uri JsonUri { get; private set; }
        public static string Json { get; private set; }
        public static SsqJson SsqJson { get; private set; }
        public static IEnumerable<SsqExcelFunction> SsqExcelFunctions { get; private set; }

        public static void Load()
        {
            JsonUri = LoadJsonUri();

            try
            {
                Json = LoadJson(JsonUri);
            }
            catch (Exception)
            {
                log.Warning("SSQ JSON file {@JsonUri} could not be loaded. Use the embedded local one.", JsonUri);
                Json = SsqResource.SsqUdf;
            }

            SsqJson = LoadSsqJson(Json);
            SsqExcelFunctions = LoadFunctions(SsqJson);
        }

        public static IEnumerable<SsqExcelFunction> GetExcelFunctions()
        {
            return SsqExcelFunctions;
        }

        public static IEnumerable<KeyValuePair<string, UserDefinedFunction>> GetUdfs(bool onlyEnabled = true)
        {
            if (onlyEnabled)
            {
                return SsqJson.UDF.Where(x => x.Value.QueryInformation.Enabled);
            }
            else
            {
                return SsqJson.UDF;
            }
        }

        private static Uri LoadJsonUri()
        {
            Uri uri;

            string jsonFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, SsqSetting.Default.JsonFile);
            if (File.Exists(jsonFile))
            {
                uri = new Uri(jsonFile, UriKind.Absolute);
            }
            else
            {
                string jsonUri = SsqSetting.Default.JsonUri;

                if (!Uri.TryCreate(jsonUri, UriKind.Absolute, out uri))
                {
                    uri = new Uri(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, jsonUri), UriKind.Absolute);
                }
            }

            log.Debug("Loaded SSQ JSON URI: {@SsqJsonUri}", uri);

            return uri;
        }

        private static string LoadJson(Uri uri)
        {
            string json = null;

            if (SsqSetting.Default.AutoUpdate)
            {
                if (uri.Scheme.Equals("http")
                    || uri.Scheme.Equals("https"))
                {
                    json = FfeWeb.GetHttpResponseContent(uri);
                }
                else if (uri.Scheme.Equals("file"))
                {
                    json = File.ReadAllText(uri.LocalPath);
                }
                else
                {
                    throw new Ssq.SsqException("Not supported URI scheme. Supported schemes are http(s) and file.");
                }
            }
            else
            {
                log.Debug("Auto Update of SSQ JSON is disabled. Use the embedded local one.");
                json = SsqResource.SsqUdf;
            }

            log.Debug("Loaded SSQ JSON text.", json);
            log.Verbose("SSQ JSON text: {@SsqJsonText}", json);

            return json;
        }

        private static SsqJson LoadSsqJson(string json)
        {
            SsqJson ssqJson = SsqJson.FromJson(json);

            log.Debug("Loaded SSQ JSON object.");

            return ssqJson;
        }

        private static IEnumerable<SsqExcelFunction> LoadFunctions(SsqJson ssqJson)
        {
            IEnumerable<SsqExcelFunction> ssqExcelFunctions = SsqDelegate.GetSsqExcelFunctions(ssqJson);

            log.Debug("Loaded SSQ Excel functions.");

            return ssqExcelFunctions;
        }
    }
}