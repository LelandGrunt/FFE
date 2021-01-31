using Serilog;

namespace FFE
{
    public abstract class SsqLog
    {
        public readonly ILogger log;

        protected SsqLog(string udf)
        {
            if (SsqSetting.Default.EnableLogging)
            {
                udf = string.IsNullOrEmpty(udf) ? "SSQ" : udf;
                Log.Logger = FfeLogger.CreateSubLogger(udf, SsqSetting.Default.LogLevel);
                log = Log.ForContext("UDF", udf);
            }
        }
    }
}