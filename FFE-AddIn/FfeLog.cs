using Serilog;

namespace FFE
{
    public abstract class FfeLog
    {
        private protected readonly ILogger log;

        protected FfeLog(string udf = "FFE")
        {
            if (FfeSetting.Default.EnableLogging)
            {
                Log.Logger = FfeLogger.CreateSubLogger(udf, FfeSetting.Default.LogLevel);
                log = Log.ForContext("UDF", udf);
            }
        }
    }
}