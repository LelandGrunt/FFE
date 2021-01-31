using ExcelDna.Logging;
using Serilog;
using Serilog.Context;
using Serilog.Core;
using Serilog.Events;
using Serilog.Filters;
using System;
using System.IO;

namespace FFE
{
    public static class FfeLogger
    {
        public static ILogger CreateDefaultLogger()
        {
            const string UDF = "FFE";

            if (FfeSetting.Default.EnableLogging)
            {
                LoggerConfiguration loggerConfiguration =
                                    new LoggerConfiguration()
                                        .MinimumLevel.ControlledBy(new LoggingLevelSwitch(FfeSetting.Default.LogLevel))
                                        .WriteTo.ExcelDnaLogDisplay(outputTemplate: FfeSetting.Default.LogTemplate,
                                                                    displayOrder: DisplayOrder.NewestFirst)
                                        .Enrich.WithProperty("UDF", UDF)
                                        .Enrich.WithThreadId();

                if (FfeSetting.Default.LogWriteToFile)
                {
                    //loggerConfiguration.Filter.ByIncludingOnly(Matching.WithProperty<string>("UDF", p => p == UDF));
                    loggerConfiguration.WriteTo.File(path: Path.Combine(AppDomain.CurrentDomain.BaseDirectory, FfeSetting.Default.LogFolder, $"{UDF}.log"),
                                                     outputTemplate: FfeSetting.Default.LogTemplate);
                }

                return loggerConfiguration.CreateLogger();
            }
            else
            {
                return Logger.None;
            }
        }

        public static ILogger CreateSubLogger(string udf, LogEventLevel loggingLevelSwitch = LogEventLevel.Debug)
        {
            if (FfeSetting.Default.EnableLogging)
            {
                LoggerConfiguration loggerConfiguration =
                                    new LoggerConfiguration()
                                        .WriteTo.Logger(Log.Logger)
                                        .MinimumLevel.ControlledBy(new LoggingLevelSwitch(loggingLevelSwitch))
                                        .Enrich.WithProperty("UDF", udf)
                                        .Enrich.WithThreadId();

                if (FfeSetting.Default.LogWriteToFile)
                {
                    loggerConfiguration.WriteTo.Logger(lc =>
                    {
                        lc.Filter.ByIncludingOnly(Matching.WithProperty<string>("UDF", p => p == udf));
                        lc.WriteTo.File(path: Path.Combine(AppDomain.CurrentDomain.BaseDirectory, FfeSetting.Default.LogFolder, $"{udf}.log"),
                                        outputTemplate: FfeSetting.Default.LogTemplate);
                    });
                }

                return loggerConfiguration.CreateLogger();
            }
            else
            {
                return Logger.None;
            }
        }
    }
}