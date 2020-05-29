using ExcelDna.Logging;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using System;
using System.IO;

namespace FFE
{
    public static class FfeLogger
    {
        public static ILogger CreateDefaultLogger(string fileName = "FFE", LogEventLevel? loggingLevelSwitch = null)
        {
            if (FfeSetting.Default.EnableLogging)
            {
                return Log.Logger = ConfigureLogging(loggingLevelSwitch ?? FfeSetting.Default.LogLevel,
                                                     fileName);
            }
            else
            {
                return Log.Logger = Logger.None;
            }
        }

        public static ILogger ConfigureLogging(LogEventLevel loggingLevelSwitch, string fileName = "FFE")
        {
            LoggerConfiguration loggerConfiguration =
                                new LoggerConfiguration()
                                    .MinimumLevel.ControlledBy(new LoggingLevelSwitch(loggingLevelSwitch))
                                    .WriteTo.ExcelDnaLogDisplay(outputTemplate: FfeSetting.Default.LogTemplate,
                                                                displayOrder: DisplayOrder.NewestFirst)
                                    .Enrich.WithThreadId();
            if (FfeSetting.Default.LogWriteToFile)
            {
                loggerConfiguration.WriteTo.File(path: Path.Combine(AppDomain.CurrentDomain.BaseDirectory, FfeSetting.Default.LogFolder, fileName + ".log"),
                                                 outputTemplate: FfeSetting.Default.LogTemplate);
            }

            return loggerConfiguration.CreateLogger();
        }
    }
}