using Serilog;
using System;

namespace FFE
{
    public abstract class FfeBase
    {
        protected FfeBase()
        {
            Log.Logger = FfeLogger.CreateDefaultLogger();
        }

        protected FfeBase(string fileName)
        {
            Log.Logger = FfeLogger.CreateDefaultLogger(fileName);
        }

        protected FfeBase(string fileName, Type type)
        {
            Log.Logger = FfeLogger.CreateDefaultLogger(fileName)
                                  .ForContext("UDF", type);
        }
    }
}