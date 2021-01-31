using Bluegrams.Application;
using ExcelDna.Integration;
using ExcelDna.Registration;
using Serilog;
using System;
using System.Linq;

namespace FFE
{
    public class FfeAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            PortableSettingsProvider.AllRoaming = true;
            PortableSettingsProvider.SettingsFileName = "FFE.config";
            PortableSettingsProvider.ApplyProvider(AvqSetting.Default);
            PortableSettingsProvider.ApplyProvider(CbqSetting.Default);
            PortableSettingsProvider.ApplyProvider(PlugInSetting.Default);
            PortableSettingsProvider.ApplyProvider(SsqSetting.Default);
            PortableSettingsProvider.ApplyProvider(FfeSetting.Default);

            Log.Logger = FfeLogger.CreateDefaultLogger();

            if (FfeSetting.Default.CheckUpdateOnStartup)
            {
                FfeUpdate.CheckUpdate(false);
            }

            if (FfeSetting.Default.RegisterFunctionsOnStartup)
            {
                RegisterFunctions();
                RegisterDelegates();
            }
        }

        public void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .Where(x => x.FunctionAttribute.Category.Equals("FFE"))
                             .Select(f =>
                             {
                                 Log.Debug("Registered FFE function: {@FfeFunction}", f.FunctionAttribute.Name);
                                 return f;
                             })
                             .RegisterFunctions();
        }

        public void RegisterDelegates()
        {
            try
            {
                SsqLoader.GetExcelFunctions()
                         .RegisterFunctions();
            }
            catch (Exception ex)
            {
                Log.Error("Exception while registering SSQ Excel functions. Message: {@ExceptionMessage}", ex.Message);
            }
        }

        public void AutoClose()
        {
            // Do nothing
        }
    }
}