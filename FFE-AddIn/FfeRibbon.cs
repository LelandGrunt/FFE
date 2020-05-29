using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Registration;
using Microsoft.Office.Interop.Excel;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Forms = System.Windows.Forms;

namespace FFE
{
    //https://github.com/Excel-DNA/Samples/tree/master/Ribbon
    [ComVisible(true)]
    public partial class FfeRibbon : ExcelRibbon
    {
        private readonly Application excelApp = null;

        private IEnumerable<string> functionNames = null;

        public FfeRibbon()
        {
            excelApp = (Application)ExcelDnaUtil.Application;
        }

        public override string GetCustomUI(string RibbonID)
        {
            return FfeResource.FfeRibbon;
        }

        public void RefreshSelection(IRibbonControl control)
        {
            if (FfeExcel.IsInEditingMode()) { return; }

            excelApp.Selection.CalculateRowMajorOrder();
        }

        public void RefreshSheet(IRibbonControl control)
        {
            Refresh(excelApp.ActiveSheet);
        }

        public void RefreshAll(IRibbonControl control)
        {
            Refresh();
        }

        public string SetStringDefault(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "ddcFfeLogLevel":
                    return DropDownItemIdToLogEventLevel.FirstOrDefault(x => x.Value == FfeSetting.Default.LogLevel).Key;
                default:
                    return null;
            }
        }

        public bool SetBoolDefault(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "cbxFfeLogging":
                    return FfeSetting.Default.EnableLogging;
                case "cbxFfeFileLogging":
                    return FfeSetting.Default.LogWriteToFile;
                case "tgbExtendedView":
                    return FfeSetting.Default.RibbonExtendedView;
                case "tgbRegisterFunctionsOnStartup":
                    return FfeSetting.Default.RegisterFunctionsOnStartup;
                case "tgbSsqAutoUpdate":
                    return SsqSetting.Default.AutoUpdate;
                case "tgbCheckUpdateOnStartup":
                    return FfeSetting.Default.CheckUpdateOnStartup;
                default:
                    return false;
            }
        }

        public bool SetVisible(IRibbonControl control)
        {
            // Visible tag has precedence for all other visible-related tags.
            if (bool.TryParse(GetTagValue(control.Tag, VisibleTagKey), out bool visible))
            {
                return visible;
            }
            else if (bool.TryParse(GetTagValue(control.Tag, ExtendedViewTagKey), out bool extendedView))
            {
                return extendedView && FfeSetting.Default.RibbonExtendedView;
            }
            else
            {
                return true;
            }
        }

        public void SetFfeLogging(IRibbonControl control, bool pressed)
        {
            FfeSetting.Default.EnableLogging = pressed;
            FfeSetting.Default.Save();

            Forms.MessageBox.Show("Logging change becomes active after Excel restart.", FfeAssembly.AssemblyTitle, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information);
        }

        public void SetFfeFileLogging(IRibbonControl control, bool pressed)
        {
            FfeSetting.Default.LogWriteToFile = pressed;
            FfeSetting.Default.Save();

            Forms.MessageBox.Show("File Logging change becomes active after Excel restart.", FfeAssembly.AssemblyTitle, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information);
        }

        public void SetFfeLogLevel(IRibbonControl control, string id, int index)
        {
            LogEventLevel logLevel = DropDownItemIdToLogEventLevel.FirstOrDefault(x => x.Key.Equals(id)).Value;

            FfeSetting.Default.LogLevel = logLevel;
            FfeSetting.Default.Save();

            // No user options dialog for specific UDF currently exists, hence set the same log level for all UDFs.
            AvqSetting.Default.LogLevel = logLevel;
            AvqSetting.Default.Save();

            CbqSetting.Default.LogLevel = logLevel;
            CbqSetting.Default.Save();

            Forms.MessageBox.Show("Log Level change becomes active after Excel restart.", FfeAssembly.AssemblyTitle, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information);
        }

        public void SetExtendedView(IRibbonControl control, bool pressed)
        {
            FfeSetting.Default.RibbonExtendedView = pressed;
            FfeSetting.Default.Save();

            Forms.MessageBox.Show("Extended View change becomes active after Excel restart.", FfeAssembly.AssemblyTitle, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information);
        }

        public void SetRegisterFunctionsOnStartup(IRibbonControl control, bool pressed)
        {
            FfeSetting.Default.RegisterFunctionsOnStartup = pressed;
            FfeSetting.Default.Save();
        }

        public void SetSsqAutoUpdate(IRibbonControl control, bool pressed)
        {
            SsqSetting.Default.AutoUpdate = pressed;
            SsqSetting.Default.Save();
        }

        public void SetCheckUpdateOnStartup(IRibbonControl control, bool pressed)
        {
            FfeSetting.Default.CheckUpdateOnStartup = pressed;
            FfeSetting.Default.Save();
        }

        public void ViewLog(IRibbonControl control)
        {
            ExcelDna.Logging.LogDisplay.Show();
        }

        public void ReRegisterFunctions(IRibbonControl control)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                SsqLoader.Load();

                FfeAddIn ffeAddIn = new FfeAddIn();
                ffeAddIn.RegisterFunctions();
                ffeAddIn.RegisterDelegates();
            });
        }

        public void OpenHelpLink(IRibbonControl control)
        {
            OpenLink(control.Tag);
        }

        public void ShowFfeAbout(IRibbonControl control)
        {
            string title = string.Format("About {0}", FfeAssembly.AssemblyTitle);
            string productName = FfeAssembly.AssemblyProduct;
            string bitness = !string.IsNullOrEmpty(FfeAssembly.AssemblyBitness) ? $"({FfeAssembly.AssemblyBitness})" : null;
            string description = FfeAssembly.AssemblyDescription;
            string version = string.Format("Version {0}", FfeAssembly.AssemblyVersion.ToString());
            string ssqVersion = string.Format("SSQ Version {0} of {1}", SsqLoader.SsqJson.Version, SsqLoader.SsqJson.VersionDate.ToShortDateString());
            string copyright = FfeAssembly.AssemblyCopyright;
            Version excelDnaVersion = System.Reflection.Assembly.GetAssembly(typeof(IExcelAddIn)).GetName().Version;
            string poweredBy = $"Powered by Excel-DNA (https://excel-dna.net) Version {excelDnaVersion.Major}.{excelDnaVersion.Minor}";

            string text = $"{productName} {bitness}\n{description}\n{version}\n{ssqVersion}\n\n{copyright}\n\n{poweredBy}";

            Forms.MessageBox.Show(text, title, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information);
        }

        public void CheckUpdate(IRibbonControl control)
        {
            FfeUpdate.CheckUpdate();
        }

        private void OpenLink(string url)
        {
            System.Diagnostics.Process.Start(url);
        }

        private const string DefaultValueTagKey = "DefaultValue";
        private const string VisibleTagKey = "Visible";
        private const string ExtendedViewTagKey = "ExtendedView";

        private string GetTagValue(string tag, string name)
        {
            string value = null;

            try
            {
                if (!string.IsNullOrEmpty(tag))
                {
                    string[] tags = tag.Split(';');
                    foreach (string t in tags)
                    {
                        string[] keyValuePair = t.Split(new[] { ":=" }, StringSplitOptions.None);
                        string key = keyValuePair[0];

                        if (key.Equals(name))
                        {
                            value = keyValuePair[1];
                            break;
                        }
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
            return value;
        }

        //TODO: Convert to JSON based mapping file.
        private Dictionary<string, LogEventLevel> DropDownItemIdToLogEventLevel = new Dictionary<string, LogEventLevel>(6)
        {
            { "ddcFfeLogLevelItemFatal", LogEventLevel.Fatal },
            { "ddcFfeLogLevelItemError", LogEventLevel.Error },
            { "ddcFfeLogLevelItemWarning", LogEventLevel.Warning },
            { "ddcFfeLogLevelItemInformation", LogEventLevel.Information },
            { "ddcFfeLogLevelItemDebug", LogEventLevel.Debug },
            { "ddcFfeLogLevelItemVerbose", LogEventLevel.Verbose }
        };

        private void Refresh(Worksheet worksheet = null, IEnumerable<string> functionNames = null)
        {
            try
            {
                if (FfeExcel.IsInEditingMode()) { return; }

                //excelApp.Calculate();
                //excelApp.CalculateFull();
                //excelApp.CalculateFullRebuild();

                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;

                this.functionNames = functionNames ?? this.functionNames ?? GetFfeFunctionNames();

                foreach (string functionName in this.functionNames)
                {
                    string formula = functionName + "(";
                    List<Range> cells = FfeExcel.FindCellsByFormula(formula, worksheet);
                    if (cells.Count > 0)
                    {
                        //excelApp.StatusBar = $"Recalculate {functionName}...";
                        FfeExcel.Recalculate(cells, functionName);
                    }
                }
            }
            finally
            {
                excelApp.ScreenUpdating = true;
                excelApp.DisplayAlerts = true;
                //excelApp.StatusBar = null;
            }
        }

        private IEnumerable<string> GetFfeFunctionNames()
        {
            /*functionNames = GetExecutingAssembly().GetTypes()
                                                  .SelectMany(t => t.GetMethods())
                                                  .Where(m => m.GetCustomAttributes(typeof(ExcelFunctionAttribute), false).Length > 0)
                                                  .Select(mi => mi.Name);*/
            // FFE functions
            functionNames = ExcelRegistration.GetExcelFunctions()
                                             .Where(x => x.FunctionAttribute.Category.Equals("FFE"))
                                             .Select(f => f.FunctionAttribute.Name);
            // + SSQ UDFs
            functionNames = functionNames.Union(SsqLoader.GetUdfs().Select(udf => udf.Value.QueryInformation.Name));

            return functionNames;
        }
    }
}