using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace FFE
{
    static class FfeAssembly
    {
        #region Assembly Attribute Accessors
        public static string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public static Version AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version;
            }
        }

        public static string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public static string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public static string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public static string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }

        // Get Bitness (32/64) based on FFE module (file) name.
        public static string AssemblyBitness
        {
            get
            {
                System.Collections.Generic.IEnumerable<ProcessModule> excelModules = Process.GetCurrentProcess().Modules.Cast<ProcessModule>();
                ProcessModule ffeModule = excelModules.FirstOrDefault(p =>
                {
                    return p.ModuleName.Equals(string.Format(FfeSetting.Default.ReleaseFileName, "x32"))
                           || p.ModuleName.Equals(string.Format(FfeSetting.Default.ReleaseFileName, "x64"));
                });
                if (ffeModule != null)
                {
                    if (ffeModule.ModuleName.Contains("x32"))
                        return "32-bit";
                    else if (ffeModule.ModuleName.Contains("x64"))
                        return "64-bit";
                    else return "";
                }
                return "";
            }
        }
        #endregion
    }
}