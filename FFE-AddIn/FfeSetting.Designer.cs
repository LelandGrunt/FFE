﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace FFE {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.8.1.0")]
    internal sealed partial class FfeSetting : global::System.Configuration.ApplicationSettingsBase {
        
        private static FfeSetting defaultInstance = ((FfeSetting)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new FfeSetting())));
        
        public static FfeSetting Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} <{ThreadId}> [{Level:u3}] {UDF} {Message:" +
            "lj}{NewLine}{Exception}")]
        public string LogTemplate {
            get {
                return ((string)(this["LogTemplate"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Logs")]
        public string LogFolder {
            get {
                return ((string)(this["LogFolder"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("FFE{0}.xll")]
        public string ReleaseFileName {
            get {
                return ((string)(this["ReleaseFileName"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool RegisterFunctionsOnStartup {
            get {
                return ((bool)(this["RegisterFunctionsOnStartup"]));
            }
            set {
                this["RegisterFunctionsOnStartup"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://api.github.com/repos/LelandGrunt/FFE/releases")]
        public string UpdateUrl {
            get {
                return ((string)(this["UpdateUrl"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://github.com/LelandGrunt/FFE/releases")]
        public string ReleasesUrl {
            get {
                return ((string)(this["ReleasesUrl"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://github.com/LelandGrunt/FFE/blob/master/CHANGELOG.md")]
        public string ChangelogUrl {
            get {
                return ((string)(this["ChangelogUrl"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool LogWriteToFile {
            get {
                return ((bool)(this["LogWriteToFile"]));
            }
            set {
                this["LogWriteToFile"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool RibbonExtendedView {
            get {
                return ((bool)(this["RibbonExtendedView"]));
            }
            set {
                this["RibbonExtendedView"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Information")]
        public global::Serilog.Events.LogEventLevel LogLevel {
            get {
                return ((global::Serilog.Events.LogEventLevel)(this["LogLevel"]));
            }
            set {
                this["LogLevel"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool CheckUpdateOnStartup {
            get {
                return ((bool)(this["CheckUpdateOnStartup"]));
            }
            set {
                this["CheckUpdateOnStartup"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool EnableLogging {
            get {
                return ((bool)(this["EnableLogging"]));
            }
            set {
                this["EnableLogging"] = value;
            }
        }
    }
}
