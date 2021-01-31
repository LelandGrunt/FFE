REM %~1 = "$(ProjectDir)"
REM %~2 = "$(ConfigurationName)"
REM %~3 = "$(DevEnvDir)"

REM Set revision number in assembly version and file version.
"%~3\TextTransform.exe" "%~1\Properties\AssemblyInfo.tt"

REM Apply configuration specific settings (Debug/Test/Release).
powershell.exe -File "%~1_Build\XmlDocTransform.ps1" -XmlFile "%~1app.Default.config" -XdtFile "%~1app.%~2.config" -XmlFileNew "%~1app.config"