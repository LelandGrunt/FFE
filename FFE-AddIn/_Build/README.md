# Build

This document describes the project specific build process of the FFE Add-in.



## Build Events

FFE uses some additional build events, executed by Visual Studio, to modify the (binary) build output and to executed post build tasks.

FFE distinguishes between common and user-specific build events.
Common build events are mandatory and independent of the user environment.
User-specific events are optional and depends on the current user environment.

In the default project configuration the build events are wrapped inside Windows batch files.



### Pre-build events

Pre-builds events are executed before the main build process is executed.

Two common pre-build events are currently defined:

1. Set revision number in assembly version and file version via text template transformation (T4) in the file `AssemblyInfo.cs`.
   Usage of `AssemblyInfo.tt` in the folder Properties.
2. Apply configuration specific settings (Debug/Test/Release) via XSL transformation.
   Usage of `XmlDocTransform.ps1` (PowerShell script) and `Microsoft.Web.XmlTransform.dll` in the _Build folder.

User-specific pre-build events are expected in the batch file PreBuildEvents.%USERNAME%.bat in folder _Build, where %USERNAME% must be replaced by the value of environment variable %USERNAME%. Use the batch file `PreBuildEvents.%USERNAME%.bat` as a template for your specific pre-build events.

```cmd
REM Common pre-build events.
IF EXIST "$(ProjectDir)_Build\PreBuildEvents.bat" CALL "$(ProjectDir)_Build\PreBuildEvents.bat" "$(ProjectDir)" "$(ConfigurationName)" "$(DevEnvDir)"

REM User-specific pre-build events.
IF EXIST "$(ProjectDir)_Build\PreBuildEvents.$(Username).bat" CALL "$(ProjectDir)_Build\PreBuildEvents.$(Username).bat" "$(ProjectDir)" "$(ConfigurationName)" "$(DevEnvDir)"
```



### Post-build events

Post-builds events are executed after the main build process is executed.

Common post-build events are currently not defined.

User-specific post-build events are expected in the batch file PostBuildEvents.%USERNAME%.bat in folder _Build, where %USERNAME% must be replaced by the value of environment variable %USERNAME%. Use the batch file `PostBuildEvents.%USERNAME%.bat` as a template for your specific post-build events. Copying the built Add-in file `FFE-AddInx[32|64]-packed.xll` to the Microsoft Office Add-in folder `%APPDATA%\Microsoft\AddIns\` of the local user is an example for a post-build event.

```cmd
REM Common post-build events.
IF EXIST "$(ProjectDir)_Build\PostBuildEvents.bat" CALL "$(ProjectDir)_Build\PostBuildEvents.bat" "$(ProjectDir)" "$(ConfigurationName)"

REM User-specific post-build events.
IF EXIST "$(ProjectDir)_Build\PostBuildEvents.$(Username).bat" CALL "$(ProjectDir)_Build\PostBuildEvents.$(Username).bat" "$(ProjectDir)" "$(ConfigurationName)"
```