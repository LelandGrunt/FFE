# Ribbon / Options

FFE comes with an own Excel ribbon tab.
The ribbon includes function specific options and general options of FFE.

<img src="Images/Ribbon_Options.md - FFE Tab.png" />

Extended View:
<img src="Images/Ribbon_Options.md - FFE Tab (Extended View).png" />

The ribbon currently contains the following groups:

* `AVQ`
  Go to [AVQ](AVQ) for a description of the AVQ options.

* [`FFE`](#ffe-group)
  The FFE group contains FFE general options.
The options effects all FFE functions or effects the basic functionality.
  
  

## FFE Group

* `Refresh All`
  (Re-)Calculates all FFE user-defined functions in all sheets.
  Status of the calculations is shown in the Excel status bar.
* `Refresh Sheet`
  (Re-)Calculates all FFE user-defined functions in the current sheet.
  Status of the calculations is shown in the Excel status bar.
* `Refresh Selection`
  (Re-)Calculates all FFE user-defined functions in the current selection.
  Status of the calculations is shown in the Excel status bar.
* `View Log`
  Shows the Excel-DNA and the FFE log entries.
  First starting point in case of issues.
  Set level of logging via `Log Level`.
* `(Re-)Register Functions`
  (Re-)Register all FFE user-defined functions.
  Helpful, when you use a customized [SsqUdf.json](SsqUdf.json) and want to change or test SSQ definitions at Excel run time.
* Options:
  * `Extended View`
    Shows/Hides defined FFE ribbon items.
    If disabled, `(Re-)Register Functions`, `Logging`, `File Logging`, `Log Level` are hidden (becomes active after Excel restart).
  * `Register UDFs on Startup`
    If disabled, no FFE user-defined functions are registered on Excel startup.
    Use `(Re-)Register Functions` to register manually.
  * `SSQ Auto Update`
    Downloads automatically the latest SSQ definitions from the FFE GitHub repository at Excel startup.
    The update may contain new FFE user-defined functions and/or update existing ones. This may be required, when a website changes its layout.
  * `Check Update on Startup`
    Checks for a new FFE release on Excel startup.
    Use `Check for Update` to check manually.
* `Logging`
  Logs events to the Excel-DNA Log Viewer and to log files (if enabled).
  Open Log Viewer via `View Log`.
* `File Logging`
  Logs events to log files (`Logging` must be enabled).
  Each user-defined function as well as the FFE framework create their own log file.
Log files are saved in the FFE root folder (where the FFE Add-in `FFE[x32|x64].xll` file is saved).
* `Log Level`
  Set level of events to log.
  * `Fatal`
    Logs only Fatal events.
  * `Error`
    Logs only Error and Fatal events.
  * `Warning`
    Logs only Warning and higher events.
  * `Information` (Default)
    Logs only Information and higher events.
  * `Debug`
    Logs only Debug and higher events.
    Use this log level if you want to debug a customized [SsqUdf.json](SsqUdf.json).
  * `Verbose`
  Logs all events.
* `Help`
  Link to the [GitHub FFE project site](https://github.com/LelandGrunt/FFE ).
* `Info`
  * `About`
    Shows the FFE Add-in *About* dialog.
  * `Check for Update`
    Checks for a new FFE release.