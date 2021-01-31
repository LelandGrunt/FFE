| Build & Test Status                                          | License                                                      | Version                                                      | Wiki                                                         | Changelog                                                    |
| ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ |
| <a href="https://github.com/LelandGrunt/FFE/actions"><img src="https://github.com/LelandGrunt/FFE/workflows/FFE-CI/badge.svg"></a> | <a href="https://github.com/LelandGrunt/FFE/blob/master/LICENSE.md"><img src="https://img.shields.io/github/license/LelandGrunt/FFE"></a> | <a href="https://github.com/LelandGrunt/FFE/releases"><img src="https://img.shields.io/github/v/release/LelandGrunt/FFE?sort=semver"></a> | <a href="https://github.com/LelandGrunt/FFE/wiki"><img src="https://img.shields.io/badge/docs-Wiki-green"></a> | <a href="https://github.com/LelandGrunt/FFE/blob/master/CHANGELOG.md"><img src="https://img.shields.io/badge/docs-Changelog-green"></a> |

# FFE (Financial Functions for Excel)

**FFE** is an Excel Add-in and adds new financial functions to Microsoft Excel.

The FFE Add-in currently provides the following functions:
1. Query of stock information (e.g. price) from different stock data providers and websites.

   * [Alpha Vantage](https://www.alphavantage.co) 
   
   * [Consorsbank](https://www.consorsbank.de)
   
   * [justETF](https://www.justetf.com)
   
   * [Yahoo Finance](https://www.finance.yahoo.com)
   
2. Validation and calculation of IBAN and bank data (e.g. BIC) via the Openiban webservice. (&#8505; **New in version 1.2.0**)

   

**Highlight**: With FFE, new websites and thus new user-defined functions can be easily added by editing a [text file](https://github.com/LelandGrunt/FFE/wiki/SsqUdf.json) ([SSQ Function](https://github.com/LelandGrunt/FFE/wiki/SSQ)).

> ```json
> "QueryInformation": {
>     "Name": "QYF",
>     "Description": "Yahoo finance website query.",
>     "Author": "Leland Grunt",
>     "Author Email": "leland.grunt@gmail.com",
>     "Version": "1.0",
>     "Version Date": "2019-12-29",
>     "Help Topic": "https://github.com/LelandGrunt/FFE",
>     "Help Link": "https://github.com/LelandGrunt/FFE",
>     "Provider": "finance.yahoo.com",
>     "Enabled": true,
>     "ExcelArgNameStockIdentifier": "Ticker",
>     "ExcelArgDescStockIdentifier": "Yahoo-Ticker",
>     "Comment": ""
> },
> "QueryParameter": {
>     "Url": "https://finance.yahoo.com/quote/{TICKER}",
>     "StockIdentifierPlaceholder": "{TICKER}",
>     "CssSelector": ".Fz\\(36px\\)",
>     "Locale": "en-US",
>     "Parser": "AngleSharp"
> }
> ```



The FFE Add-in currently adds in detail the following user-defined functions:

| Group    | Function   | Description                                                  |
| -------- | ---------- | ------------------------------------------------------------ |
| AVQ      | QAVQ       | Query of stock values via the [Alpha Vantage GLOBAL_QUOTE API](https://www.alphavantage.co/documentation/#latestprice). |
| AVQ      | QAVID      | Query of stock values via the [Alpha Vantage TIME_SERIES_INTRADAY API](https://www.alphavantage.co/documentation/#intraday). |
| AVQ      | QAVD       | Query of stock values via the [Alpha Vantage TIME_SERIES_DAILY API](https://www.alphavantage.co/documentation/#daily). |
| AVQ      | QAVDA      | Query of stock values via the [Alpha Vantage TIME_SERIES_DAILY_ADJUSTED API](https://www.alphavantage.co/documentation/#dailyadj). |
| AVQ      | QAVW       | Query of stock values via the [Alpha Vantage TIME_SERIES_WEEKLY API](https://www.alphavantage.co/documentation/#weekly). |
| AVQ      | QAVWA      | Query of stock values via the [Alpha Vantage TIME_SERIES_WEEKLY_ADJUSTED API](https://www.alphavantage.co/documentation/#weeklyadj). |
| AVQ      | QAVM       | Query of stock values via the [Alpha Vantage TIME_SERIES_MONTHLY API](https://www.alphavantage.co/documentation/#monthly). |
| AVQ      | QAVMA      | Query of stock values via the [Alpha Vantage TIME_SERIES_MONTHLY_ADJUSTED API](https://www.alphavantage.co/documentation/#monthlyadj). |
| AVQ      | QAVTS      | Wrapper for all Alpha Vantage time series APIs (except GLOBAL_QUOTE). |
| CBQ      | QCB        | Query of stock prices from the [Consorsbank](https://www.consorsbank.de) website. |
| CBQ      | QCBF       | Query of stock prices from the [Consorsbank](https://www.consorsbank.de) website with currency format. |
| SSQ      | QJE        | Query of stock prices from the [justETF](https://www.justetf.com) website. |
| SSQ      | QYF        | Query of stock prices from the [Yahoo Finance](https://www.finance.yahoo.com) website. |
| Openiban | QBIC       | Query of BICs (Business Identifier Codes).                   |
| Openiban | QIBAN      | Query of IBANs (International Bank Account Numbers).         |
| Openiban | QCOUNTRIES | Query of country names and codes supported by the function QIBAN. |

<img src="https://raw.githubusercontent.com/LelandGrunt/FFE/master/FFE-AddIn/_Doc/Images/README.md%20-%20FFE%20Examples.png" alt="FFE Examples" width="75%" height="75%" />

FFE uses the [Excel-DNA](https://excel-dna.net/) library under the hood. Excel-DNA brings the full power of .NET to Excel.



## Table of Content

* [Installation](#installation)
* [Usage Examples](#usage-examples)
* [Known Issues](#known-issues)
* [Support](#support)
* [Documentation](#documentation)
* [Roadmap](#roadmap)
* [Contributing](#contributing)
* [Changelog](#changelog)
* [License](#license)
* [Credits](#credits)
* [Legal Statement](#legal-statement)



## Installation

The installation is as simple as copying a file.
FFE itself is just one file (`FFE[x32|x64].xll`) which Excel can load as an Add-in.
You can load the Add-in file in two ways:

1. Copy the FFE Add-in file ([FFEx32.xll](https://github.com/LelandGrunt/FFE/releases/latest/download/FFEx32.xll) or [FFEx64.xll](https://github.com/LelandGrunt/FFE/releases/latest/download/FFEx64.xll)) to the folder 
   `%USERPROFILE%\AppData\Roaming\Microsoft\AddIns\` 
   or 
   `%APPDATA%\Microsoft\AddIns\`.
   In this case, Excel automatically loads the FFE Add-in on startup.
   Choose `FFEx32.xll` for a Excel version in 32-bit and `FFEx64.xll` for a Excel version in 64-bit.
   This is the easiest way.
2. In Excel via `File | (Excel) Options | Add-ins | Excel Add-Ins` (< Excel 2013) or via Ribbon `Developer | (Excel) Add-ins` (>= Excel 2013). 
   If Developer Ribbon is not available, add via `File | Options | Customize Ribbon` and then enable `Developer` in the `Customize the Ribbon` list.
   Change your `Trust Center Settings`, if necessary.

Hint: The folder where the FFE file is located is used as a root folder for FFE. Log files (if logging is enabled) are created, as well as a customized [SsqUdf.json](https://github.com/LelandGrunt/FFE/wiki/SsqUdf.json) file is searched, in this folder.



**Additional steps for the AVQ function group (optional):**

For the Alpha Vantage provider, a API key is mandatory.

1. Get a free Alpha Vantage API Key [here](https://www.alphavantage.co/support/#api-key) (only e-mail is required).
2. Set the Alpha Vantage API Key via the `Set API Key` button in the new ribbon tab `FFE`.
   
   

**The FFE Add-in was tested with:**

* Excel 2003 (32-bit)
* Excel 2007 (32-bit)
* Excel 2010 (32-/64-bit)
* Excel 2013 (32-/64-bit)
* Excel 2016 (32-/64-bit)
* Excel 2019 (64-bit)

For Excel 2003 no ribbon interface is available. Use and edit the [FFE.config](https://github.com/LelandGrunt/FFE/releases/latest/download/FFE.config) file to set the [FFE options](https://github.com/LelandGrunt/FFE/wiki/Ribbon_Options) (incl. Alpha Vantage API Key). Save the file in the FFE root folder where the FFE Add-in (FFEx32.xll|FFEx64.xll) is located.



**AutoUpdate/New Version Check**

FFE has a check for a new update on Excel startup. The automatic update check can be disabled via ribbon button `FFE | Options | Check Update on Startup`. 
Check manually via `FFE | Info | Check for Update`.
The update may contain new FFE user-defined functions and/or updates existing ones.




## Usage Examples
Use FFE in your personal Excel based asset reporting to update your current stock values.
<img src="https://raw.githubusercontent.com/LelandGrunt/FFE/master/FFE-AddIn/_Doc/Images/README.md%20-%20My%20Asset%20Report.png" alt="Example for an Asset Report" width="75%" height="75%" />

| Function                                                     | Excel Formula                          | Result                                                       |
| ------------------------------------------------------------ | -------------------------------------- | ------------------------------------------------------------ |
| [QAVQ](https://github.com/LelandGrunt/FFE/wiki/AVQ#qavq)     | =QAVQ("MSFT")                          | Returns the latest stock price of Microsoft Corporation.     |
| [QAVID](https://github.com/LelandGrunt/FFE/wiki/AVQ#qavid)   | =QAVID("MSFT","volume")                | Returns the latest trading volume of Microsoft Corporation.  |
| [QAVD](https://github.com/LelandGrunt/FFE/wiki/AVQ#qavd)     | =QAVD("MSFT","high",-2)                | Returns the "high" stock quote of Microsoft Corporation from two days ago. |
| [QAVDA](https://github.com/LelandGrunt/FFE/wiki/AVQ#qavda)   | =QAVDA("MSFT","open",5)                | Returns the 5th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation. |
| [QAVW](https://github.com/LelandGrunt/FFE/wiki/AVQ#qavw)     | =QAVW("MSFT","volume",,"2020-01-03")   | Returns the trading volume of Microsoft Corporation of 2020-01-03 from the [TIME_SERIES_WEEKLY](https://www.alphavantage.co/documentation/#weekly) API. |
| [QAVWA](https://github.com/LelandGrunt/FFE/wiki/AVQ#qavwa)   | =QAVWA("MSFT",,,"2020-01-01",TRUE)     | Returns the "close" stock quote of Microsoft Corporation of 2019-12-27.  Hint: No new year's day stock quote is available, latest available stock quote is then 2019-12-27 (best_match = TRUE). |
| [QAVM](https://github.com/LelandGrunt/FFE/wiki/AVQ#qavm)     | =QAVM("MSFT")                          | Returns the recent "close" stock quote of Microsoft Corporation from the [TIME_SERIES_MONTHLY](https://www.alphavantage.co/documentation/#monthly) API. |
| [QAVMA](https://github.com/LelandGrunt/FFE/wiki/AVQ#qavma)   | =QAVMA("MSFT","high")                  | Returns the recent "high" stock quote of Microsoft Corporation from the [TIME_SERIES_MONTHLY_ADJUSTED](https://www.alphavantage.co/documentation/#monthlyadj) API. |
| [QAVTS](https://github.com/LelandGrunt/FFE/wiki/AVQ#qavts)   | =QAVTS("MSFT","high","weekly",-2)      | Returns the "high" stock quote of Microsoft Corporation from the [TIME_SERIES_WEEKLY](https://www.alphavantage.co/documentation/#weekly) API from two weeks ago. |
| [QCB](https://github.com/LelandGrunt/FFE/wiki/CBQ)           | =QCB("US5949181045")                   | Returns the current stock ask price of Microsoft Corporation from the [Consorsbank](https://www.consorsbank.de) website. |
| [QCB](https://github.com/LelandGrunt/FFE/wiki/CBQ)           | =QCB("US5949181045","GAT")             | Returns the current stock ask price of Microsoft Corporation of the stock exchange *Tradegate* (=GAT) from the [Consorsbank](https://www.consorsbank.de) website. |
| [QJE](https://github.com/LelandGrunt/FFE/wiki/SSQ)           | =QJE("IE00B4L5Y983")                   | Returns the current stock price of ETF "iShares Core MSCI World UCITS ETF USD (Acc)" from the www.justetf.com website. |
| [QYF](https://github.com/LelandGrunt/FFE/wiki/SSQ)           | =QYF("MSFT")                           | Returns the current real time stock price of Microsoft Corporation from the [Yahoo Finance](https://finance.yahoo.com) website. |
| [QBIC](https://github.com/LelandGrunt/FFE/wiki/Openiban)     | =QBIC("DE89370400440532013000")        | Returns the BIC (Business Identifier Code) of the IBAN *DE89370400440532013000*. |
| [QIBAN](https://github.com/LelandGrunt/FFE/wiki/Openiban)    | =QIBAN("DE", "37040044", "0532013000") | Returns the IBAN (International Bank Account Number) for given bank code and account number for the given country. |
| [QCOUNTRIES](https://github.com/LelandGrunt/FFE/wiki/Openiban) | =QCOUNTRIES()                          | Returns an array of country names and codes supported by the function `QIBAN`. |



## Known Issues

**Excel starts without opening the start screen or a blank workbook.**
This is a bug in Excel (at least Office 365 MSO version 16.0.11328.20468) whenever any .xll is installed.
The issue appears to be fixed in version 16.0.12013.20000.

**Lost FFE settings**
The FFE settings are saved in the `%LOCALAPPDATA%\Microsoft_Corporation\FullTrustSandbox(Excel-DN_Path_<Id>\<Excel Version>\user.config` file. If a new Office/Excel version with a new build number is installed, then a new sub folder is created, and the existing setting/config file is not found.
Since FFE version 1.2.0 the FFE settings are portable and saved directly in the FFE root folder where also the FFE Add-in (FFEx32.xll|FFEx64.xll) is located.



## Support

Via [Issues](https://github.com/LelandGrunt/FFE/issues) or [Discussions](https://github.com/LelandGrunt/FFE/discussions).

Please check the [known issues](#known-issues) and the [existing issue entries](https://github.com/LelandGrunt/FFE/issues) first before creating a new one.



## Documentation

Go to [Wiki](https://github.com/LelandGrunt/FFE/wiki)



## Roadmap

* AVQ: Addition of other available Alpha Vantage APIs (e.g. [CURRENCY_EXCHANGE_RATE](https://www.alphavantage.co/documentation/#currency-exchange)).

* FFE Framework: Plugin interface for custom user-defined functions/methods.

* FFE Framework: AutoUpdate of FFE Add-in (not only check).

* FFE Framework: User-defined functions as [asynchronous calls](https://github.com/Excel-DNA/Registration/blob/master/Source/Samples/Registration.Sample/AsyncFunctionExamples.cs).

* FFE Framework: Replacement of ribbon icons by nicer and more meaningful ones.

* FFE Framework: Addition of a JSON parser.

* FFE Framework: Addition of RTD-based functions/RTD support.

* SSQ: Adding of additional providers (voting via GitHub reactions).

* SSQ: Adding of other web scraping libraries (e.g. [OpenScraping](https://github.com/Microsoft/openscraping-lib-csharp)).

* SSQ: Grouping of SSQ functions with same website/provider.

* SSQ: User interface for adding and editing SSQ query definitions (editing of [SsqUdf.json](https://github.com/LelandGrunt/FFE/wiki/SsqUdf.json)).

* SSQ: Implementing type converter for returned value.

* SSQ: Selection by index for XPath and CssSelector query methods (like `RegExMatchIndex`).

* New UDF: https://frugalisten.de/rechner/ via MS Solver Foundation.



## Contributing
Contributors are welcome!

Project contributing is possible in several ways:
1. Add new user-defined functions by add new web query providers via the FFE [SSQ](https://github.com/LelandGrunt/FFE/wiki/SSQ) function (for IT Experts).
   This is the easiest way and requires no (deeper) programming skills.
   Just edit the [SsqUdf.json](https://github.com/LelandGrunt/FFE/wiki/SsqUdf.json) file and define declaratively which value should be extracted from a website and how the query should be displayed in Excel.
   Go to [SSQ](https://github.com/LelandGrunt/FFE/wiki/SSQ) for a detailed description and to [SsqUdf.json](https://github.com/LelandGrunt/FFE/wiki/SsqUdf.json) for a how-to description.
   [Example](FFE-AddIn/_Doc/Attachments/SsqUdf.json) is available.
2. Add new user-defined functions by writing new C# (static) methods (for Developers).
   The method must return a type that Excel can interpret (usually simple types such as string or double). Go to [Excel-DNA](https://excel-dna.net/) for more information and a getting-started guide.
   [FFE](FFE-AddIn/FfeAttribute.cs) and [Excel-DNA](https://github.com/Excel-DNA/ExcelDna/blob/master/Source/ExcelDna.Integration/ExcelAttributes.cs) attributes are necessary to register the methods/user-defined functions in the FFE context in Excel.
   [Example](FFE-AddIn/CBQ/Cbq.cs) is available.
3. Extend the FFE functionality (for Developers).
   See [Roadmap](#roadmap).
4. Analyze and resolve issues or answer questions (for FFE Experts).
   See [Issues](https://github.com/LelandGrunt/FFE/issues) or [Discussions](https://github.com/LelandGrunt/FFE/discussions).



## Changelog
Go to [Changelog](CHANGELOG.md).



## License
FFE is licensed under the [MIT license](LICENSE.md).



## Credits
Major dependencies are:
* [Excel-DNA](https://www.nuget.org/packages/ExcelDna.AddIn/)

* [HtmlAgilityPack](https://www.nuget.org/packages/HtmlAgilityPack/)

* [AngleSharp](https://www.nuget.org/packages/AngleSharp/)

* [Serilog](https://www.nuget.org/packages/Serilog/)

  
## Legal Statement
FFE use web scraping to extract data from websites.
Web scraping may be illegal in specific/your countries or prohibited by the terms of use of the website.
Do not bypass the web scraping anti-block-techniques of the websites.
Only use FFE for your private and do not extract data in mass.