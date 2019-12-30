# SSQ Function

With **SSQ** (Simple Stock Query) the user can define and add new functions (user-defined functions) to Excel in a declaratively way without coding and re-compiling of the FFE Add-in.

The heart of SSQ is a text file ([SsqUdf.json](SsqUdf.json)) in JSON format, which contains the definitions of the functions.
Based on these definitions, FFE dynamically creates new user-defined functions at Excel startup or during run time via the FFE  `(Re-)Register Functions` ribbon button.

FFE contains an embedded version of the SSQ text file ([SsqUdf.json](https://github.com/LelandGrunt/FFE/blob/master/FFE-AddIn/SSQ/SsqUdf.json)) with predefined functions. The FFE QJE function is defined in this file, for example.

The user can create his own version of this text file and thus add new or replace existing functions (which are defined in the embedded version). The own file has precedence over the embedded one.
This text file must be saved under the name `SsqUdf.json` in the root folder where the FFE Add-in (FFEx32|FFEx64.xll) is located. Changes in the JSON file are applied at Excel startup or during Excel run time via the FFE  `(Re-)Register Functions` ribbon button. The error because of repeated and overwritten functions can be ignored, if existing functions want to be changed.

All FFE user-defined functions are registered under the Excel function category `FFE`.
<img src="Images/SSQ.md - Excel function category FFE.png" style="zoom:33%;" />



SSQ contains currently the following embedded functions (with example):

| Excel Formula        | Result                                                       |
| -------------------- | ------------------------------------------------------------ |
| =QJE("IE00B4L5Y983") | Returns the current stock price of ETF "iShares Core MSCI World UCITS ETF USD (Acc)" from the www.justetf.com website. |
| =QYF("MSFT")         | Returns the current real time stock price of Microsoft Corporation from the [Yahoo Finance](https://finance.yahoo.com) website. |

SSQ downloads automatically the latest definitions from the FFE GitHub repository at Excel startup (Auto Update). The update may contain new FFE user-defined functions and/or update existing ones. This may be required, when a website changes its layout. The Auto Update function can be disabled via option `SSQ Auto Update`.
<img src="Images/SSQ.md - SSQ Auto Update.png" style="zoom:33%;" />



## Syntax

QJE(isin)

| Argument name   | Description                                                  |
| --------------- | ------------------------------------------------------------ |
| isin (required) | The International Securities Identification Number (ISIN) of the stock. Can be a string, or a cell reference like A2. |

QYF(ticker)

| Argument name     | Description                                                  |
| ----------------- | ------------------------------------------------------------ |
| ticker (required) | The ticker of the stock. Can be a string, or a cell reference like A2. |



## Examples

Function examples:
<img src="Images/SSQ.md - SSQ Examples.png" style="zoom:50%;" />

<a href="Attachments/SSQ Examples.xlsx">SSQ Examples.xlsx</a>