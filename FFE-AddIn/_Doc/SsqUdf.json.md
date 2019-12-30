# SsqUdf.json

The **SsqUdf.json** file contains all user-defined function definitions for the FFE [SSQ](SSQ) function.
The file itself is in JSON format and easy to read by humans and machines, and editable with a text editor of your choice.

Each user-defined function definition has two parts:

1. [QueryInformation](#queryinformation):
The first part describes the name and how the function should be displayed in Excel in the function dialog.
   <img src="Images/SsqUdf.json.md - QueryInformation.png" style="zoom:50%;" />
   How it looks like in the Excel function dialog:
   <img src="Images/SsqUdt.json.md - Excel function dialog (QGF).png" style="zoom: 33%;" />
   `Help on this function` is linked to the website which is defined under `Help Topic.`
   The `Help Link` object is currently not used.
   The `Comment` object value is an internal comment and used in the JSON file only.
   
2. [QueryParameter](#queryparameter):
   The second part defines the website and how the stock value should be query from it.
   <img src="Images/SsqUdf.json.md - QueryParameter.png" style="zoom: 50%;" />

At the beginning of the [SsqUdf.json](https://raw.githubusercontent.com/LelandGrunt/FFE/master/FFE-AddIn/SSQ/SsqUdf.json) file there is an additional meta data section which describes all possible objects in `QueryInformation` and ` QueryParameter`.



## QueryInformation

In the QueryInformation part the information for the Excel function dialog is set.
`Help Topic`, `Help Link`, `ExcelArgNameStockIdentifier`, `ExcelArgDescStockIdentifier` and `Comment` are optional. All other objects are mandatory.

| JSON Object                 | Mandatory | Description                                                  |
| --------------------------- | --------- | ------------------------------------------------------------ |
| Name                        | Yes       | Name of the user-defined function (UDF).                     |
| Description                 | Yes       | Description of the UDF (max. 30 characters).                 |
| Author                      | Yes       | Author who wrote this UDF query specification.               |
| Author Email                | Yes       | Authors email.                                               |
| Version                     | Yes       | Version of the UDF query specification. Format: MAJOR.MINOR  |
| Version Date                | Yes       | Date of the UDF query specification. Format: YYYY-MM-DD      |
| Help Topic                  | No        | Link to the UDF help. Linked with the Excel function dialog. Format: http(s):// |
| Help Link                   | No        | Link to the authors UDF help site. Format: http(s)://        |
| Provider                    | Yes       | Source of the UDF query result.                              |
| Enabled                     | Yes       | Register the definition as UDF? Format: true/false           |
| ExcelArgNameStockIdentifier | No        | Name of the UDF stock identifier argument.                   |
| ExcelArgDescStockIdentifier | No        | Description of the UDF stock identifier argument.            |
| Comment                     | No        | Internal comment. Used in the JSON file only.                |

If you want to share your query definitions with other users or want to add to the FFE project, then set all objects.



## QueryParameter

The QueryParameter part defines the website and how the stock value should be query from it.
`Url` and one of the methods (`XPath`, `CssSelector`, `RegExPattern`) are mandatory. All other objects are optional.

| JSON Object                                                  | Mandatory | Description                                                  |
| ------------------------------------------------------------ | --------- | ------------------------------------------------------------ |
| Url                                                          | Yes       | The URL to the website with the stock value. `Url` must contain the `StockIdentifierPlaceholder` object value (default is `{ISIN_TICKER_WKN}`). The `StockIdentifierPlaceholder` is replaced by the first Excel function argument at run time (usually the ISIN/ticker/WKN). |
| StockIdentifierPlaceholder                                   | No        | The placeholder for the stock identifier (ISIN, ticker, symbol or WKN). Default is `{ISIN_TICKER_WKN}`. |
| [XPath](#query-method)                                       | Yes/No    | The XPath expression to query the stock value from the website. Either `XPath`, `CssSelector` or `RegExPattern` object is mandatory. If all objects exists, then `XPath` is prioritized. |
| [CssSelector](#query-method)                                 | Yes/No    | The CSS selector expression to query the stock value from the website. Either `XPath`, `CssSelector` or `RegExPattern` object is mandatory. If all objects exists, then `XPath` is prioritized. |
| [RegExPattern](#query-method)                                | Yes/No    | The regular expression to query the stock value from the website. Either `XPath`, `CssSelector` or `RegExPattern` object is mandatory. If `XPath` or `CssSelector` exists, then regular expression is evaluated on XPath/CssSelector selection result. |
| [RegExGroupName](#query-method)                              | No        | The group name must be part of the regular expression. Only used if `RegExPattern` object exists. Default is `quote`. |
| [RegExMatchIndex](#query-method)                             | No        | The index (zero-based), if the regular expression pattern returns multiple matches. Only used if `RegExPattern` object exists. Default is `0`. |
| [Locale](https://docs.microsoft.com/openspecs/windows_protocols/ms-lcid/a9eac961-e77d-41a6-90a5-ce1a8b0cdb9c) | No        | The `Locale` is the culture that match to the query result. Default is the Locale of the local computer. Must be a predefined CultureInfo identifier (.NET). For example, use `en-US` for stock values which are formatted in the english format on the website. |
| [Parser](#query-parser)                                      | No        | The HTTP client/parser to use for the website query. Valid values are `HAP`, `AngleSharp`, `HttpClient`, `WebClient` or `Auto`. Default is `Auto`. |



### Query Method

The query method defines how the stock value is extracted/queried from the website.
The value of the query method is the search-/query-expression, which is used to select the element (=stock value) from the HTML source code of the website.

Three query methods are supported by FFE (all examples returns the same query result):

1. XML Path ([XPath](https://en.wikipedia.org/wiki/XPath)):
   With the XPath query language elements in a HTML page can be selected by navigating through the HTML DOM tree.  Additional search criterias can limit the selection results. Although the query language is designed for XML documents, in most cases it also works with HTML.
   Example: `//strong[contains(@class,'price')]` select all strong tags which contains the CSS class *price*.
2. [CSS Selector](https://en.wikipedia.org/wiki/Cascading_Style_Sheets#Selector):
   With CSS selectors elements in a HTML page can be selected by CSS pattern matching rules.
   Example: `.price` selects all elements with CSS class *price*.
3. [Regular Expression](https://en.wikipedia.org/wiki/Regular_expression) (RegEx):
   With regular expressions texts can be searched by patterns. In our case the returned HTML source code is searched to find the stock value of our choice.
   Example: `<strong class="price.*.>.*?(?<quote>\d*,\d*)</strong>` selects all HTML source texts which matches the search expression. Here: all numbers with a comma as (decimal) separator, which are embedded in a HTML strong tag and where (CSS) class is the first attribute with value *price*. The inner selection in regular expression group `quote` `(?<quote>\d*,\d*)` is then returned.

Note: The XPath and CSS Selector query methods returns always the first element that matches the expression, while the regular expression query method supports to select the element by index.



**How to find the value for query method object?**

The XPath and the CSS selector expression can normaly be copied by each modern web browser.

1. Open the website with the stock value.
2. Open the developer tools (`F12`).
3. Use the "select tool" to select the stock value in the HTML page.
4. Right click the selection and copy XPath or CSS selector via the copy function.

Other possibilities are to use browser extensions/add-ons (recommended), e.g.

* Ruto - XPath Finder for [Chrome](https://chrome.google.com/webstore/detail/ruto-xpath-finder/ilcoelkkcokgeeijnopjnolmmighnppp) and [Firefox](https://addons.mozilla.org/de/firefox/addon/rutoxpath)
* XPath Helper for [Chrome](https://chrome.google.com/webstore/detail/xpath-helper/hgimnogjllphhhkhlmebbmlgjoejdpjl)
* ChroPath for [Chrome](https://chrome.google.com/webstore/detail/chropath/ljngjbnaijcbncmcnjfhigebomdlkcjo) and [Firefox](https://addons.mozilla.org/de/firefox/addon/chropath-for-firefox)
* [xPath Finder](https://addons.mozilla.org/de/firefox/addon/xpath_finder/) for Firefox

For one query method different expressions for the same query result can be exists. Choose the one which performs the **best** and which consider **variants**. For example, XPath expressions which contains a fix currency are not working with stocks in other currencies. Or a CSS class for rising stock prices may not consider the falling stock prices.

Note: The query result is tried to parse to a decimal value at the end so that Excel can calculate with the value. If parsing fails, the query result is interpreted as a string value.

The expression with the best query performance can be found with the [FFE benchmark tool](https://github.com/LelandGrunt/FFE/tree/master/FFE-AddIn.Benchmark) based on [BenchmarkDotNet](https://benchmarkdotnet.org/). Define the benchmark test with different expressions in file `ParserMethodExpression.json` and then use the `FFE.WebParserMethodExpressionPerformanceWebTests` class to execute.

> .\FFE-AddIn_Benchmark.exe --f FFE.WebParserMethodExpressionPerformanceWeb*



**How to analyse not working expressions?**

In log level *Debug* the HTML source code of the last executed query is saved in the folder *Logs* in the FFE root folder. A separate subfolder is created for each query parser and the files are created with name pattern `PageSource_<Url>.html`. Open the *offline* files in the browser and view the result returned by the query parsers. The query method expression can also tested and copied from the *offline* files.



### Query Parser

The query parser defines the HTTP client used to request the website and the HTML parser used to parse the HTML source code.

Four query parsers are available:

1. [Html Agility Pack](https://html-agility-pack.net/) (HAP)
2. [AngleSharp](https://anglesharp.github.io/)
3. [HttpClient](https://docs.microsoft.com/en-us/dotnet/api/system.net.http.httpclient)
4. [WebClient](https://docs.microsoft.com/en-us/dotnet/api/system.net.webclient)

`HAP` and `AngleSharp` contains a HTTP client and a HTML parser.
`HttpClient` and `WebClient` are HTTP clients only and "routes" the request result (HTML) to the HAP parser (for XPath expressions) or to the AngleSharp parser (for CSS selector expressions).
The differences are in performance and how the HTML source code is interpreted.
The `Auto` value selects the query parser based on specified query method. `Auto` does not select the parser which performance the best.



**Which parser should I use?**

The answer is ambiguous und depends on the website and its HTML source code and the [query method](#query-method) you choose.
Based on benchmark tests the `AngleSharp` parser in combination with the `XPath` method performs well.

In some cases your query method expression may not work with all query parsers. Try another one and choose the one which performs the best.