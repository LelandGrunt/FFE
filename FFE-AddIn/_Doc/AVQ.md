# AVQ Function

The **AVQ** (Alpha Vantage Queries) function group returns stock data from the [Alpha Vantage](https://www.alphavantage.co) (AV) provider.

Alpha Vantage provides a free API for realtime and historical data on stocks and other finance data in JSON or CSV formats. AVQ is a wrapper to get these data via a Excel user-defined function to Excel. AVQ currently supports the following Alpha Vantage API:

* [TIME_SERIES_DAILY](https://www.alphavantage.co/documentation/#daily): Daily time series (date, daily open, daily high, daily low, daily close, daily volume)

Examples:

* **=QAVD("MSFT")**
* **=QAVD("MSFT", "volume")**

Note: AVQ requires a free Alpha Vantage API Key, that can be requested on [www.alphavantage.co](https://www.alphavantage.co/support/#api-key) (only e-mail is required).
The free API service is limited up to 5 API requests per minute and 500 requests per day. Use the [Batch Query](AVQ-Batch-Query) to reduce the number of requests or buy a [premium API key](https://www.alphavantage.co/premium/) which provides a higher API call volume.

AVQ is an independent development and has no relationship to Alpha Vantage.
In general, the same Alpha Vantage [term of services](https://www.alphavantage.co/terms_of_service/) apply.

The QAVD function uses the Alpha Vantage [TIME_SERIES_DAILY](https://www.alphavantage.co/documentation/#daily) API to return daily time series (open, high, low, close) and volumes for the given stock symbol.

| Excel Formula                        | Result                                                       |
| ------------------------------------ | ------------------------------------------------------------ |
| =QAVD("MSFT")                        | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVD("MSFT","close")                | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVD("MSFT","high",-2)              | Returns the "high" stock quote of Microsoft Corporation from two days ago. |
| =QAVD("MSFT","open",5)               | Returns the 5th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation. |
| =QAVD("MSFT","volume",,"2019-12-27") | Returns the trading volume of Microsoft Corporation of 2019-12-27. Note: The given date must be contained in the AV response. |



## Syntax

QAVD(symbol,[info],[trading_day],[trading_date])

| Argument Name           | Description                                                  |
| ----------------------- | ------------------------------------------------------------ |
| symbol (required)       | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)         | The stock info to return. Currently supported values are `open`, `high`, `low`, `close` and `volume`. Default is `close`. |
| trading_day (optional)  | Day/X-th item of the time series. If `day` = 0 then most recent data point is selected (Date unspecific). If `day` < 0 then data point of current date minus `day` is selected (Date specific). If `day` > 0 then the x-th (x = day) data point is selected (Date unspecific). Default is `0`. |
| trading_date (optional) | The trading date. Note: The compact output of the TIME_SERIES_DAILY API returns only the last 100 data points. If the trading date is given, then it overrules the trading day argument. No default value. |



## Examples

Function examples:
<img src="Images/AVQ.md - AVQ Examples.png" />

<a href="Attachments/AVQ Examples.xlsx">AVQ Examples.xlsx</a>



## Best Practices

1. If you retrieve the FFE formula error `#AV_CALL_LIMIT_REACHED`, then you reached the Alpha Vantage API call volume limit (of 5 requests per minute or 500 requests per day for the free service, see also [Alpha Vantage FAQs](https://www.alphavantage.co/support/#support)). For this case FFE has the `Refresh #AV-Errors` ribbon button. This button re-calculates only AVQ functions, that previously returned an #AV_-error.
   <img src="Images/AVQ.md - Refresh AV-Errors.png" style="zoom: 33%" />
   
   Buy a [premium API key](https://www.alphavantage.co/premium/) from Alpha Vantage with a higher API call volume for mass requests.
   Note: FFE cannot provide support for issues related to premium API keys.
   
2. If you request more than five stock information (price or volume) per minute, then you will retrieve the `#AV_CALL_LIMIT_REACHED` FFE formula error from the 6th request onwards.
   Alternatively, use the [AVQ batch query](AVQ-Batch-Query) function or wait a minute to afterwards request the next five stock information with the `Refresh #AV-Errors` button (see also point 1.).
   
   

## Common Problems

1. No Alpha Vantage API Key is set.
   Set Alpha Vantage API Key via FFE ribbon button `Set API Key`.
   Default key is "demo", which only works with limited symbols (e.g. "MSFT").

2. To many Alpha Vantage requests per minutes or day (FFE formula error `#AV_CALL_LIMIT_REACHED`).
   See also [best practices](#best-practices).



## Options

All AVQ specific options are available via the ribbon group AVQ in the ribbon tab FFE.
<img src="Images/AVQ.md - Options.png" style="zoom:50%;" />

* `Set API Key`
  Sets the Alpha Vantage API Key.
* AVQ Batch Query options:
  Go to [AVQ Batch Query](AVQ-Batch-Query).
* `Refresh #AV-Errors`
  (Re-)Calculates all AVQ user-defined functions which previously returned an FFE formula #AV_-error.
* `Help`
  * `Alpha Vantage: Homepage`
    Link to the [Alpha Vantage Homepage](https://www.alphavantage.co).
  * `Alpha Vantage: Term of Services`
    Link to the [Alpha Vantage Term of Services](https://www.alphavantage.co/terms_of_service/).
  * `Alpha Vantage: Get Free API Key`
    Link to the Alpha Vantage [Claim your API Key](https://www.alphavantage.co/support/#api-key) web page.