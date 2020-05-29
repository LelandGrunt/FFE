# AVQ Function

The **AVQ** (Alpha Vantage Queries) function group returns stock data from the [Alpha Vantage](https://www.alphavantage.co) (AV) provider.

Alpha Vantage provides a free API for realtime and historical data on stocks and other finance data in JSON or CSV formats. AVQ is a wrapper to get these data via a Excel user-defined function to Excel. AVQ currently supports the following Alpha Vantage APIs:

* [GLOBAL_QUOTE](https://www.alphavantage.co/documentation/#latestprice): Returns the latest price and volume information.

  See Excel Function [QAVQ](#qavq).

* [TIME_SERIES_INTRADAY](https://www.alphavantage.co/documentation/#intraday): Intraday time series (open, high, low, close, volume).

  See Excel Function [QAVID](#qavid).

* [TIME_SERIES_DAILY](https://www.alphavantage.co/documentation/#daily): Daily time series (daily open, daily high, daily low, daily close, daily volume).

  See Excel Function [QAVD](#qavd).

* [TIME_SERIES_DAILY_ADJUSTED](https://www.alphavantage.co/documentation/#dailyadj): Daily time series (daily open, daily high, daily low, daily close, daily volume, daily adjusted close, and split/dividend).

  See Excel Function [QAVDA](#qavda).

* [TIME_SERIES_WEEKLY](https://www.alphavantage.co/documentation/#weekly): Weekly time series (last trading day of each week, weekly open, weekly high, weekly low, weekly close, weekly volume).

  See Excel Function [QAVW](#qavw).

* [TIME_SERIES_WEEKLY_ADJUSTED](https://www.alphavantage.co/documentation/#weeklyadj): Weekly adjusted time series (last trading day of each week, weekly open, weekly high, weekly low, weekly close, weekly adjusted close, weekly  volume, weekly dividend).

  See Excel Function [QAVWA](#qavwa).

* [TIME_SERIES_MONTHLY](https://www.alphavantage.co/documentation/#monthly): Monthly time series (last trading day of each month, monthly open, monthly high, monthly low, monthly close, monthly volume).

  See Excel Function [QAVM](#qavm).

* [TIME_SERIES_MONTHLY_ADJUSTED](https://www.alphavantage.co/documentation/#monthlyadj): Monthly adjusted time series (last trading day of each month, monthly  open, monthly high, monthly low, monthly close, monthly adjusted close,  monthly volume, monthly dividend).

  See Excel Function [QAVMA](#qavma).

Examples:

* **=QAVQ("MSFT")**
* **=QAVID("MSFT", "volume")**
* **=QAVD("MSFT","volume",,"2019-12-27")**

Note: AVQ requires a free Alpha Vantage API Key, that can be requested on [www.alphavantage.co](https://www.alphavantage.co/support/#api-key) (only e-mail is required).
The free API service is limited up to 5 API requests per minute and 500 requests per day. Buy a [premium API key](https://www.alphavantage.co/premium/) which provides a higher API call volume.

AVQ is an independent development and has no relationship to Alpha Vantage.
In general, the same Alpha Vantage [term of services](https://www.alphavantage.co/terms_of_service/) apply.



## Functions

* [QAVQ](#qavq)
* [QAVID](#qavid)
* [QAVD](#qavd)
* [QAVDA](#qavda)
* [QAVW](#qavw)
* [QAVWA](#qavwa)
* [QAVM](#qavm)
* [QAVMA](#qavma)
* [QAVTS](#qavts)



### QAVQ

The QAVQ function uses the Alpha Vantage [GLOBAL_QUOTE](https://www.alphavantage.co/documentation/#latestprice) API to return the latest price and volume information for the given stock symbol.

| Excel Formula          | Result                                              |
| ---------------------- | --------------------------------------------------- |
| =QAVQ("MSFT")          | Returns the latest price of Microsoft Corporation.  |
| =QAVQ("MSFT","volume") | Returns the latest volume of Microsoft Corporation. |

**Syntax**

QAVQ(symbol,[info])

| Argument Name     | Description                                                  |
| ----------------- | ------------------------------------------------------------ |
| symbol (required) | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)   | The stock info to return. Supported values are `open`, `high`, `low`, `price`, `volume`, `latest trading day`, `previous close`, `change` and `change percent` (case insensitive). Default is `price`. |

**Examples**
<img src="Images/AVQ.md - QAVQ Examples.png" />

------

### QAVID

The QAVID function uses the Alpha Vantage [TIME_SERIES_INTRADAY](https://www.alphavantage.co/documentation/#intraday) API to return realtime or historical stock data for the given stock symbol.

| Excel Formula               | Result                                                       |
| --------------------------- | ------------------------------------------------------------ |
| =QAVID("MSFT")              | Returns the latest "close" stock quote of Microsoft Corporation. |
| =QAVID("MSFT","volume")     | Returns the latest trading volume of Microsoft Corporation.  |
| =QAVID("MSFT",,4,"15min")   | Returns the 4th "close" stock quote of Microsoft Corporation based on a 15 minutes interval. Hint: Should be the stock quote that was valid for 45 minutes ago. |
| =QAVID("MSFT",,2,"60min")   | Returns the 2nd "close" stock quote of Microsoft Corporation based on a 60 minutes interval. Hint: Should be the stock quote that was valid for 60 minutes ago. |
| =QAVID("MSFT",,101,,"full") | Returns the 101st "close" stock quote of Microsoft Corporation from the Alpha Vantage query result based on a 5 minutes interval. |

**Syntax**

QAVID(symbol,[info],[data_point_index],[interval],[output_size])

| Argument Name               | Description                                                  |
| --------------------------- | ------------------------------------------------------------ |
| symbol (required)           | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)             | The stock info to return. Supported values are `open`, `high`, `low`, `close` and `volume` (case insensitive). Default is `close`. |
| data_point_index (optional) | The index (zero-based) of the time series data point. Default is `0`. |
| interval (optional)         | The time interval between two data points. Valid values are `1min`, `5min`, `15mn`, `30min` and `60min`. Default is `5min`. |
| output_size (optional)      | The output size of the returned data points. Valid values are `compact` or `full` (case insensitive). Default is `compact`. |

**Examples**
<img src="Images/AVQ.md - QAVID Examples.png" />

------

### QAVD

The QAVD function uses the Alpha Vantage [TIME_SERIES_DAILY](https://www.alphavantage.co/documentation/#daily) API to return daily stock data for the given stock symbol.

| Excel Formula                        | Result                                                       |
| ------------------------------------ | ------------------------------------------------------------ |
| =QAVD("MSFT")                        | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVD("MSFT","close")                | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVD("MSFT","high",-2)              | Returns the "high" stock quote of Microsoft Corporation from two days ago. |
| =QAVD("MSFT","open",5)               | Returns the 5th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation. |
| =QAVD("MSFT","volume",,"2019-12-27") | Returns the trading volume of Microsoft Corporation of 2019-12-27. Note: The given trading date must be contained in the AV response. |
| =QAVD("MSFT",,,"2020-01-01",TRUE)    | Returns the "close" stock quote of Microsoft Corporation of 2019-12-31. Hint: No new year's day stock quote is available, latest available stock quote is then 2019-12-31 (best_match = TRUE). |
| =QAVD("MSFT","close",101,,,"full")   | Returns the 101st "close" stock quote from the Alpha Vantage query result of Microsoft Corporation. |

**Syntax**

QAVD(symbol,[info],[trading_day],[trading_date],[best_match],[output_size])

| Argument Name           | Description                                                  |
| ----------------------- | ------------------------------------------------------------ |
| symbol (required)       | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)         | The stock info to return. Supported values are `open`, `high`, `low`, `close` and `volume` (case insensitive). Default is `close`. |
| trading_day (optional)  | Day/X-th item of the time series. If `trading_day` = 0 then most recent data point is selected (Date unspecific). If `trading_day` < 0 then data point of current date minus `trading_day` is selected (Date specific). If `trading_day` > 0 then the x-th (x = day) data point is selected (Date unspecific). Default is `0`. |
| trading_date (optional) | The trading date. Note: The compact output size of the API returns only the last 100 data points/trading days. If the trading date is given, then it overrules the trading day argument. No default value. |
| best_match (optional)   | Finds best trading data point, if no trading (day/date) argument matches. Valid values are `FALSE` (= 0) or `TRUE` (<> 0). Default is `FALSE`. |
| output_size (optional)  | The output size of the returned data points. Valid values are `compact` or `full` (case insensitive). Default is `compact`. |

**Examples**
<img src="Images/AVQ.md - QAVD Examples.png" />

------

### QAVDA

The QAVDA function uses the Alpha Vantage [TIME_SERIES_DAILY_ADJUSTED](https://www.alphavantage.co/documentation/#dailyadj) API to return daily adjusted stock data for the given stock symbol.

| Excel Formula                         | Result                                                       |
| ------------------------------------- | ------------------------------------------------------------ |
| =QAVDA("MSFT")                        | Returns the recent "adjusted close" stock quote of Microsoft Corporation. |
| =QAVDA("MSFT","close")                | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVDA("MSFT","high",-2)              | Returns the "high" stock quote of Microsoft Corporation from two days ago. |
| =QAVDA("MSFT","open",5)               | Returns the 5th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation. |
| =QAVDA("MSFT","volume",,"2019-12-27") | Returns the trading volume of Microsoft Corporation of 2019-12-27. Note: The given trading date must be contained in the AV response. |
| =QAVDA("MSFT",,,"2020-01-01",TRUE)    | Returns the "adjusted close" stock quote of Microsoft Corporation of 2019-12-31. Hint: No new year's day stock quote is available, latest available stock quote is then 2019-12-31 (best_match = TRUE). |
| =QAVDA("MSFT","close",101,,,"full")   | Returns the 101st "close" stock quote from the Alpha Vantage query result of Microsoft Corporation. |

**Syntax**

QAVDA(symbol,[info],[trading_day],[trading_date],[best_match],[output_size])

| Argument Name           | Description                                                  |
| ----------------------- | ------------------------------------------------------------ |
| symbol (required)       | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)         | The stock info to return. Supported values are `open`, `high`, `low`, `close`, `adjusted close`, `volume`, `dividend amount` and `split coefficient` (case insensitive). Default is `adjusted close`. |
| trading_day (optional)  | Day/X-th item of the time series. If `trading_day` = 0 then most recent data point is selected (Date unspecific). If `trading_day` < 0 then data point of current date minus `trading_day` is selected (Date specific). If `trading_day` > 0 then the x-th (x = day) data point is selected (Date unspecific). Default is `0`. |
| trading_date (optional) | The trading date. Note: The compact output size of the API returns only the last 100 data points/trading days. If the trading date is given, then it overrules the trading day argument. No default value. |
| best_match (optional)   | Finds best trading data point, if no trading (day/date) argument matches. Valid values are `FALSE` (= 0) or `TRUE` (<> 0). Default is `FALSE`. |
| output_size (optional)  | The output size of the returned data points. Valid values are `compact` or `full` (case insensitive). Default is `compact`. |

**Examples**
<img src="Images/AVQ.md - QAVDA Examples.png" />

------

### QAVW

The QAVW function uses the Alpha Vantage [TIME_SERIES_WEEKLY](https://www.alphavantage.co/documentation/#weekly) API to return weekly stock data for the given stock symbol. The latest data point is the prices and volume information for the week (or partial week) that contains the current trading day. All other data points contain weekly stock information from last trading day of the week (usually of Friday).

| Excel Formula                        | Result                                                       |
| ------------------------------------ | ------------------------------------------------------------ |
| =QAVW("MSFT")                        | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVW("MSFT","close")                | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVW("MSFT","high",-2)              | Returns the "high" stock quote of Microsoft Corporation from two weeks ago. Hint: Trading date is calculated by subtraction `trading week` (in weeks) from the current date. The calculated trading date must be contained in the AV response. |
| =QAVW("MSFT","open",5)               | Returns the 5th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation. |
| =QAVW("MSFT","volume",,"2020-01-03") | Returns the trading volume of Microsoft Corporation of 2020-01-03. Note: The given trading date must be contained in the AV response. |
| =QAVW("MSFT",,,"2020-01-01",TRUE)    | Returns the "close" stock quote of Microsoft Corporation of 2019-12-27. Hint: No new year's day stock quote is available, latest available stock quote is then 2019-12-27 (best_match = TRUE). |

**Syntax**

QAVW(symbol,[info],[trading_week],[trading_date],[best_match])

| Argument Name           | Description                                                  |
| ----------------------- | ------------------------------------------------------------ |
| symbol (required)       | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)         | The stock info to return. Supported values are `open`, `high`, `low`, `close` and `volume` (case insensitive). Default is `close`. |
| trading_week (optional) | Week/X-th item of the time series. If `trading_week` = 0 then most recent data point is selected (Date unspecific). If `trading_week` < 0 then data point of current date minus `trading_week` is selected (Date specific). If `trading_week` > 0 then the x-th (x = week) data point is selected (Date unspecific). Default is `0`. |
| trading_date (optional) | The trading date. If the trading date is given, then it overrules the trading week argument. No default value. |
| best_match (optional)   | Finds best trading data point, if no trading (week/date) argument matches. Valid values are `FALSE` (= 0) or `TRUE` (<> 0). Default is `FALSE`. |

**Examples**
<img src="Images/AVQ.md - QAVW Examples.png" />

------

### QAVWA

The QAVWA function uses the Alpha Vantage [TIME_SERIES_WEEKLY_ADJUSTED](https://www.alphavantage.co/documentation/#weeklyadj) API to return weekly adjusted stock data for the given stock symbol. The latest data point is the prices and volume information for the week (or partial week) that contains the current trading day. All other data points contain weekly stock information from last trading day of the week (usually of Friday).

| Excel Formula                         | Result                                                       |
| ------------------------------------- | ------------------------------------------------------------ |
| =QAVWA("MSFT")                        | Returns the recent "adjusted close" stock quote of Microsoft Corporation. |
| =QAVWA("MSFT","close")                | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVWA("MSFT","high",-2)              | Returns the "high" stock quote of Microsoft Corporation from two weeks ago. Hint: Trading date is calculated by subtraction `trading week` (in weeks) from the current date. The calculated trading date must be contained in the AV response. |
| =QAVWA("MSFT","open",5)               | Returns the 5th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation. |
| =QAVWA("MSFT","volume",,"2020-01-03") | Returns the trading volume of Microsoft Corporation of 2020-01-03. Note: The given trading date must be contained in the AV response. |
| =QAVWA("MSFT",,,"2020-01-01",TRUE)    | Returns the "close" stock quote of Microsoft Corporation of 2019-12-27. Hint: No new year's day stock quote is available, latest available stock quote is then 2019-12-27 (best_match = TRUE). |

**Syntax**

QAVWA(symbol,[info],[trading_week],[trading_date],[best_match])

| Argument Name           | Description                                                  |
| ----------------------- | ------------------------------------------------------------ |
| symbol (required)       | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)         | The stock info to return. Supported values are `open`, `high`, `low`, `close`, `adjusted close`, `volume` and `dividend amount` (case insensitive). Default is `adjusted close`. |
| trading_week (optional) | Week/X-th item of the time series. If `trading_week` = 0 then most recent data point is selected (Date unspecific). If `trading_week` < 0 then data point of current date minus `trading_week` is selected (Date specific). If `trading_week` > 0 then the x-th (x = week) data point is selected (Date unspecific). Default is `0`. |
| trading_date (optional) | The trading date. If the trading date is given, then it overrules the trading week argument. No default value. |
| best_match (optional)   | Finds best trading data point, if no trading (week/date) argument matches. Valid values are `FALSE` (= 0) or `TRUE` (<> 0). Default is `FALSE`. |

**Examples**
<img src="Images/AVQ.md - QAVWA Examples.png" />

------

### QAVM

The QAVM function uses the Alpha Vantage [TIME_SERIES_MONTHLY](https://www.alphavantage.co/documentation/#monthly) API to return monthly stock data for the given stock symbol. The latest data point is the prices and volume information for the month (or partial month) that contains the current trading day. All other data points contain monthly stock information from the last trading day of the month (usually the last working day of the month).

| Excel Formula                        | Result                                                       |
| ------------------------------------ | ------------------------------------------------------------ |
| =QAVM("MSFT")                        | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVM("MSFT","close")                | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVM("MSFT","high",-2)              | Returns the "high" stock quote of Microsoft Corporation from two months ago. Hint: Trading date is calculated by subtraction `trading month` (in months) from the current date. The calculated trading date must be contained in the AV response. |
| =QAVM("MSFT","open",5)               | Returns the 5th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation. |
| =QAVM("MSFT","volume",,"2020-01-31") | Returns the trading volume of Microsoft Corporation of 2020-01-31. Note: The given date must be contained in the AV response. |
| =QAVM("MSFT",,,"2019-09-30",TRUE)    | Returns the "close" stock quote of Microsoft Corporation of 2019-09-29. Hint: No stock quote is available on weekends, latest available stock quote is then 2019-12-29 (best_match = TRUE). |

**Syntax**

QAVM(symbol,[info],[trading_month],[trading_date],[best_match])

| Argument Name            | Description                                                  |
| ------------------------ | ------------------------------------------------------------ |
| symbol (required)        | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)          | The stock info to return. Supported values are `open`, `high`, `low`, `close` and `volume` (case insensitive). Default is `close`. |
| trading_month (optional) | Month/X-th item of the time series. If `trading_month` = 0 then most recent data point is selected (Date unspecific). If `trading_month` < 0 then data point of current date minus `trading_month` is selected (Date specific). If `trading_month` > 0 then the x-th (x = month) data point is selected (Date unspecific). Default is `0`. |
| trading_date (optional)  | The trading date. If the trading date is given, then it overrules the trading month argument. No default value. |
| best_match (optional)    | Finds best trading data point, if no trading (month/date) argument matches. Valid values are `FALSE` (= 0) or `TRUE` (<> 0). Default is `FALSE`. |

**Examples**
<img src="Images/AVQ.md - QAVM Examples.png" />

------

### QAVMA

The QAVMA function uses the Alpha Vantage [TIME_SERIES_MONTHLY_ADJUSTED](https://www.alphavantage.co/documentation/#monthlyadj) API to return monthly adjusted stock data for the given stock symbol. The latest data point is the prices and volume information for the month (or partial month) that contains the current trading day. All other data points contain monthly stock information from the last trading day of the month (usually the last working day of the month).

| Excel Formula                         | Result                                                       |
| ------------------------------------- | ------------------------------------------------------------ |
| =QAVMA("MSFT")                        | Returns the recent "adjusted close" stock quote of Microsoft Corporation. |
| =QAVMA("MSFT","close")                | Returns the recent "close" stock quote of Microsoft Corporation. |
| =QAVMA("MSFT","high",-2)              | Returns the "high" stock quote of Microsoft Corporation from two weeks ago. Hint: Trading date is calculated by subtraction `trading month` (in months) from the current date. The calculated trading date must be contained in the AV response. |
| =QAVMA("MSFT","open",5)               | Returns the 5th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation. |
| =QAVMA("MSFT","volume",,"2020-01-03") | Returns the trading volume of Microsoft Corporation of 2020-01-03. Note: The given date must be contained in the AV response. |
| =QAVMA("MSFT",,,"2020-01-01",TRUE)    | Returns the "close" stock quote of Microsoft Corporation of 2019-12-27. Hint: No new year's day stock quote is available, latest available stock quote is then 2019-12-27 (best_match = TRUE). |

**Syntax**

QAVMA(symbol,[info],[trading_month],[trading_date],[best_match])

| Argument Name           | Description                                                  |
| ----------------------- | ------------------------------------------------------------ |
| symbol (required)       | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)         | The stock info to return. Supported values are `open`, `high`, `low`, `close`, `adjusted close`, `volume` and `dividend amount` (case insensitive). Default is `adjusted close`. |
| trading_week (optional) | Month/X-th item of the time series. If `trading_month` = 0 then most recent data point is selected (Date unspecific). If `trading_month` < 0 then data point of current date minus `trading_month` is selected (Date specific). If `trading_month` > 0 then the x-th (x = month) data point is selected (Date unspecific). Default is `0`. |
| trading_date (optional) | The trading date. If the trading date is given, then it overrules the trading month argument. No default value. |
| best_match (optional)   | Finds best trading data point, if no trading (month/date) argument matches. Valid values are `FALSE` (= 0) or `TRUE` (<> 0). Default is `FALSE`. |

**Examples**
<img src="Images/AVQ.md - QAVMA Examples.png" />

------

### QAVTS

The QAVMA function is a wrapper for all Alpha Vantage time series APIs (except GLOBAL_QUOTE). The interval argument determines the API that is used.

| Excel Formula                                | Result                                                       |
| -------------------------------------------- | ------------------------------------------------------------ |
| =QAVTS("MSFT")                               | Returns the recent "close" stock quote of Microsoft Corporation from the [TIME_SERIES_DAILY](https://www.alphavantage.co/documentation/#daily) API. |
| =QAVTS("MSFT","close")                       | Returns the recent "close" stock quote of Microsoft Corporation from the [TIME_SERIES_DAILY](https://www.alphavantage.co/documentation/#daily) API. |
| =QAVTS("MSFT","high","weekly",-2)            | Returns the "high" stock quote of Microsoft Corporation from the [TIME_SERIES_WEEKLY](https://www.alphavantage.co/documentation/#weekly) API from two weeks ago. |
| =QAVTS("MSFT","open","5min")                 | Returns the recent "open" stock quote of Microsoft Corporation from the [TIME_SERIES_INTRADAY](https://www.alphavantage.co/documentation/#intraday) API. |
| =QAVTS("MSFT","volume",,,"2020-01-03",TRUE)  | Returns the trading volume of Microsoft Corporation of 2020-01-03 from the [TIME_SERIES_DAILY_ADJUSTED](https://www.alphavantage.co/documentation/#dailyadj) API. Note: The given date must be contained in the AV response. |
| =QAVTS("MSFT",,"daily",101,,,"compact",TRUE) | Returns the oldest available "close" stock quote of Microsoft Corporation from the [TIME_SERIES_DAILY](https://www.alphavantage.co/documentation/#daily) API. Hint: In output_size `compact` only 100 data points are available, oldest available stock quote is the one from the 100th data point (best_match = TRUE). |

**Syntax**

QAVTS(symbol,[info],[interval],[trading_day],[trading_date],[adjusted],[output_size],[best_match])

| Argument Name            | Description                                                  |
| ------------------------ | ------------------------------------------------------------ |
| symbol (required)        | The symbol of the stock. The symbol can be a string, or a cell reference like A2. |
| info (optional)          | The stock info to return. Supported values are `open`, `high`, `low`, `close`, `adjusted close`, `volume`, `dividend amount` and `split coefficient` (case insensitive). Default is `close`. |
| interval (optional)      | Defines the AlphaVantage API to use. Valid values are `daily`, `weekly`, `monthly`, `1min`, `5min`, `15mn`, `30min` and `60min` (case insensitive). Default is `daily`. |
| trading_day (optional)*  | Day/X-th item of the time series. If `trading_day` = 0 then most recent data point is selected (Date unspecific). If `trading_day` < 0 then data point of current date minus `trading_day` is selected (Date specific). If `trading_day` > 0 then the x-th (x = day) data point is selected (Date unspecific). Default is `0`. |
| trading_date (optional)* | The trading date. If the trading date is given, then it overrules the trading day argument. No default value. |
| adjusted (optional)°     | Query adjusted values or not. Valid values are `FALSE` (= 0) or `TRUE` (<> 0). Default is `FALSE`. |
| output_size (optional)^  | The output size of the returned data points. Valid values are `compact` or `full` (case insensitive). Default is `compact`. |
| best_match (optional)*   | Finds best trading data point, if no trading (day/date) argument matches. Valid values are `FALSE` (= 0) or `TRUE` (<> 0). Default is `FALSE`. |

\* Used for interval `daily`, `weekly`, and `monthly` only.

° Used for interval `daily` only.

^ Used for interval `daily` and [`1`\|`5`\|`15`\|`30`\|`60`]`min` only.

**Examples**
<img src="Images/AVQ.md - QAVTS Examples.png" />



## Examples

The examples shown above can be downloaded <a href="Attachments/AVQ Examples.xlsx">here</a>.



## Best Practices

1. If you retrieve the FFE formula error `#AV_CALL_LIMIT_REACHED`, then you reached the Alpha Vantage API call volume limit (of 5 requests per minute or 500 requests per day for the free service, see also [Alpha Vantage FAQs](https://www.alphavantage.co/support/#support)). For this case FFE has the `Refresh #AV-Errors` ribbon button. This button re-calculates only AVQ functions, that previously returned an #AV_-error.
   <img src="Images/AVQ.md - Refresh AV-Errors.png" style="zoom: 33%" />
   
   Buy a [premium API key](https://www.alphavantage.co/premium/) from Alpha Vantage with a higher API call volume for mass requests.
   Note: FFE cannot provide support for issues related to premium API keys.
   
   
   

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
* `Refresh #AV-Errors`
  (Re-)Calculates all AVQ user-defined functions which previously returned an FFE formula #AV_-error.
* `Help`
  * `Alpha Vantage: Homepage`
    Link to the [Alpha Vantage Homepage](https://www.alphavantage.co).
  * `Alpha Vantage: Term of Services`
    Link to the [Alpha Vantage Term of Services](https://www.alphavantage.co/terms_of_service/).
  * `Alpha Vantage: Get Free API Key`
    Link to the Alpha Vantage [Claim your API Key](https://www.alphavantage.co/support/#api-key) web page.