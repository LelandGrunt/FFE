# CBQ Functions

The **CBQ** (Consorsbank Queries) function group returns the stock prices from the [Consorsbank](https://www.consorsbank.de) website.

Example:

* **=QCB("US5949181045")**
* **=QCBF("US5949181045")**

Consorsbank is a german direct bank and a brand of BNP Paribas.
Consorsbank provides stock information (e.g. price) via their website.

| Excel Formula               | Result                                                       |
| --------------------------- | ------------------------------------------------------------ |
| =QCB("US5949181045")        | Returns the current stock ask price of Microsoft Corporation. |
| =QCB("US5949181045", "GAT") | Returns the current stock ask price of Microsoft Corporation from the stock exchange *Tradegate* (=GAT). |
| =QCBF("US5949181045")       | Returns the current stock ask price of Microsoft Corporation with formatting. Format is `#,##0.00 <Currency Code>`. |

Examples for stock exchanges:

| Stock Exchange | Code |
| -------------- | ---- |
| Düsseldorf     | DUS  |
| Frankfurt      | FRA  |
| Hamburg        | HAM  |
| Hannover       | HAN  |
| München        | MUN  |
| Nasdaq         | NMS  |
| Stuttgart      | STU  |
| Tradegate      | GAT  |
| Xetra          | GER  |

The stock exchange which is used by default depends on the Consorsbank website.



## Syntax

QCB\[F\](isin_wkn,[stock_exchange_code])

| Argument Name                  | Description                                                  |
| ------------------------------ | ------------------------------------------------------------ |
| isin_wkn (mandatory)           | The International Securities Identification Number or Wertpapierkennnummer (Germany). |
| stock_exchange_code (optional) | The code of stock exchange (see also above examples for stock exchanges). |



## Examples

Function examples:
<img src="Images/CBQ.md - CBQ Examples.png" width="50%" height="50%" />

<a href="Attachments/CBQ Examples.xlsx">CBQ Examples.xlsx</a>