# AVQ Batch Query Function

The **AVQ Batch Query** allows to query stock quotes for multiple symbols with one request.

In comparison to the `AVQD` user-defined function, the batch query queries stock quotes for multiple symbols with one Alpha Vantage API call. This allows to increase the number of stock quote queries, which are limited by the Alpha Vantage free API service.

The AVQ Batch Query function is not a formula function, but a function that is called via ribbon button `FFE | AVQ | Batch Query`.

The batch query requires a Excel-selection (also named as range) of stock symbols to work. The selection must have two columns. The first column contains the symbols, in the second column the result (stock quote) of the batch query is returned. Multiple Excel-selections with the same selection principle are allowed (see also [Examples](#examples)).
The selection(s) must match with the selected AVQ batch query mode, that defines the kind of selection you choose.
<img src="Images/AVQ Batch Query.md - Modes.png" style="zoom: 50%;" />

* `Contiguous Range`
  The selection contains two columns, which are located next to each other. The range is one area, that is not discontinuous by an other selection.
  The result (stock quote) of the Alpha Vantage batch query is returned next (right) to the symbol column.
  <img src="Images/AVQ Batch Query.md - Contiguous Range Mode.png" style="zoom: 50%;" />
  
* `Non-Contiguous Range`
  The selections contains a total of two columns, which are discontinuous by other columns.
  The result (stock quote) of the Alpha Vantage batch query is returned in the second column.
  <img src="Images/AVQ Batch Query.md - Non-Contiguous Range Mode.png" style="zoom:50%;" />
  
* `Named Range`
  The selections were previously saved as named ranges. The named ranges contains two columns each, which must be a contiguous range (support for non-contiguous ranges are on the roadmap).
  Use * (asterisk) as a wildcard character at the end of *Range Name* 
  
  <img src="Images/AVQ Batch Query.md - Range Name.png" style="zoom:33%;" />
  
  to use all named ranges which starts with the given name.
  The result (stock quote) of the AVQ batch query is returned in the second column.
  <img src="Images/AVQ Batch Query.md - Named Range Mode.png" style="zoom:50%;" />
  
  <img src="Images/AVQ Batch Query.md - Named Ranges - Name Manager.png" style="zoom: 33%;" />



## Options

**Batch Comment**
The `Batch Comment` option adds the stock quote timestamp (retrieved from the Alpha Vantage Batch API) in local time and the provider as a Excel comment to each stock quote cell. This makes the traceability of the stock quotes easier.

<img src="Images/AVQ Batch Query.md - Batch Comment.png" style="zoom:33%;" />
<img src="Images/AVQ Batch Query.md - Batch Comment Example.png" style="zoom:33%;" />



## Examples

<a href="Attachments/AVQ Batch Query Examples.xlsx">AVQ Batch Query Examples.xlsx</a> ([free Alpha Vantage API Key](https://www.alphavantage.co/support/#api-key) required)