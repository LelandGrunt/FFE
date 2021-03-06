# Changelog

This document records all notable changes to FFE.



## [1.2.0](https://github.com/LelandGrunt/FFE/compare/v1.1.0...v1.2.0) - 2021-01-31

### Added

* New [Openiban](https://github.com/lelandgrunt/ffe/wiki/Openiban) functions:

  * QBIC
  * QIBAN
  * QCOUNTRIES
* AVQ option: Stop Refresh on Error

### Changed

- Make settings portable

### Fixed

* Refresh All / Sheet function
* (File) Logging



## [1.1.0](https://github.com/LelandGrunt/FFE/compare/v1.0.0...v1.1.0) - 2020-05-29

### Added

* New [AVQ](https://github.com/lelandgrunt/ffe/wiki/AVQ) functions:
  * QAVQ
  * QAVID
  * QAVDA
  * QAVW
  * QAVWA
  * QAVM
  * QAVMA
  * QAVTS
* New [BestMatch](https://github.com/lelandgrunt/ffe/wiki/AVQ#qavd) argument.
* New [OutputSize](https://github.com/lelandgrunt/ffe/wiki/AVQ#qavd) argument (QAVD / QAVDA).

### Removed

* AVQ Batch Query function: BATCH_STOCK_QUOTES API was removed by [Alpha Vantage](https://www.alphavantage.co/).



## 1.0.0 - 2019-12-29

* First public release.