# Openiban Functions

The **Openiban** function group validate and calculate IBANs (International Bank Account Numbers) by using the [Openiban](https://openiban.com) webservice.

Examples:

* **=FFE.OPENIBAN.QBIC("DE89370400440532013000")**
* **=FFE.OPENIBAN.QIBAN("DE", "37040044", "0532013000")**
* **=FFE.OPENIBAN.QCOUNTRIES()**

| Excel Formula                                       | Result                                                       |
| --------------------------------------------------- | ------------------------------------------------------------ |
| =FFE.OPENIBAN.QBIC("DE89370400440532013000")        | Returns the BIC (Business Identifier Code) for the given IBAN. |
| =FFE.OPENIBAN.QIBAN("DE", "37040044", "0532013000") | Returns the IBAN (International Bank Account Number) for given bank code and account number for the given country. Supported countries are returned by function `QCOUNTRIES`. |
| =FFE.OPENIBAN.QCOUNTRIES()                          | Returns an array of country names and codes supported by the function `QIBAN`. |



## Functions

* [QBIC](#qbic)
* [QIBAN](#qiban)
* [QCOUNTRIES](#qcountries)



### QBIC

The QBIC function uses the [Openiban](https://openiban.com/) REST Webservice to return the BIC (Business Identifier Code) for the given IBAN.

| Excel Formula                                | Result                                          |
| -------------------------------------------- | ----------------------------------------------- |
| =FFE.OPENIBAN.QBIC("DE89370400440532013000") | Returns the BIC to IBAN DE89370400440532013000. |

**Syntax**

QBIC(iban)

| Argument Name   | Description                                                  |
| --------------- | ------------------------------------------------------------ |
| iban (required) | The IBAN. The IBAN can be a string, or a cell reference like A2. |

**Examples**
<img src="Images/Openiban.md - QBIC Examples.png" width="100%" height="100%" />



### QIBAN

The QIBAN function uses the [Openiban](https://openiban.com/) REST Webservice to return the IBAN (International Bank Account Number) for given bank code and account number and the given country.

| Excel Formula                                     | Result                                                       |
| ------------------------------------------------- | ------------------------------------------------------------ |
| =FFE.OPENIBAN.QIBAN("DE","37040044","0532013000") | Returns the german (*DE*) IBAN for bank code *37040044* and account number *0532013000*. Supported countries are returned by function `QCOUNTRIES`. |

**Syntax**

QIBAN(country_code,bank_code,account_number)

| Argument Name             | Description                                                  |
| ------------------------- | ------------------------------------------------------------ |
| country_code (required)   | The country code (ISO 3166-1 alpha-2 code) for which the IBAN should be calculated. |
| bank_code (required)      | The bank code for which the IBAN should be calculated.       |
| account_number (required) | The account number for which the IBAN should be calculated.  |

**Examples**
<img src="Images/Openiban.md - QIBAN Examples.png" width="100%" height="100%" />



### QCOUNTRIES

The QCOUNTRIES function uses the [Openiban](https://openiban.com/) REST Webservice to return an array of country names and codes supported by the function `QIBAN`.

| Excel Formula               | Result                                                       |
| --------------------------- | ------------------------------------------------------------ |
| =FFE.OPENIBAN.QCOUNTRIES()) | Returns an array of country names and codes supported by the function `QIBAN`. |

**Syntax**

QCOUNTRIES()

**Examples**
<img src="Images/Openiban.md - QCOUNTRIES Examples.png" width="100%" height="100%" />



## Examples

The examples shown above can be downloaded <a href="Attachments/Openiban Examples.xlsx">here</a>.
