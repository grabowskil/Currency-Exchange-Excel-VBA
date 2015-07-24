Currency-Exchange-Excel-VBA
===========================

Pulls specified .json from yahoo finance to fetch exchange rates.

#### Install
To execute, implement following libraries for Excel VBA
 * Microsoft HTML Object Library
 * Microsoft Internet Controls
 * Microsoft_JScript

Import the .bas-file and you are ready to go

#### How to use
Inside of another script or in the worksheet use `GetExchange(Currency 1, Currency 2)` (use the [ISO 4217 names](http://en.wikipedia.org/wiki/ISO_4217#Active_codes "wikipedia.org")) to pull the exchange rate from yahoo's server. Make sure, your IE (8 or higher) is configured to display `.json`-files.
Because every new calculation requests a new pull I advise you to use this function (and refreshes) wisely, to keep ressource usage low, create a worksheet with the few exchange rates you need and don't use this function / formula excessively.

#### Configure Internet Explorer to disply .json-files
Create a new file anywhere you want and call it `IE-json.reg` (or use the `.reg`-file from this git).
```
Windows Registry Editor Version 5.00;
[HKEY_CLASSES_ROOT\MIME\Database\Content Type\application/json]
"CLSID"="{25336920-03F9-11cf-8FD0-00AA00686F13}"
"Encoding"=hex:08,00,00,00
```
credits: http://www.codeproject.com/Tips/216175/View-JSON-in-Internet-Explorer
