# market_track

Google App script for Google Sheets to track exchange rates, international stocks and funds


## Purpose

Enable one to safely track a large numbre of stocks, exchange rates and funds over decades.

This [Google Apps Script](https://developers.google.com/apps-script/) for [Google Sheets](https://developers.google.com/apps-script/guides/sheets) is effectively making a cache of the historical
values of stocks you want to track. This permits analysing and graphing without having
to constantly use [GOOGLEFINANCE()](https://support.google.com/docs/answer/3093281) function, which eventually becomes throttled by Google
Apps when you start tracking lot of stocks. This is a problem when you track a hundreds
of stocks over decades.	

	
## Why

Your favorite app will likely not be supported in 10 years, neither your favorite web site
maintained by a startup with no income. On the other hand, spreadsheets have a 99% chance
of being around in 2040.

Static data can be printed on dead trees (paper) and analysed 10 years later, so it's
better to retrieve the data and keep a static copy than rely on the functions to retrieve
the data on the fly. For example, many finance functions do not want to go back more than
10 years ago, which can be a real problem. You can import a CSV and mix
with the sheets here. Even if the code here doesn't work in 2040, **your data is safe** and
you can export it to CSV to import it back into _Microsoft Office 2038_.


## Actions


### Setup

1. Create a [new Google Sheet](https://docs.google.com/spreadsheets/create).
2. Select menu `Tools`, `Scripts editor`.
3. Paste the content of `code.gs` into the editor.
4. Save, close and reopen the sheet.
5. Select menu `Stocks`, `Add new sheet`. If you can't find it, it'll be on the far right right after `Help`. If it's not there, the script didn't load.

To update to a new version, redo steps 2, 3 and 4 on an existing sheet, replacing the old code.


### Track a new stock, ETF or fund

Use the menu `Stocks`, `Add new sheet` or press the button on the right. Give it
the ticker symbol and the starting date to track the stock. For example, try `GOOG`
with date `2015` will start on the first day the stock market was open in 2015.


### Track currency exchange rate

Same as stock, use CURRENCY:FROMTO as one word. For example, USD to CAD is
`CURRENCY:USDCAD`.


### Update a sheet

Use the menu `Stocks`, `Update current sheet`.


### Update all sheets

Use the menu `Stocks`, `Update all sheets`.


### Start over with a sheet

Just delete the sheet and create it again.
