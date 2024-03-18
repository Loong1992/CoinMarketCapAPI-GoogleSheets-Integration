# CoinMarketCapAPI-GoogleSheets-Integration
Effortlessly track cryptocurrency prices in Google Sheets with our simple script, using CoinMarketCap's API. Update prices with one click, costing just 1 API credit each time. The Basic plan offers 10,000 credits per month, ensuring you stay within limits and keep your data up-to-date. Say goodbye to API limits and hello to hassle-free tracking.

### Usage
#### The first parameter is accepting either A1 notation or a cryptocurrency symbol. The second parameter specifies the category you wish to query. For example:
##### Get BTC name
```
=getCryptoDataBySymbol(A1, "name")
```
```
=getCryptoDataBySymbol("BTC", "name")
```
##### Get BTC rank
```
=getCryptoDataBySymbol(A1, "cmc_rank")
```
```
=getCryptoDataBySymbol("BTC", "cmc_rank")
```
##### Get BTC price (USD)
```
=getCryptoDataBySymbol(A1, "quote.USD.price")
```
```
=getCryptoDataBySymbol("BTC, "quote.USD.price")
```

### Acknowledgment
I would like to extend my gratitude to [@dmarcosl]( https://github.com/dmarcosl/coinmarketcap-google-sheet-script ) for generously sharing their open-source code, which served as the foundational inspiration for this project.
