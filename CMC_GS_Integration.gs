// Global variables
var API_KEY = 'YOUR_API_KEY'; // Replace with your API key
var options = {
  "method": "GET",
  "contentType": "application/json",
  "headers": {
    "Accept": "application/json",
    "X-CMC_PRO_API_KEY": API_KEY
  },
  "muteHttpExceptions": true,
  "validateHttpsCertificates": true
};

// Function to fetch cryptocurrency data
function fetchCryptoData(symbols) {
  try {
    var url = "https://pro-api.coinmarketcap.com/v2/cryptocurrency/quotes/latest?convert=USD" + "&symbol=" + symbols.toString();
    var response = UrlFetchApp.fetch(url, options);
    var responseJSON = JSON.parse(response.getContentText());
    var datas = [];
    if (responseJSON.data) {
      symbols.forEach(symbol => datas.push(responseJSON.data[symbol][0]));
    } else {
      console.error('Error: Unexpected API response format');
      return null;
    }
    return datas;
  } catch (error) {
    console.error('Error fetching cryptocurrency data:', error);
    return null;
  }
}

// Function to update cryptocurrency data in spreadsheet
function updateCryptosData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getDataRange();
  const formulas = range.getFormulas();
  const num_cols = range.getNumColumns();
  const num_rows = range.getNumRows();
  let symbol_positions = {};
  let symbols = [];
  let category_positions = {};
  let categories = [];

  // Loop through spreadsheet to find formulas
  for (let row = 0; row < num_rows; row++) {
    for (let col = 0; col < num_cols; col++) {
      var formula = formulas[row][col];
      if (formula && formula.indexOf("getCryptoDataBySymbol") > -1) {
        let value = formula.slice(formula.indexOf("getCryptoDataBySymbol"));
        let valueSplit = value.split(",");
        let symbolOrA1Notation = valueSplit[0].split("(")[1].trim().replace(/"/g, "");
        var a1NotationPattern = /^[A-Za-z]+\d+$/;
        if (a1NotationPattern.test(symbolOrA1Notation)) {
          symbol = sheet.getRange(symbolOrA1Notation).getValue();
        } else {
          symbol = symbolOrA1Notation;
        }
        let category = valueSplit[1].split(")")[0].trim().replace(/"/g, "");

        if (!symbol_positions[symbol]) {
          symbol_positions[symbol] = [];
        }
        symbol_positions[symbol].push([row, col]);
        if (!symbols.includes(symbol)) {
          symbols.push(symbol);
        }

        if (!category_positions[category]) {
          category_positions[category] = [];
        }
        category_positions[category].push([row, col]);
        if (!categories.includes(category)) {
          categories.push(category);
        }
      }
    }
  }

  // Fetch cryptocurrency data
  const datas = fetchCryptoData(symbols);

	// Update spreadsheet cells with fetched data
	Object.keys(symbol_positions).forEach(function(symbol) {
	  var position_list = symbol_positions[symbol];
	  position_list.forEach(function(position) {
		const [row, col] = position;
		const index = symbols.indexOf(symbol);
		const cell = range.getCell(row + 1, col + 1);
		let result;
		try {
		  result = categories[index].split('.').reduce(function(obj, prop) {
			return obj[prop];
		  }, datas[index]);
		} catch (error) {
		  console.error('Error accessing property: ' + categories[index]);
		}
		const value = typeof result === 'string' ? '"' + result + '"' : result;
		cell.setFormula(`=getCryptoDataBySymbol("${symbol}", "${categories[index]}", ${value})`);
	  });
	});
}

// Function to get cryptocurrency data by symbol and category
function getCryptoDataBySymbol(symbol, category, value) {
  if (!value) {
    value = "Please use the menu option to retrieve the price.";
  }
  return value;
}

// Function to create custom menu
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("CoinMarketCap")
    .addItem("Update", "updateCryptosData")
    .addToUi();
}