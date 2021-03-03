// Copyright 2021, www.allstacksdeveloper.com, All rights reserved.
/**
 * Add tab to the menu bar
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive()
  var menuItems = [
    { name: 'Generate Historical Prices Sheets', functionName: 'generateHistoricalPricesSheets' },
    { name: 'Generate Daily Evolution of Portfolio', functionName: 'generateDailyEvolution' },
    { name: 'Delete Historical Prices Sheets', functionName: 'deleteHistoricalPricesSheets' }
  ]
  spreadsheet.addMenu('Portfolio Tools', menuItems)
}

/**
 * Generate a sheet for each unique symbol found in the Transactions tab.
 * Each sheet contains historical prices of the symbol until today.
 */
function generateHistoricalPricesSheets() {
  var symbols = extractSymbolsFromTransactions()
  var firstTransactionDate = extractFirstTransactionDate()
  symbols.forEach(symbol => generateHistoricalPriceSheetForSymbol(symbol, firstTransactionDate))
}

/**
 * Delete all sheets for all symbols found in the Transactions tab.
 */
function deleteHistoricalPricesSheets() {
  var symbols = extractSymbolsFromTransactions()
  symbols.forEach(symbol => deleteHistoricalPriceSheetForSymbol(symbol))
}

/**
 * From the Transactions sheet and all historical prices sheets for all symbols:
 * - Compute daily evolution of the portfolio from the first transaction date
 * - Write evolutions into Evolutions sheet
 * 
 * 'Date', 'Invested Money', 'Cash', 'Market Value', 'Portfolio Value', 'Gain', 'Gain Percentage'
 */
function generateDailyEvolution() {
  var transactions = extractTransactions()
  var indexes = extractIndexes()

  var portfolioByTransactionDate = computePortfolioByTransactionDate(transactions)

  var historicalPricesBySymbol = {}
  var symbols = extractSymbolsFromTransactions()
  symbols.forEach(symbol => {
    historicalPricesBySymbol[symbol] = getHistoricalPricesBySymbol(symbol)
  })
  indexes.forEach(index => {
    historicalPricesBySymbol[index.symbol] = getHistoricalPricesBySymbol(index.symbol)
  })

  var firstTransactionDate = transactions[0].date

  // Compute Evolutions

  var evolutionHeaders = ['Date', 'Invested Money', 'Cash', 'Market Value', 'Portfolio Value', 'Gain', 'Gain Percentage']
  indexes.forEach(index => {
    evolutionHeaders.push(index.name)
  })
  var evolutions = [evolutionHeaders]
  var portfolioSnapshot
  for (var aDate = firstTransactionDate; aDate <= new Date(); aDate.setDate(aDate.getDate() + 1)) {
    var dString = getDateString(aDate)
    var invested = 0
    var cash = 0
    var value = 0
    portfolioSnapshot = portfolioByTransactionDate[dString] ? portfolioByTransactionDate[dString] : portfolioSnapshot
    if (portfolioSnapshot) {
      for (const key in portfolioSnapshot) {
        switch (key) {
          case 'cash':
            cash = portfolioSnapshot.cash
            break
          case 'invested':
            invested = portfolioSnapshot.invested
            break
          default:
            var symbol = key
            var numShares = portfolioSnapshot[symbol]
            if (numShares > 0) {
              var priceOfSymbolOnDate = historicalPricesBySymbol[symbol][dString]
              if (priceOfSymbolOnDate) {
                value += numShares * priceOfSymbolOnDate
              } else {
                value = -1
                break
              }
            }
            break
        }
      }
    }
    if (value > -1) {
      var portfolioValue = value + cash
      var gain = portfolioValue - invested
      var gainPercentage = gain / invested
      var evolution = [new Date(aDate.getTime()), invested, cash, value, portfolioValue, gain, gainPercentage]
      var lineNumber = evolutions.length + 1
      indexes.forEach(index => {
        evolution.push(historicalPricesBySymbol[index.symbol][dString])
      })
      evolutions.push(evolution)
    }
  }

  // Write the evolutions
  var spreadsheet = SpreadsheetApp.getActive()
  var sheetName = 'Evolutions'
  var sheet = spreadsheet.getSheetByName(sheetName)
  if (sheet) {
    sheet.clear()
    sheet.activate()
  } else {
    sheet = spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets())
  }

  // evolutions[0] is the header row
  sheet.getRange(1, 1, evolutions.length, evolutions[0].length).setValues(evolutions)
}

/**
 * Compute the composition of portfolio on each day of transaction.
 * A composition of portfolio contains:
 * - Amount of invested money so far (Deposit - Withdrawal)
 * - Amount of available cash
 * - Number of shares for each bought stock
 * 
 * {
 *  invested: 10000,
 *  cash: 2001.42,
 *  APPL: 400,
 *  GOOGL: 500
 * }
 * @param {Array} transactions 
 */
function computePortfolioByTransactionDate(transactions) {
  var portfolioByDate = {}
  var portfolioSnapshot = {
    invested: 0,
    cash: 0
  }
  for (var i = 0; i < transactions.length; i++) {
    var transaction = transactions[i]
    var tDate = getDateString(transaction.date)
    var tType = transaction.type
    var tSymbol = transaction.symbol
    var tAmount = transaction.amount
    var tShares = transaction.shares
    if (tType === 'BUY' || tType === 'SELL') {
      if (!portfolioSnapshot.hasOwnProperty(tSymbol)) {
        portfolioSnapshot[tSymbol] = 0
      }
      portfolioSnapshot[tSymbol] += Number(tShares)
    }
    if (tType === 'DEPOSIT' || tType === 'WITHDRAWAL') {
      portfolioSnapshot.invested += Number(tAmount)
    }
    portfolioSnapshot.cash += Number(tAmount)
    var portfolioCloned = {}
    Object.assign(portfolioCloned, portfolioSnapshot)
    portfolioByDate[tDate] = portfolioCloned
  }
  return portfolioByDate
}

/**
 * Get the first transaction date from the sheet Transactions.
 */
function extractFirstTransactionDate() {
  var spreadsheet = SpreadsheetApp.getActive()
  var transactionsSheet = spreadsheet.getSheetByName('Transactions')
  var rows = transactionsSheet.getRange('A:E').getValues()

  var transactions = []
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i]
    if (row[0].length === 0) {
      break
    }

    transactions.push({
      date: row[0]
    })
  }

  // Order by date ascending
  transactions.sort((t1, t2) => t1.date < t2.date ? -1 : 1)

  console.log(transactions[0].date)
  return transactions[0].date
}

/**
 * Extract and return array of transactions ordered by date ascending from the sheet Transactions.
 */
function extractTransactions() {
  var spreadsheet = SpreadsheetApp.getActive()
  var transactionsSheet = spreadsheet.getSheetByName('Transactions')
  var rows = transactionsSheet.getRange('A:E').getValues()

  var transactions = []
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i]
    if (row[0].length === 0) {
      break
    }

    transactions.push({
      date: row[0],
      type: row[1],
      symbol: row[2],
      amount: row[3],
      shares: row[4]
    })
  }

  // Order by date ascending
  transactions.sort((t1, t2) => t1.date < t2.date ? -1 : 1)

  return transactions
}

/**
 * Extract unique symbols from the Transactions sheets
 */
function extractSymbolsFromTransactions() {
  var spreadsheet = SpreadsheetApp.getActive()
  var transactionsSheet = spreadsheet.getSheetByName('Transactions')
  var rows = transactionsSheet.getRange('C2:C').getValues()

  var symbols = new Set() // Use Set to avoid duplicates
  for (var i = 0; i < rows.length; i++) {
    var symbol = rows[i][0]
    if (symbol.length > 0) {
      symbols.add(symbol)
    }
  }

  // Convert from Set to array
  return [...symbols]
}

/**
 * Create a new sheet whose name is name of the symbol
 * for its historical prices until today.
 * 
 * The formula below is added to A1 cell.
 * GOOGLEFINANCE("SYMBOL", "price", "1/1/2014", TODAY(), "DAILY")
 * @param {String} symbol 
 */
function generateHistoricalPriceSheetForSymbol(symbol, fromDate) {
  // Create a new empty sheet for the symbol
  var spreadsheet = SpreadsheetApp.getActive()
  var sheetName = symbol
  var symbolSheet = spreadsheet.getSheetByName(sheetName)
  if (symbolSheet) {
    symbolSheet.clear()
    symbolSheet.activate()
  } else {
    symbolSheet = spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets())
  }

  let fromDateFunction = "DATE(" + fromDate.getFullYear() + "," + (fromDate.getMonth() + 1) + "," + fromDate.getDate() + ")"
  var historicalPricesFormula = 'GOOGLEFINANCE("' + symbol + '", "price", ' + fromDateFunction + ', TODAY(), "DAILY")'
  symbolSheet.getRange('A1').setFormula(historicalPricesFormula)

  symbolSheet.getRange('A:A').setNumberFormat('dd/mm/yyyy')
  symbolSheet.getRange('B:B').setNumberFormat('#.###')
}

/**
 * Delete the sheet whose name is name of the symbol
 * @param {String} symbol 
 */
function deleteHistoricalPriceSheetForSymbol(symbol) {
  var spreadsheet = SpreadsheetApp.getActive()
  var sheetName = symbol
  var symbolSheet = spreadsheet.getSheetByName(sheetName)
  if (symbolSheet) {
    spreadsheet.deleteSheet(symbolSheet)
  }
}

/**
 * Each symbol has its own sheet of its name for its historical prices.
 * 
 * Return a map from date to close price on that date.
 * {
 *  '2020-01-31': 22.30,
 *  '2020-02-01': 21.54
 * }
 * @param {String} symbol 
 */
function getHistoricalPricesBySymbol(symbol) {
  var spreadsheet = SpreadsheetApp.getActive()
  var historySheet = spreadsheet.getSheetByName(symbol)
  var rows = historySheet.getRange('A:B').getValues()

  var priceByDate = {}
  for (var i = 1; i < rows.length; i++) { // Start from 1 to ignore headers
    var tDate = rows[i][0]
    if (tDate) {
      tDate = getDateString(rows[i][0])
      var close = rows[i][1]
      priceByDate[tDate] = close
    } else {
      break // it means empty row.
    }
  }
  return priceByDate
}

/**
 * Extract name and symbol of indexes from the sheet Indexes.
 */
function extractIndexes() {
  var spreadsheet = SpreadsheetApp.getActive()
  var sheet = spreadsheet.getSheetByName('Indexes')
  var rows = sheet.getRange('A:B').getValues()

  var indexes = []
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i]
    if (row[0].length === 0) {
      break
    }

    indexes.push({
      name: row[0],
      symbol: row[1]
    })
  }

  return indexes
}

/**
 * Generate a historical prices sheet for each index
 */
function generateHistoricalPricesForIndexes() {
  var indexes = extractIndexes()
  var firstTransactionDate = extractFirstTransactionDate()
  indexes.forEach(index => generateHistoricalPriceSheetForSymbol(index.symbol, firstTransactionDate))
}

function getDateString(aDate) {
  return aDate.getFullYear() + '-' + (aDate.getMonth() + 1) + '-' + aDate.getDate()
}
