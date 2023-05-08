var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var portfolioSheet = spreadsheet.getSheetByName('Portfolio');

function addSell(rowData) {
  var sellSheet = spreadsheet.getSheetByName('Sell');
  var row = 2;
  var lastRow = rowData.length + row;
  
  // Remove last column from array (Sell Index)
  var rowData = rowData.map(
    function(val) {
      return val.slice(0, -1);
    });
  
  sellSheet.insertRowsAfter(row, rowData.length);
  sellSheet.getRange('A3:F' + lastRow).setValues(rowData);
  
  var fromRange = sellSheet.getRange('G2:I2');
  fromRange.copyTo(sellSheet.getRange('G3:I' + lastRow), {contentsOnly:false});
}

function addBuy(rowData) {
  var buySheet = spreadsheet.getSheetByName('Buy');
  var lastRow = rowData.length + 1;
  
  // Remove last column from array (Buy Index)
  var rowData = rowData.map(
    function(val) {
      return val.slice(0, -1);
    });
  
  buySheet.insertRowsBefore(2, rowData.length);
  buySheet.getRange('A2:G' + lastRow).setValues(rowData);
}

function sell() {
  var sellRange = portfolioSheet.getRange('Sell');
  var positions = portfolioSheet.getRange('Position');
  const numRows = sellRange.getNumRows();
  var sellData = [];
  
  for (var i = 1; i <= numRows; i++) {
    var sellQty = sellRange.getCell(i, 1).getValue();
    var sellAp = sellRange.getCell(i, 2).getValue();
    
    if (sellQty > 0 && sellAp > 0) {
      /*
        0: Date
        1: Symbol
        2: Qty
        3: AP
        4: Sell Qty
        5: Sell AP
        6: Sell Index
      */
      sellData.push([
        new Date(),
        positions.getCell(i, 1).getValue(),
        positions.getCell(i, 2).getValue(),
        positions.getCell(i, 3).getValue(),
        sellQty,
        sellAp,
        i
      ]);
    }
  }
  
  if (sellData.length > 0) {
    sellData.sort(sortFunction);
    //console.log(sellData);
    
    addSell(sellData);
    
    for (var i = 0; i < sellData.length; i++) {
      var qty = parseInt(sellData[i][2]);
      var sellQty = parseInt(sellData[i][4]);
      var newQty = qty - sellQty;
      var sellIndex = sellData[i][6];
      
      positions.getCell(sellIndex, 2).setValue(newQty);
    }
  }
}

function buy() {
  var buyRange = portfolioSheet.getRange('Buy');
  var positions = portfolioSheet.getRange('Position');
  const numRows = buyRange.getNumRows();
  var buyData = [];
  
  for (var i = 1; i <= numRows; i++) {
    var buyQty = parseInt(buyRange.getCell(i, 1).getValue());
    var buyAp = parseFloat(buyRange.getCell(i, 2).getValue());
    
    if (buyQty > 0 && buyAp > 0) {
      var qty = parseInt(0 + positions.getCell(i, 2).getValue());
      var ap = parseFloat(0 + positions.getCell(i, 3).getValue());
      
      if (ap == 0) {
        ap = buyAp;
      }
      
      var newAp = ((qty * ap) + (buyQty * buyAp)) / (qty + buyQty);
      
      /*
        0: Date
        1: Symbol
        2: Qty
        3: AP
        4: Buy qty
        5: Buy AP
        6: New AP
        7: Buy Index
      */
      buyData.push([
        new Date(),
        positions.getCell(i, 1).getValue(),
        qty,
        ap,
        buyQty,
        buyAp,
        newAp,
        i
      ]);
    }
  }
  
  if (buyData.length > 0) {
    buyData.sort(sortFunction);
    //console.log(buyData);
    
    addBuy(buyData);
    
    for (var i = 0; i < buyData.length; i++) {
      var qty = buyData[i][2];
      var buyQty = buyData[i][4];
      var newAp = buyData[i][6];
      var newQty = qty + buyQty;
      var buyIndex = buyData[i][7];
      
      positions.getCell(buyIndex, 2).setValue(newQty);
      positions.getCell(buyIndex, 3).setValue(newAp);
    }
  }
}

function clearTrades() {
  portfolioSheet.getRange('Sell').setValue('');
  portfolioSheet.getRange('Buy').setValue('');
}

function clearPrices() {
  portfolioSheet.getRange('SellPrice').setValue('');
  portfolioSheet.getRange('BuyPrice').setValue('');
}

function setTrades(mode) {
  try {
    var targetQuantities = portfolioSheet.getRange('TargetQuantity');
    var prices = portfolioSheet.getRange('Price');
    var sellRange = portfolioSheet.getRange('Sell');
    var buyRange = portfolioSheet.getRange('Buy');
    var tradeCompensation = portfolioSheet.getRange('TradeCompensation');
    const numRows = targetQuantities.getNumRows();
    
    tradeCompensation.setValue('');

    for (var i2 = 0; i2 < 2; i2++) {
      for (var i = 1; i <= numRows; i++) {
        var targetQuantityCell = targetQuantities.getCell(i, 1);
        var qty = targetQuantityCell.getValue();
        var price = prices.getCell(i, 1).getValue();
        
        if ((!mode || mode == 'sell') && targetQuantityCell.getBackgroundColor() == '#ff9900') { // orange
          var quantityCell = sellRange.getCell(i, 1);
          var priceCell = sellRange.getCell(i, 2);

          quantityCell.setValue(qty * -1);
          priceCell.setValue(price);
        }
        else if ((!mode || mode == 'buy') && targetQuantityCell.getBackgroundColor() == '#34a853') { // green
          var quantityCell = buyRange.getCell(i, 1);
          var priceCell = buyRange.getCell(i, 2);

          quantityCell.setValue(qty);
          priceCell.setValue(price);
        }
      }

      // Compensation
      var tradeTotal = parseFloat(portfolioSheet.getRange('TradeTotal').getValue());
      var cash = parseFloat(0 + portfolioSheet.getRange('Cash').getValue());

      tradeCompensation.setValue(tradeTotal + cash);
    }

    tradeCompensation.setValue('');
    
  } catch (err) {
    logError(err.stack);
  }
}

function setPrices() {
  try {
    var prices = portfolioSheet.getRange('Price');
    var sellRange = portfolioSheet.getRange('Sell');
    var buyRange = portfolioSheet.getRange('Buy');
    const numRows = prices.getNumRows();
    
    // Sell range
    for (var i = 1; i <= numRows; i++) {
      var quantityCell = sellRange.getCell(i, 1);
      var qty = quantityCell.getValue();
      
      if (qty > 0) {
        var price = prices.getCell(i, 1).getValue();
        var priceCell = sellRange.getCell(i, 2);
        
        priceCell.setValue(price);
      }
    }

    // Buy range
    for (var i = 1; i <= numRows; i++) {
      var quantityCell = buyRange.getCell(i, 1);
      var qty = quantityCell.getValue();
      
      if (qty > 0) {
        var price = prices.getCell(i, 1).getValue();
        var priceCell = buyRange.getCell(i, 2);
        
        priceCell.setValue(price);
      }
    }
  } catch (err) {
    logError(err.stack);
  }
}

function setSell() {
  setTrades('sell');
}

function setBuy() {
  setTrades('buy');
}

function setBalance() {
  var tradeTotalCell = portfolioSheet.getRange('TradeTotal');
  var dayTotalCell = portfolioSheet.getRange('DayTotal');
  var cashCell = portfolioSheet.getRange('Cash');
  
  dayTotalCell.setFormula(dayTotalCell.getValue() + ' + ' + tradeTotalCell.getValue());
  cashCell.setValue('');
}

function endTrades() {
  try {
    //var lock = LockService.getScriptLock();
    //lock.waitLock(20000);

    var dayTotalCell = portfolioSheet.getRange('DayTotal');
    var cashCell = portfolioSheet.getRange('Cash');

    if (dayTotalCell.getValue() != 0) {
      if (!confirm('Day Total is not empty')) {
        //SpreadsheetApp.getUi().alert('exit');
        return;
      }
    }

    if (cashCell.getValue() != 0) {
      if (!confirm('Cash is not empty')) {
        //SpreadsheetApp.getUi().alert('exit');
        return;
      }
    }

    sell();
    buy();
    setBalance();
    clearTrades();
    
    //lock.releaseLock();
  } catch (err) {
    logError(err.stack);
  }
}

function logError(message) {
  MailApp.sendEmail('ehercoles@gmail.com', 'GAS error', message);
}

function confirm(message) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     message,
     'Are you sure you want to continue?',
      ui.ButtonSet.YES_NO);
  return result == ui.Button.YES;
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Trade')
      .addItem('Set', 'setTrades')
      .addItem('Set Sell', 'setSell')
      .addItem('Set Buy', 'setBuy')
      .addItem('Set Prices', 'setPrices')
      .addItem('Clear Prices', 'clearPrices')
      .addItem('Clear', 'clearTrades')
      .addItem('End', 'endTrades')
      .addToUi();
}
