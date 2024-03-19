var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

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
  var sellRange = spreadsheet.getRangeByName('Sell');
  var positions = spreadsheet.getRangeByName('Position');
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
  var buyRange = spreadsheet.getRangeByName('Buy');
  var positions = spreadsheet.getRangeByName('Position');
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

function clearOrders() {
  spreadsheet.getRangeByName('Sell').setValue('');
  spreadsheet.getRangeByName('Buy').setValue('');
}

function clearPrices() {
  spreadsheet.getRangeByName('SellPrice').setValue('');
  spreadsheet.getRangeByName('BuyPrice').setValue('');
}

function setOrders(mode) {
  try {
    var targetQuantities = spreadsheet.getRangeByName('TargetQuantity');
    var prices = spreadsheet.getRangeByName('Price');
    var sellRange = spreadsheet.getRangeByName('Sell');
    var buyRange = spreadsheet.getRangeByName('Buy');
    const numRows = targetQuantities.getNumRows();
    
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
    
  } catch (err) {
    logError(err.stack);
  }
}

function setPrices() {
  try {
    var prices = spreadsheet.getRangeByName('Price');
    var sellRange = spreadsheet.getRangeByName('Sell');
    var buyRange = spreadsheet.getRangeByName('Buy');
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
  setOrders('sell');
}

function setBuy() {
  setOrders('buy');
}

function setBalance(dayTotal) {

  var orderTotal = spreadsheet.getRangeByName('OrderTotal');
  var cash = spreadsheet.getRangeByName('Cash');
  
  dayTotal.setFormula(dayTotal.getValue() + ' + ' + orderTotal.getValue());
  cash.setValue('');
}

function fillOrders() {
  
  try {
    var dayTotal = spreadsheet.getRangeByName('DayTotal');

    if ((dayTotal.getValue() != 0) && (!confirm('Day Total is not empty'))) {
      return;
    }

    buy();
    sell();
    setBalance(dayTotal);
    clearOrders();

  } catch (err) {
    logError(err.stack);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('*Order')
      .addItem('Set', 'setOrders')
      .addItem('Set Sell', 'setSell')
      .addItem('Set Buy', 'setBuy')
      .addItem('Set Prices', 'setPrices')
      .addItem('Clear Prices', 'clearPrices')
      .addItem('Clear', 'clearOrders')
      .addSeparator()
      .addItem('Fill', 'fillOrders')
      .addToUi();
}
