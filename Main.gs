var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function addOrder(orderType) {

  var orderRange = spreadsheet.getRangeByName(orderType);
  var positions = spreadsheet.getRangeByName('Position');
  var orderData = [];
  var isBuy = orderType == 'Buy';
  const numRows = orderRange.getNumRows();
  
  for (let i = 1; i <= numRows; i++) {

    let orderQty = parseInt(orderRange.getCell(i, 1).getValue());
    let orderAp = parseFloat(orderRange.getCell(i, 2).getValue());
    
    if (orderQty > 0 && orderAp > 0) {
      let qty = parseInt(0 + positions.getCell(i, 2).getValue());
      let ap = parseFloat(0 + positions.getCell(i, 3).getValue());
      
      if (ap == 0) {
        ap = orderAp;
      }
      
      if (isBuy) {
        var newQty = qty + orderQty;
        var newAp = ((qty * ap) + (orderQty * orderAp)) / (qty + orderQty);

      } else {
        var newQty = qty - orderQty;
        var newAp = ap;
      }
      
      orderData.push([
        new Date(),
        positions.getCell(i, 1).getValue(),
        qty,
        ap,
        orderQty,
        orderAp, // Sheet "Sell" limit
        newAp, // Sheet "Buy" limit
        newQty,
        i]); // order index
    }
  }
  
  if (orderData.length == 0) return;

  //#region Set position
  let numCol = orderData[0].length;

  for (let i = 0; i < orderData.length; i++) {
    
    let orderIndex = orderData[i][numCol-1];

    positions.getCell(orderIndex, 2).setValue(orderData[i][numCol-2]); // Qty
    positions.getCell(orderIndex, 3).setValue(orderData[i][numCol-3]); // AP
  }
  //#endregion
  
  //#region Add order
  if (isBuy) {
    orderData = Util.sliceColumn(orderData, 0, -2);
    
  } else {
    orderData = Util.sliceColumn(orderData, 0, -3);
  }

  orderData.sort(Util.sortFunction);
  //Logger.log(orderData);

  let orderSheet = spreadsheet.getSheetByName(orderType);
  const rowStart = 3;
  const rowCount = orderData.length;
  
  numCol = orderData[0].length;
  orderSheet.insertRowsAfter(rowStart-1, rowCount);
  orderSheet.getRange(rowStart, 1, rowCount, numCol).setValues(orderData);
  
  // Copy formula to the new cells
  if (!isBuy) {
    const colStart = numCol + 1;
    const colCount = 3;

    let fromRange = orderSheet.getRange(rowStart-1, colStart, 1, colCount);
    fromRange.copyTo(orderSheet.getRange(rowStart, colStart, rowCount, colCount), {contentsOnly:false});
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
    Util.logError(err.stack);
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
    Util.logError(err.stack);
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

    if ((dayTotal.getValue() != 0) && (!Util.confirm('Day Total is not empty'))) return;

    //SpreadsheetApp.getUi().alert('TEST');

    addOrder('Buy');
    addOrder('Sell');
    setBalance(dayTotal);
    clearOrders();

  } catch (err) {
    Util.logError(err.stack);
  }
}

function incrementThreshold() {

    const range = spreadsheet.getRangeByName('Threshold');
    const value = range.getValue();
    const rule = range.getDataValidation();
    
    if (rule == null) return;

    //const criteria = rule.getCriteriaType();
    const args = rule.getCriteriaValues();
    const validationValues = args[0].getValues().filter(Number);
    const maxValue = validationValues[validationValues.length - 1];

    //Logger.log(validationValues);

    if (value < maxValue) {
      range.setValue(value + 1);
    }
}

function decrementThreshold() {
  
    const range = spreadsheet.getRangeByName('Threshold');
    const value = range.getValue();

    if (value > 1) {
      range.setValue(value - 1);
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
