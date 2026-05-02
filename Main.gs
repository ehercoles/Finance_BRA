let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {

  SpreadsheetApp.getUi()
      .createMenu('*Order')
      .addItem('Set', 'setOrders')
      .addItem('Clear Prices', 'clearPrices')
      .addItem('Clear', 'clearOrders')
      .addSeparator()
      .addItem('Fill', 'fillOrders')
      .addToUi();
}

function addOrder(orderType) {

  let portfolioSheet = spreadsheet.getSheetByName("Portfolio");
  let orderRange = spreadsheet.getRangeByName(orderType);
  let positionRange = spreadsheet.getRangeByName('Position');
  let orderData = [];
  const isBuy = orderType == 'Buy';
  const numRows = orderRange.getNumRows();
  
  //#region Set order
  for (let i=1; i<=numRows; i++) {

    let orderSymbol = positionRange.getCell(i, 1).getValue();
    let orderQty = parseInt(orderRange.getCell(i, 1).getValue());
    let orderPrice = parseFloat(orderRange.getCell(i, 2).getValue());
    
    if (orderQty > 0 && orderPrice > 0) {

      let qty = parseInt(0 + positionRange.getCell(i, 2).getValue());
      let avgCost = parseFloat(0 + positionRange.getCell(i, 3).getValue());
      
      if (isBuy) {
        
        if (qty == 0) {

          avgCost = orderPrice;
        }

        var newQty = qty + orderQty;
        var newAvgCost = ((qty * avgCost) + (orderQty * orderPrice)) / (qty + orderQty);

      } else {

        var newQty = qty - orderQty;
        var newAvgCost = avgCost;
      }
      
      orderData.push([
        new Date(),
        orderSymbol,
        qty,
        avgCost,
        orderQty,
        orderPrice, // Sheet "Sell" limit
        newQty,
        newAvgCost, // Sheet "Buy" limit
        i+2]); // Order index
    }
  }

  if (orderData.length == 0) return;
  //#endregion

  //#region Set position
  let numCol = orderData[0].length;

  for (let i=0; i<orderData.length; i++) {
    
    let orderRow = orderData[i];
    let orderIndex = orderRow[numCol-1];
    let qty = orderRow[6];
    let avgCost = orderRow[7];
    let values = [[qty, avgCost]];
    let positionRow = portfolioSheet.getRange(orderIndex, 2, 1, 2);

    if (qty == 0) {

      positionRow.setValue("");

    } else {
      
      positionRow.setValues(values);
    }
  }
  //#endregion
  
  //#region Add order
  if (isBuy) {
    
    orderData = Util.sliceColumn(orderData, 0, -1);
    
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
  //#endregion
}

function clearOrders() {

  spreadsheet.getRangeByName('Sell').setValue('');
  spreadsheet.getRangeByName('Buy').setValue('');
}

function clearPrices() {

  spreadsheet.getRangeByName('SellPrice').setValue('');
  spreadsheet.getRangeByName('BuyPrice').setValue('');
}

function setOrders() {

  try {

    let targetQuantityRange = spreadsheet.getRangeByName('TargetQuantity');
    let priceRange = spreadsheet.getRangeByName('Price');
    let buyRange = spreadsheet.getRangeByName('Buy');
    let sellRange = spreadsheet.getRangeByName('Sell');
    const numRows = targetQuantityRange.getNumRows();
    const buyOrders = new Array(numRows).fill([undefined,undefined]);
    const sellOrders = new Array(numRows).fill([undefined,undefined]);

    for (let i=1; i<=numRows; i++) {

      let targetQuantityCell = targetQuantityRange.getCell(i, 1);
      let qty = targetQuantityCell.getValue();
      let price = priceRange.getCell(i,1).getValue();
      
      if (targetQuantityCell.getBackgroundColor() == '#ff9900') { // orange

        sellOrders[i-1] = [qty*-1, price];

      } else if (targetQuantityCell.getBackgroundColor() == '#34a853') { // green

        buyOrders[i-1] = [qty, price];
      }
    }

    buyRange.setValues(buyOrders);
    sellRange.setValues(sellOrders);

  } catch (err) {

    Util.logError(err.stack);
  }
}

//#region BRA version: do not replace nor replicate the code below
function setBalance(dayTotal) {

  let orderTotal = spreadsheet.getRangeByName('OrderTotal');
  let cash = spreadsheet.getRangeByName('Cash');
  
  dayTotal.setFormula(dayTotal.getValue() + ' + ' + orderTotal.getValue());
  cash.setValue('');
}

function fillOrders() {
  
  try {

    let dayTotal = spreadsheet.getRangeByName('DayTotal');

    if ((dayTotal.getValue() != 0) && (!Util.confirm('Day total is not empty'))) return;

    //SpreadsheetApp.getUi().alert('TEST');

    addOrder('Buy');
    addOrder('Sell');
    setBalance(dayTotal);
    clearOrders();

  } catch (err) {
    
    Util.logError(err.stack);
  }
}
//#endregion
