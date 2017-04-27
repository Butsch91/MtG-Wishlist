// **********************************************
// function fcnAddLine()
//
// 
//
// **********************************************

function fcnAddLine()
{
  // Load active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var LastColumn = actSht.getMaxColumns();
  var actRng = ss.getActiveRange();
  var actRow = actRng.getRowIndex();
  var NewRow = actRow + 1;
  
  var DfltTotal = '=IF(INDIRECT("R[0]C[2]",FALSE) = "Buy",INDIRECT("R[0]C[-4]",FALSE)*INDIRECT("R[0]C[-1]",FALSE),"")';
  var DfltSort = '=If(INDIRECT("R[0]C[1]",FALSE)<>"",switch(INDIRECT("R[0]C[1]",FALSE),Lists!$B$3,Lists!$A$3,Lists!$B$4,Lists!$A$4,Lists!$B$5,Lists!$A$5,Lists!$B$6,Lists!$A$6,Lists!$B$7,Lists!$A$7,Lists!$B$8,Lists!$A$8,Lists!$B$9,Lists!$A$9,Lists!$B$10,Lists!$A$10,Lists!$B$11,Lists!$A$11,Lists!$B$12,Lists!$A$12),"")';
  var DfltBuyStore = '=CONCATENATE(INDIRECT("R[0]C[-8]",FALSE)," ",INDIRECT("R[0]C[-7]",FALSE))';
  var DfltBuyOnline = '=CONCATENATE(INDIRECT("R[0]C[-9]",FALSE)," x ",INDIRECT("R[0]C[-8]",FALSE)," - ", INDIRECT("R[0]C[-7]",FALSE))';
  
  var DfltValues =    [[1,'','','$0.00',DfltTotal,DfltSort,'','',DfltBuyStore,DfltBuyOnline,'']];
  var DfltValuesBuy = [[1,'','','$0.00',DfltTotal,DfltSort,'Buy','',DfltBuyStore,DfltBuyOnline,'']];
  
  var CardRow = actSht.getRange(NewRow,1,1,11);
  var CardRowFonts = [['black','black','black','black','black','black','black','black','black','black','black']];
  var CardRowColors = [[null,null,null,null,null,null,null,null,null,null,null]];
  var CardRowWeight = [['normal','normal','normal','normal','normal','normal','normal','normal','normal','normal','normal']];
  var CardRowSize = [[10,10,10,10,10,10,10,10,10,10,10]];
  var CardRowAlign = [['center','left','left','center','center','center','left','left','left','left','left']];
  
  var CellTotalRef = actSht.getRange(actRow,5);

  var CellCostNew = actSht.getRange(NewRow,4);
  var CellTotalNew = actSht.getRange(NewRow,5);
  var CellOrderNb = actSht.getRange(NewRow,6);

  var CellConcatRef = actSht.getRange(actRow,8); 
  var CellConcatNew = actSht.getRange(NewRow,8); 
    
  if (actRow > 5){
    actSht.insertRowAfter(actRow);

    CardRow.setFontColors(CardRowFonts);
    CardRow.setBackgrounds(CardRowColors);
    CardRow.setFontWeights(CardRowWeight);
    CardRow.setFontSizes(CardRowSize);
    CardRow.setHorizontalAlignments(CardRowAlign);

    if(actShtName != 'Buy') CardRow.setValues(DfltValues);
    if(actShtName == 'Buy') CardRow.setValues(DfltValuesBuy);
    
    CellCostNew.setNumberFormat('$0.00');
    CellTotalNew.setNumberFormat('$0.00');
    CellOrderNb.setNumberFormat('0');
  }
}


// **********************************************
// function fcnFiveAddLine()
//
// 
//
// **********************************************

function fcnFiveAddLines()
{
  // Load active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actRng = ss.getActiveRange();
  var actRow = actRng.getRowIndex();
  
  if (actRow > 6){
    for(var Row = 0; Row < 5; Row++){
      fcnAddLine();
    }
  }
}

// **********************************************
// function fcnGathererLink()
//
// Creates a link to the Gatherer Card Page for
// the selected card
// 
// **********************************************

function fcnGathererLink() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var actCell = actSht.getActiveCell();
  var CardRow = 2;
  var Row = actCell.getRow();
  var Col = actCell.getColumn();
  var CardName;
  //var TestSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test");
  
  //TestSht.getRange(1,2).setValue(actShtName);
  //TestSht.getRange(2,2).setValue(CardName);
  
  if((actShtName == 'Buy' || actShtName == 'Sell' || actShtName == 'Wish') && Row == CardRow) { 
    CardName = actCell.getValue();
    if(CardName != '') actCell.setValue('=HYPERLINK("http://gatherer.wizards.com/Pages/Card/Details.aspx?name='+CardName+'","'+CardName+'")');
    actCell.setFontColor('black');
    actCell.setFontLine('none');
  }
}

// **********************************************
// function fcnGathererLinkAllCards()
//
// Creates a link to the Gatherer Card Page for
// each card in the list
// 
// **********************************************

function fcnGathererLinkAllCards() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxRow = actSht.getMaxRows();
  var Row = 5;
  var CardCol = 2;
  var CardName;
  var CardRng;
  //var TestSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test");
  
  //TestSht.getRange(1,2).setValue(actShtName);
  //TestSht.getRange(2,2).setValue(CardName);
  
  if(actShtName == 'Buy' || actShtName == 'Sell' || actShtName == 'Wish') { 
    for (Row; Row <= MaxRow; Row++){
      CardRng = actSht.getRange(Row, CardCol);
      CardName = CardRng.getValue();
      if(CardName != '') CardRng.setValue('=HYPERLINK("http://gatherer.wizards.com/Pages/Card/Details.aspx?name='+CardName+'","'+CardName+'")');
      CardRng.setFontColor('black');
      CardRng.setFontLine('none');
    }
  }
}


// **********************************************
// function fcnRemoveCardsInDeck()
//
// Finds each card with the "In Deck" status
// and deletes the row
// 
// **********************************************

function fcnRemoveCardsInDeck() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actRng = ss.getActiveRange();
  
  actRng.sort([{column: 6, ascending: true}, {column: 4, ascending: true}, {column: 3, ascending: true}]);
  
  // Select the Deck title and verify (or select the range prior to execute the function?)
  // Finds the next Deck Start (look for "Deck" status) and finds the appropriate range
  // Finds "In Deck" status for each card and deletes the row
}


// **********************************************
// function fcnTransferToBuyList()
//
// Transfers all cards with the "Buy" status to 
// the Buy Tab 
// 
// **********************************************

function fcnTransferToBuyList() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var WishSht = ss.getSheetByName('Wish');
  var WishMaxRows = WishSht.getMaxRows();
  var WishMaxCols = WishSht.getMaxColumns();
  var WishCardData;
  var WishCardName;
  var WishStatus;
  var WishColCard = 2;
  var WishColStatus = 7;
  var WishRowStart = 7;

  var BuySht = ss.getSheetByName('Buy');
  var BuyMaxRows = BuySht.getMaxRows();
  var BuyMaxCols = BuySht.getMaxColumns();
  var BuyCard;
  var BuyCardFound = 0;
  var BuyEmptyRow = 0;
  var BuyColCard = 2;
  var BuyColStatus = 7;
  var BuyRowStart = 7;
  var NumCol = 4;
  
  var DfltTotal = '=IF(INDIRECT("R[0]C[2]",FALSE) = "Buy",INDIRECT("R[0]C[-4]",FALSE)*INDIRECT("R[0]C[-1]",FALSE),"")';
  var DfltSort = '=If(INDIRECT("R[0]C[1]",FALSE)<>"",switch(INDIRECT("R[0]C[1]",FALSE),Lists!$B$3,Lists!$A$3,Lists!$B$4,Lists!$A$4,Lists!$B$5,Lists!$A$5,Lists!$B$6,Lists!$A$6,Lists!$B$7,Lists!$A$7,Lists!$B$8,Lists!$A$8,Lists!$B$9,Lists!$A$9,Lists!$B$10,Lists!$A$10,Lists!$B$11,Lists!$A$11,Lists!$B$12,Lists!$A$12),"")';
  var DfltBuyOnline = '=CONCATENATE(INDIRECT("R[0]C[-8]",FALSE)," ",INDIRECT("R[0]C[-7]",FALSE))';
  var DfltBuyStore = '=CONCATENATE(INDIRECT("R[0]C[-9]",FALSE)," x ",INDIRECT("R[0]C[-8]",FALSE)," - ", INDIRECT("R[0]C[-7]",FALSE))';
  var DfltValuesBuy = [[1,'','','$0.00',DfltTotal,DfltSort,'Buy','',DfltBuyOnline,DfltBuyStore,'']];
  
    var TestSht = ss.getSheetByName('Test');
  TestSht.clear();
  
  // Clears all Cards in Buy List
  BuySht.getRange(BuyRowStart,1,BuyMaxRows - BuyRowStart,4).setValue('');
  
  // Browse through all Cards in Wishlist
  for (var WishRow = WishRowStart; WishRow <= WishMaxRows; WishRow++){
        
    WishStatus = WishSht.getRange(WishRow,WishColStatus).getValue();
    // Finds a Card with the Status "Buy"
    if (WishStatus == 'Buy'){
      // Saves the Card Name from the Wish List
      WishCardName = WishSht.getRange(WishRow,2).getValue();
            
      // Initializes Data for Buy List Search
      BuyMaxRows = BuySht.getMaxRows();
      BuyEmptyRow = 0;
      BuyCardFound = 0;
      
      // Search if card from Wish List is already in Buy List 
      for (var BuyRow = BuyRowStart; BuyRow <= BuyMaxRows; BuyRow ++){
        BuyCard = BuySht.getRange(BuyRow,BuyColCard).getValue();
        // Saves the first Empty Row in the Buy List
        if (BuyCard == '' && BuyEmptyRow == 0) BuyEmptyRow = BuyRow;
        
        // Raise the Flag if Card is found in the Buy List 
        if (BuyCard == WishCardName){
          var BuyCardFndNb = BuySht.getRange(BuyRow,1).getValue();
          BuySht.getRange(BuyRow,1).setValue(BuyCardFndNb + 1);
          BuyCardFound = 1;
        }
      }
      
      // If we reach the Buy List end and no Empty Row is found, insert a new one
      if (BuyCardFound == 0 && BuyEmptyRow == 0){
        BuySht.insertRowAfter(BuyMaxRows - 1);
        // Sets New Row Values 
        BuySht.getRange(BuyMaxRows,1,1,11).setValues(DfltValues);
        BuyEmptyRow = BuyMaxRows;
      }
      
      // Card was not found, add it to the Buy List at the Empty Row Found 
      if (BuyCardFound == 0 && BuyEmptyRow != 0){
        // Copies the Card from the Wish List to the Buy List and clears Data Validation
        WishSht.getRange(WishRow,1,1,NumCol).copyValuesToRange(BuySht, 1, 4, BuyEmptyRow, BuyEmptyRow);
        BuySht.getRange(BuyEmptyRow,1,1,NumCol).clearDataValidations();
        BuySht.getRange(BuyEmptyRow, 11).setValue(new Date()).setNumberFormat('yyyy-MM-dd HH:mm:ss');
      }
    }
  }
}
