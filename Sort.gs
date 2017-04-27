// **********************************************
// function fcnSortByStatus()
//
// Sorts all cards in the deck selected
// according to Status
// 
// **********************************************

function fcnSortByStatus() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var MaxRows = actSht.getMaxRows();
  var MaxCols = actSht.getMaxColumns();
  var NumCols = MaxCols - 1;
  var SectRange;
  var SectFirstRow = actSht.getActiveCell().getRow();
  var SectNumRows;
  var EndValue;

  if (actShtName == 'Wish'){  
    // Finds the End of the Section to determine the number or rows in the section
    for (var Row = SectFirstRow; Row <= MaxRows; Row++){
      EndValue = actSht.getRange(Row, 2).getValue();
      if (EndValue == 'End'){
        SectNumRows = Row - SectFirstRow - 1;
        Row = MaxRows + 1;
      }
    }
    
    SectRange = actSht.getRange(SectFirstRow+1, 2, SectNumRows, NumCols);
    SectRange.sort(6);
  } 
}


// **********************************************
// function fcnSortByName()
//
// Sorts all cards in the deck selected
// according to Card Name
// 
// **********************************************

function fcnSortByName() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var MaxRows = actSht.getMaxRows();
  var MaxCols = actSht.getMaxColumns();
  var NumCols = MaxCols - 1;
  var SectRange;
  var SectFirstRow = actSht.getActiveCell().getRow();
  var SectNumRows;
  var EndValue;

  if (actShtName == 'Wish'){  
    // Finds the End of the Section to determine the number or rows in the section
    for (var Row = SectFirstRow; Row <= MaxRows; Row++){
      EndValue = actSht.getRange(Row, 2).getValue();
      if (EndValue == 'End'){
        SectNumRows = Row - SectFirstRow - 1;
        Row = MaxRows + 1;
      }
    }
    
    SectRange = actSht.getRange(SectFirstRow+1, 2, SectNumRows, NumCols);
    SectRange.sort(2);
  }
}

// **********************************************
// function fcnSortBySet()
//
// Sorts all cards in the deck selected
// according to Set
// 
// **********************************************

function fcnSortBySet() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var MaxRows = actSht.getMaxRows();
  var MaxCols = actSht.getMaxColumns();
  var NumCols = MaxCols - 1;
  var SectRange;
  var SectFirstRow = actSht.getActiveCell().getRow();
  var SectNumRows;
  var EndValue;

  if (actShtName == 'Wish'){  
    // Finds the End of the Section to determine the number or rows in the section
    for (var Row = SectFirstRow; Row <= MaxRows; Row++){
      EndValue = actSht.getRange(Row, 2).getValue();
      if (EndValue == 'End'){
        SectNumRows = Row - SectFirstRow - 1;
        Row = MaxRows + 1;
      }
    }
    
    SectRange = actSht.getRange(SectFirstRow+1, 2, SectNumRows, NumCols);
    SectRange.sort([3,2]);
  }
}

// **********************************************
// function fcnSortBySetBuy()
//
// Sorts all cards in the Buy List
// according to the Card Set
// 
// **********************************************

function fcnSortBySetBuy() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var MaxRows = actSht.getMaxRows();
  var MaxCols = actSht.getMaxColumns();
  var SectRange;

  var BuyColSet = 3;
  var BuyRowStart = 7;

  if (actShtName == 'Buy'){  
    SectRange = actSht.getRange(BuyRowStart, 1, MaxRows - BuyRowStart, MaxCols);
    SectRange.sort([3,2]);
  }
}