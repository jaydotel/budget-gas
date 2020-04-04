/*
0. Make some test cases e.g. If I feed in searchTerm for payee then what particular do I get <-- Easier to check rather than me manually looking in the spreadsheet
1. Combine statements into one
2. Check Memo only if didn't find in Payee
3. Add conditional formatting rules
4. A function to add category particulars and automatically sort it
*/

// TODO: Haven't tested if this works
function getColByName(sheet, colName, row) {
  var data = sheet.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  if (col != -1) {
    return col;
  }
}

// Gets column of data without extending past the last row
function getColumns(sheet, startRow, startCol, endCol) {
  endCol = endCol || startCol;

  let lastRow = 0;

  // TODO: Remove test value
  if (sheet.getName() === "Bank Statement") {
    lastRow = 26; 
  } else {
    lastRow = sheet.getLastRow();
  }
  
  return sheet.getRange(startCol + startRow + ':' + endCol + lastRow).getValues();
}

/*
 * Extracts a column from a Range i.e. 2d array
 * TODO: Make this more dynamic by checking the column header and then you just need a parameter for column header rather than counting the column number
 *
 * @param {Object} range
 * @param {string} columnIndex
 */
function extractColumnFromArray(range, columnIndex) {
  return range.map(columns => columns[columnIndex]);
}

function launchApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // TODO: Store all spreadsheet sheets/columns into an object so it's not as messy. This way we can just do ss.statement.payees etc... Also, we can just pass in an object instead of all variables. And can use this for future Google sheet projects. All functions like getCol, getRow... contained within this function. Need to detect for each row/col where the end of the data is.
  const statementSheet = ss.getSheetByName('Bank Statement');
  const statement = getColumns(statementSheet, 9, 'A', 'H');
  const dates = extractColumnFromArray(statement, 0);
  const payees = extractColumnFromArray(statement, 4);
  const memos = extractColumnFromArray(statement, 5);
  const statementStartRow = 9;
  
  const expensesMappingTableSheet = ss.getSheetByName('ExpensesMappingTable');
  const expensesMappingTable = getColumns(expensesMappingTableSheet, 2, 'A', 'B');
  const searchTerms = extractColumnFromArray(expensesMappingTable, 0);
  const categories = extractColumnFromArray(expensesMappingTable, 1);

  /* 
    1. For each transaction/payee, check if it can be mapped back to a category
    2. 
  */

  let transactionsMappedToCategories = payees.map((payee, index) => mapTransactionToCategories(payee, memos, index, searchTerms, categories));

  const particularStartCell = 'J' + (statementStartRow); // TODO: Change to actual column
  const particularEndCell = 'J' + (statementStartRow + transactionsMappedToCategories.length - 1);

  // Logger.log(transactionsMappedToCategories);
  Logger.log(statementStartRow);
  Logger.log(statementStartRow + transactionsMappedToCategories.length);
  Logger.log(transactionsMappedToCategories.length);
  Logger.log(transactionsMappedToCategories);

  statementSheet.getRange(particularStartCell + ':' + particularEndCell).setValues(transactionsMappedToCategories); // TODO: Do this all at once...

  // Loop through payees with searchTerm. If found, return particular for that row where it was found
  // while()
}

function mapTransactionToCategories(payee, memos, payeesIndex, searchTerms, categories) {
  // Ignore string case
  payee = payee.toUpperCase();

  let matches = [];

  // Find the categories for this transaction
  searchTerms.forEach((searchTerm, searchTermsIndex) => {
    searchTerm = searchTerm.toUpperCase();


    let memo = memos[payeesIndex].toUpperCase();
    Logger.log('payee is ' + payee + ', memo is ' + memo + ', searchTerm is ' + searchTerm + ', result is ' + payee.indexOf(searchTerm));
    if (payee.indexOf(searchTerm) > -1 || memo.indexOf(searchTerm) > -1) {
      matches.push(categories[searchTermsIndex]);
    }
  });

  Logger.log('matches: ', matches);

  if (matches.length === 0) {
    const notFoundMessage = 'No matches';
    matches.push(notFoundMessage);
  }

  return mergeArrayElements(matches); // ['Personal Care, HelloWorld']
}

// Merges all array elements into one element then returns a new array
function mergeArrayElements(array) {
  return [array.join()];
}