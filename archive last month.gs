/** 
 * Purpose: Place last month's time clock submissions into a new spreadsheet 
 * and delete them from the main sheet 
 * */
function archive() {
  try {

    var months = ['January', 'Febuary', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

    //a variable to hold the current month
    let month = new Date ().getMonth ();

    //a variable to hold the year of last month's submissions, if we are in January, the year should be last year
    let year;
    if (month === 0)
      year = 11;
    
    else 
      year = new Date ().getFullYear ();

    const sheet = SpreadsheetApp.getActiveSpreadsheet ().getSheetByName ('Form Responses 1');
    const newSheet = SpreadsheetApp.getActiveSpreadsheet ().insertSheet ('Archive Responses ' + months [month - 1] + ' ' + year);

    //copy the formatting of the first row to the new sheet
    sheet.getRange ('A1:K1').copyTo (newSheet.getRange ('A1:K1'));
    //make the columns fit the words
    newSheet.autoResizeColumns (1, 11);
    //make words bold as they are in the original sheet
    newSheet.getRange (1, 1, 1, 11).setTextStyle (SpreadsheetApp.newTextStyle ().setBold (true).build ()); 

    var i = 2;
    for (i; i < sheet.getMaxRows (); i++) {

      const date = new Date (sheet.getRange ('A'+ i).getValue ());

      //if the row being looked at was a submission from the last month
      if (date.getMonth () < month || new Date ().getFullYear () > date.getFullYear ()) {

       newSheet.appendRow (sheet.getRange (i + ':' + i).getValues ()[0]);

       //add back the formatting, which is just the background color of the Clock In /Clock Out column
       if (newSheet.getRange (newSheet.getLastRow (), 4).getValue () === 'Clock OUT')
        newSheet.getRange (newSheet.getLastRow (), 4).setBackground ('#e06666');

        else
          newSheet.getRange (newSheet.getLastRow (), 4).setBackground ('#b7e1cd');
        
      }

      //if we reached a row whose date is not in the previous' month, stop looping as the rest of the rows/submissions will be for the current month
      else
        break;
    }

    //resize the columns to fit their text
    newSheet.autoResizeColumns (1, newSheet.getLastColumn ());

    //delete last month's submissions from the main sheet
    sheet.deleteRows (2, i);
  }
  catch (error) {
    console.log (error);
  }
}
