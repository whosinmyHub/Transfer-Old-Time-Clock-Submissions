/** 
 * Purpose: Place last month's time clock submissions into a new spreadsheet 
 * and delete them from the main sheet 
 * */
function archive() {
  try {
    
    //Form Responses 1
    let sheet = SpreadsheetApp.getActiveSheet ();

    var months = ["January", "Feburary", "March", "April", "May", "June", "July",
    "August", "September", "October", "November", "December"];

    //get the int representing last month
    var lastMonth = new Date ().getMonth () - 1;

    //turn that int to a name using the months array
    var lastMonthName = lastMonth == -1 ? months[11] : months[lastMonth];

    //get the year which last month was in, (these ternary operators are used to see if last month was December, because if this month is January the index would be 0 and 0 - 1 is not 11 but rather -1 so we need to manually check that)
    var year = lastMonth == -1 ? new Date ().getFullYear () - 1 : new Date ().getFullYear ();

    //create a new sheet for to put last month's responses in
    let newSheet = SpreadsheetApp.getActiveSpreadsheet ().insertSheet (1);
    var name = "Archive Responses " + lastMonthName + " " + year;
    newSheet.setName (name);

    //copy over the formatting to newSheet
    sheet.getRange ('A1:L1').copyTo (newSheet.getRange ('A1:L1'));
    newSheet.getRange ('A1:L1').setTextStyle (SpreadsheetApp.newTextStyle ().setBold (true).build ());
    
    //copy over last month's submissions
    var thisMonth = new Date ().getMonth ();
    for (i = 2; i < sheet.getMaxRows (); i++) {

      //continue the loop as long as we don't encouter a submission of this month
      if (new Date (sheet.getRange ('A2').getValue ()).getMonth () < thisMonth)
      {
        newSheet.appendRow (sheet.getRange (2 + ':' + 2).getValues ()[0]);
        sheet.deleteRow (2);

        //change the CLOCK IN or CLOCK OUT column to correct background color
        if (newSheet.getRange ('D' + i).getValue () === 'Clock IN')
          newSheet.getRange ('D' + i).setBackground ('#b7e1cd');
        else
          newSheet.getRange ('D' + i).setBackground ('#e06666');

      }
      else 
        break;
    }

    //resize all the columns to fit the text properly
    newSheet.autoResizeColumns (1, newSheet.getMaxColumns ()); 

    
  }
  catch (error) {
    Logger.log (error);
  }
}
