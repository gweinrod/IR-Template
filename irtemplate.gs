//Purpose: To standardize IRS, saving time for consultants and clinicians and allowing more for instruction

var document = DocumentApp.getActiveDocument()
var userName = Session.getEffectiveUser().getUsername();

//init menu
function onInstall(){
  onOpen();
}

//menu
function onOpen() {
  var menu = DocumentApp.getUi().createMenu('IR Template')
      .addItem('Insert Instructor Information', 'insertInstructor')
      .addItem('VV:P2P   - Picture to Picture', 'insertP2P')
  menu.addToUi();
}

//writes the user's name, date
function insertInstructor() {
  var cursor = DocumentApp.getActiveDocument().getCursor();
  var date = Utilities.formatDate(new Date(), "PST", "MMM dd yyyy @h:mm a ")
  if (cursor) {
    var element = cursor.insertText(userName + " " + date + "\n");
    if (element) {
      element.setBold(true);
      length = element.getText().length
      var position;
      try {
        position = document.newPosition(element, cursor.getOffset() + length);
      } catch (error) {
        position = document.newPosition(element, cursor.getOffset() + length - 1)
      }
      document.setCursor(position);
    } else {
      locationError();
    }
  } else {
    cursorError();
  }
}

//writes template for Picture to Picture
function insertP2P() 
{
    var cursor = DocumentApp.getActiveDocument().getCursor();
    if (cursor) {
    var element = cursor.insertText("Picture to Picture\nBook:\t\tLevel:\t\tNumber:\t\t\nPICS: %\tind:\t\tpr:");
    if (element) {
      element.setBold(true);
      length = element.getText().length
      var position;
      try {
        position = document.newPosition(element, cursor.getOffset() + length);
      } catch (error) {
        position = document.newPosition(element, cursor.getOffset() + length - 1)
      }
      document.setCursor(position);
    } else locationError();
    } else cursorError();
}

function locationError() {
  DocumentApp.getUi().alert('Cannot insert text at this cursor location.');
}

function cursorError() {
  DocumentApp.getUi().alert('Cannot find a cursor in the document.');
}
