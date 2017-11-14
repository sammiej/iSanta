function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var numRows = sheet.getLastRow()-1;
  var data = sheet.getRange(startRow, 1, numRows, 2).getValues();        
  var names = sheet.getRange(startRow, 2, numRows, 1).getValues();         
  
  var randNames = randomizeNames(names, numRows);
  
  for (i in data) {
    var row = data[i];
    var emailAddress = row[0];
    var name = row[1];
    var subject = name + ", you are secret santa for...";
    var message = randNames[i];
    var gif = UrlFetchApp.fetch("https://i.pinimg.com/originals/c5/23/47/c523475df32a2010a20f74cb22dffd4a.gif")  // Change link to your gif of choice
                         .getBlob()
                         .setName("gif");
    MailApp.sendEmail({
      to: emailAddress, 
      subject: subject, 
      htmlBody: message + "<br><br><img src='cid:gif'>",
      inlineImages: {
        gif: gif
      }
    });
  }
  
  /** 
    * Uncomment this to to see the shuffler in action (the 3rd column will be filled with each person's matched santa. but comment out lines 19-26 to avoid email spamming)
    *
  for (var i=2; i<=numRows+1; i++) {
    var cell = sheet.getRange(i, 3);
    cell.setValue(randNames[i-2]);
  }*/
}


function randomizeNames(names, n) {
  var copyNames = names.slice(0);            
  
  for (var i=n-1; i>=0; i--) {
    do {
      var rand = Math.floor(Math.random() * n);    
    } while (names[rand] == copyNames[i]);         
    
    var temp = names[i];
    names[i] = names[rand];
    names[rand] = temp;
  }
  
  return names;
}
