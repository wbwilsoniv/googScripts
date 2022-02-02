// New Version changes: 
// 1) getFrom > getTo method
// 2) mailContent added for order number
// 3) created ss includes order number

function GetAddresses() {
    // Get the active spreadsheet
   var ss = SpreadsheetApp.getActiveSpreadsheet();  
 
   // Label to search
   var userInputSheet = ss.getSheets()[0];
   
   var labelName = userInputSheet.getRange("B2").getValue();
   
   // Create / empty the target sheet
   var sheetName = "Label: " + labelName;
   var sheet = ss.getSheetByName (sheetName) || ss.insertSheet (sheetName, ss.getSheets().length);
   sheet.clear();
   
   // Get all messages in a nested array (threads -> messages)
   var addressesOnly = [];
   var messageData = [];
 
   var startIndex = 0;
   var pageSize = 100;
   while (1)
   {
     // Search in pages of 100
     var threads = GmailApp.search ("label:" + labelName, startIndex, pageSize);
     if (threads.length == 0)
       break;
     else
       startIndex += pageSize;
        
     // Get all messages for the current batch of threads
     var messages = GmailApp.getMessagesForThreads (threads);
     
     // Loop over all messages
     for (var i = 0; i < messages.length ; i++)
     {
       // Loop over all messages in this thread
       for (var j = 0; j < messages[i].length; j++)
       {
         var mailFrom = messages[i][j].getTo ();
         var mailDate = messages[i][j].getDate ();
         var mailContent = messages[i][j].getBody();
         // mailFrom format may be either one of these:
         // name@domain.com
         // any text <name@domain.com>
         // "any text" <name@domain.com>
         
         var name = "";
         var email = "";
         var order = "";
         var orderNums = mailContent.match (/#\d{4}/);
         if (orderNums){
           order = orderNums;
         }
         else 
         {
           order = 0000;
         }

         var matches = mailFrom.match (/\s*"?([^"]*)"?\s+<(.+)>/);
         if (matches)
         {
           name = matches[1];
           email = matches[2];
         }
         else
         {
           email = mailFrom;
         }
         // Check if (and where) we have this already
         var index = addressesOnly.indexOf (mailFrom);
         if (index > -1)
         {
           // We already have this address -> remove it (so that the result is ordered by data from new to old)
           addressesOnly.splice(index, 1);
           messageData.splice(index, 1);
         }
         
         // Add the data
         addressesOnly.push (mailFrom);
         messageData.push ([name, email, order, mailDate]);
       }
     }
   }
   
   // Add data to corresponding sheet
   sheet.getRange (1, 1, messageData.length, 4).setValues (messageData);
 }
 
 
 //
 // Adds a menu to easily call the script
 //
 function onOpen ()
 {
   var sheet = SpreadsheetApp.getActiveSpreadsheet ();
   
   var menu = [ 
     {name: "Extract email addresses",functionName: "GetAddresses"}
   ];  
   
   sheet.addMenu ("WW Scripts", menu);    
 }
 