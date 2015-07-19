 /* Script enabling to transfer content of emails from the label to excel spreadsheet */

function getemails() 
{
   /* This enables the script to access the specified label, in this case “Your Label */
   
   var label = GmailApp.getUserLabelByName(“Your Label“);
   var threads = label.getThreads();

       for (var i = 0; i < threads.length; i++) 
       { 
 
   var messages=threads[i].getMessages();  
 
       for (var j = 0; j < messages.length; j++) 
       {
    
    /*this, tells the script what subject lines I am after: */
       
   var message=messages[j];
       
       process(message);     

       }     
       }
   

function process(message) 
   {
     
     /* access plain body of an e-mail and open a specified spreadsheet */
     
     var body= message.getPlainBody();
     var id= "GoogleSheetId";
     var ss = SpreadsheetApp.openById(id);
     var sheet = ss.getActiveSheet();
     
     /* create a new row in the spreadsheet and paste the body of an e-mail */
     
     sheet.appendRow([body]);
     
     /* once completed mark the message as read and then delete it */
 
         markMessagesRead(message);
             
             deleteMessage(message);           
    }  
   
         function markMessagesRead(message) 
            {
             message.markRead();
            }
             function deleteMessage(message)
                {
                 message.moveToTrash();

                }
}   
