/*
Author: Anthony Edward Pitts
Data: 2/6/2019
Version: 2.0

Anthony Pitts' GitHub: https://github.com/anthonypitts2022
*/
                                                              //Personalized Email Automator

//This is an application which runs in a google sheet that can send mass emails that are unique to each recipient based on variable values
     //that are defined uniquely per email adress in the currently active google sheet. 
          //A template for the sheet:    https://docs.google.com/spreadsheets/d/1rAxuRpHd5VnrryhLL7ETMg_MtRhEiOYTr8By_NDA8jw/edit#gid=0
          //A template for the doc with email body:    https://docs.google.com/document/d/1wYBpMXac8AuErLD9VYFQ4SxIHG934zoyIeYmNbVjpMY/edit
// every part of email that user wants to be variable should be marked between <> and the word inside <> should be the first row of a column with 
     //user-specific inputs of that variable under it, in that column.
//assumes every recipient is getting the same subject line, but could easily be variable - I thought this would be more helpful

//Current speed: 40 emails: 6 seconds   
            // 1000 emails: 1 minute 36 seconds

function sendEmail(){
  
  //Sends a warning message to confirm that user wants to send the emails
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Please confirm that you want to send these emails.', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {} 
  else if (response == ui.Button.NO) { return;} 
  else { return;}
  
  var spreadsht = SpreadsheetApp.getActive();
  
  var EmailSendSht = spreadsht.getSheetByName("EmailSender");
  
  //gets subject line
  EmailSendSht.activate();
  var EmailSendData = EmailSendSht.getRange(1,1,EmailSendSht.getLastRow(), EmailSendSht.getLastColumn()).getValues(); 
  var subjectColumnIndex = 1;
  var subjectRowIndex = 0;
  
  var docUrl = EmailSendSht.getRange(2,2).getValue().toString();
  var doc = DocumentApp.openByUrl(docUrl); //gets URL of doc with email body
  
  var Sheet = spreadsht.getSheetByName("DataSheet");
  Sheet.activate();

  //gets each paragraph in individual elements of array
  var origParagraphsArray=doc.getParagraphs();
  for(i=0;i<origParagraphsArray.length; i++){
    origParagraphsArray[i] = origParagraphsArray[i].getText();
  }
  
  //gets spreadsheet data into array
  var sheet_data = Sheet.getRange(1,1,Sheet.getLastRow(), Sheet.getLastColumn()).getValues(); 
  // So sheet_data[2][1] would report the cell at row 3 column 2 -- sheet_data[0][x] would report a category of variable
  
  var emailColumnIndex;
  for(emailColumnIndex=0; false == (sheet_data[0][emailColumnIndex].toString().toLowerCase().equals('email')); emailColumnIndex++){
  }                               //finds column index of email column
  
  for(var personCounter=1; personCounter<Sheet.getLastRow(); personCounter++){// personCount is recording each person getting email 
      if(Sheet.getRange(personCounter+1,emailColumnIndex+1).isBlank() ==true){//if there is an empty row (specifically, empty email adress)
          continue;
      }
      var wordArray = [];
      var paragraphsArray = new Array(origParagraphsArray.length);//resets the paragraphsArray back to original version for new person
         for (var q = 0; q < origParagraphsArray.length; q++) {
            paragraphsArray[q] = origParagraphsArray[q];
         }
      for(var paraCounter=0; paraCounter<origParagraphsArray.length; paraCounter++){// paraCounter is recording each paragraph in doc
          wordArray.length = 0;
          var wordArray = paragraphsArray[paraCounter].toString().split(' '); //wordArray holds a word per element of current paragraph
          for(var wordCounter=0; wordCounter<wordArray.length; wordCounter++){ // wordCounter is recording each word in current paragraph
              var currentWord = wordArray[wordCounter].toString();
              if(currentWord.charAt(0).equals('<')){ //a.k.a. "if this is a word to be replaced with unique word to person a"
                  var wordCategory = currentWord.substring(currentWord.lastIndexOf("<") + 1, currentWord.lastIndexOf(">"));//gets keyword in angle brackets
                  for(var t=0; t<sheet_data[0].length; t++){ //t is looping to find categtory of unique word
                      if(wordCategory.equals(sheet_data[0][t].toString())){ //a.k.a "if it found the word category specified in b/w angle brackets"
                          currentWord = sheet_data[personCounter][t].toString() + currentWord.split('>').pop();   //fills in brackets with unique word
                      }
                  }             
              }
          wordArray[wordCounter] = currentWord;
          }
          var updatedParagraph = "";
          for(var h=0;h<wordArray.length;h++){
              updatedParagraph = updatedParagraph + wordArray[h].toString() + " "; // sets updated paragraphs with unique words into current paragraph
          }                                                                         
              paragraphsArray[paraCounter] = updatedParagraph;                             
          }
      var emailBody = ""; //the final message being sent
      for(var paraCounter=0; paraCounter<paragraphsArray.length; paraCounter++){//puts indentation in front of each paragraph
          paragraphsArray[paraCounter] = "\n     " + paragraphsArray[paraCounter].toString(); 
          if(paragraphsArray[paraCounter].toString().substr(7,9).equals('  ')){//if this is an empty paragraph, "delete" it
              paragraphsArray[paraCounter] = "";
          }
           emailBody = emailBody + paragraphsArray[paraCounter].toString();
      }
      
     // Good lines for testing data to be sent 
     /*
      Logger.log('Person Counter: ' + personCounter);
      Logger.log(EmailSendData[subjectRowIndex][subjectColumnIndex]);
      Logger.log('email: ' + sheet_data[personCounter][emailColumnIndex]);
      Logger.log('body: ' + emailBody);
      Logger.log(' END OF THIS PERSON DATA '); */
      
      MailApp.sendEmail(sheet_data[personCounter][emailColumnIndex], EmailSendData[subjectRowIndex][subjectColumnIndex], emailBody); //format= 
                                                                                                                             //(emailAdress,subject,message)
  }
}
