/*
  MELTBot Version 1.0.1
  Revision tracker: https://tinyurl.com/MeltBotRevTracker
  
  This form gathers help requests from students, checks for assorted requirements, and contacts the appropriate techs with relevant job information.
  
  Author:
  Victor Kojenov
  vkojenov@pdx.edu
*/

//Function to send an email
function SendEmail(recipients, subject_line, message, reply_to) {
  MailApp.sendEmail
  ({
    to:""+recipients+"",
    subject:subject_line,
    htmlBody:message+"<br/>",
    replyTo: ""+reply_to
  });
  
}

//Function to get contact info for melt techs for a particular machine:
function GetMeltTechs(machine) {
  var emails = "psumelt@gmail.com, chuning@pdx.edu";
  var techInfo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MELT Tech Info");
  var tech_lastRow = techInfo.getLastRow();
  
  //Determine column based on machine type
  var machCol = 0;
  for (var i = 3; i<9; i++) {
    if (techInfo.getRange(2, i).getValue() == machine) {
      machCol = i;
    }
  }
  
  //Add techs who can operate the machinery to the email list
  for (var i=3; i<(tech_lastRow+1); i++) {
    if (techInfo.getRange(i, machCol).getValue() == "Y") {
      emails = emails + ", " + techInfo.getRange(i, 2).getValue();
    }
  }
  
  //Return the email list
  return emails;
}

//Function to gather all tech e-mails
function GetAllTechs() {
  var emails = "psumelt@gmail.com, chuning@pdx.edu";
  var techInfo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MELT Tech Info");
  var tech_lastRow = techInfo.getLastRow();
  
  //Goes through each row and adds tech email to variable
  for (var i = 3; i<(tech_lastRow + 1); i++) {
    emails = emails + ", " + techInfo.getRange(i, 2).getValue();
  }
  
  //Returns emails
  return emails;
}


//MAIN FUNCTION

function Request() { 
  //Setting up spreadsheet and checking for a request:
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var requests = spreadsheet.getSheetByName("Project Requests");
  var req_lastRow = requests.getLastRow();
  
  if (req_lastRow != 1) {      //Executes if there is a new request
    //Setting up some arrays and variables to categorize useful data
    
    //Student's contact info:
    var contactInfo = [
      requests.getRange(req_lastRow, 2).getValue(),    //E-mail address
      requests.getRange(req_lastRow, 3).getValue(),    //Full name
      requests.getRange(req_lastRow, 4).getValue()];   //Student type
    
    //Request type:
    var reqType = requests.getRange(req_lastRow, 5).getValue();
    
    if (reqType == "Project Request") {
      
      //Shop safety check
      var shopSafetyProj = requests.getRange(req_lastRow, 6).getValue();  //Shop safety check
      if (shopSafetyProj == "No") {
        //Shop safety e-mail - Parameters
        var meltTechs = GetAllTechs();
        var studentSubject = "Shop Safety Class Information"
        var studentMsg = "Hi " + contactInfo[1] + ", <br/> <br/> Unfortunately, the MELT Team cannot help you with your project as you haven't taken the Machine Shop Safety course. Don't be discouraged, though, the process is very simple and takes two sessions. You can get find out more at: http://www.pdx.edu/mme/safety-class-registration. <br/> <br/>Once you have completed both parts of the Safety class, we'd be more than happy to help you with your project! <br/> <br/> Best regards, <br/> MELT Bot";
        
        SendEmail(contactInfo[0], studentSubject, studentMsg, meltTechs);
      
      } else if (shopSafetyProj == "Yes") {
        //Variable setup for Project Request
        var projType = requests.getRange(req_lastRow, 8).getValue();        //Project type
        var projMach = requests.getRange(req_lastRow, 9).getValue();        //Which machine to use
        var projDesc = requests.getRange(req_lastRow, 10).getValue();       //Short project description

        var meltTechs = GetMeltTechs(projMach);        
        //Email to techs - Parameters
        var techSubject = "New " + projMach + " " + projType;
        var techMsg = "Hey " + projMach + " techs! <br/> <br/> " + contactInfo[1] + ", a " + contactInfo[2] + ", has a " + projMach + " " + projType + " for you. They've provided the following info for their project: <br/> <br/>" + projDesc.italics() + "<br/> <br/> Please respond to their request within 48 hours to schedule a time to help them. Responding to this e-mail will contact the correct people. Also, don't forget to reply all, and generate an invoice for any jobs outside of the ME department! <br/> <br/> Best regards, <br/> MELT Bot";
        
        //Sending Email to techs
        SendEmail(meltTechs, techSubject, techMsg, (contactInfo[0] + ", " + meltTechs));
        
        //Email to Student - Parameters
        var studentSubject = "Your " + projMach + " Project Request";
        var studentMsg = "Hey " + contactInfo[1] + ", <br/> <br/> Thanks for your project submission! I've sent an email to the " + projMach + " techs with your project information, and a tech will contact you soon to schedule a time to work on your project. Please allow up to 48 hours for a tech to contact you, and please keep in mind that they may not be able to get to your project immediately. <br/> In the meantime, if you have any relevant project files (Solidworks parts/drawings, hand sketches, etc.), please reply and attach them. It will help the techs get a better idea of what needs to be done! <br/> <br/> Best regards, <br/> MELT Bot";
        
        //Sending Email to student
        SendEmail(contactInfo[0], studentSubject, studentMsg, meltTechs);
      }
    } else if (reqType == "Machine Training") {
      //Shop safety check
      var shopSafetyMT = requests.getRange(req_lastRow, 7).getValue();
      if (shopSafetyMT == "No") {
        //Shop safety e-mail - Parameters
        var studentSubject = "Shop Safety Class Information"
        var studentMsg = "Hi " + contactInfo[1] + ", <br/> <br/> Unfortunately, the MELT Team cannot train you as you haven't taken the Machine Shop Safety course. Don't be discouraged, though, the process is very simple and takes two sessions. You can get find out more at: http://www.pdx.edu/mme/safety-class-registration. <br/> <br/>Once you have completed both parts of the Safety class, we'd be more than happy to get you trained! <br/> <br/> Best regards, <br/> MELT Bot";
        
        SendEmail(contactInfo[0], studentSubject, studentMsg, meltTechs);
        
      } else if (shopSafetyMT == "Yes") {
        var machType = requests.getRange(req_lastRow, 12).getValue();   //Machine type
        
        //E-mail to techs - Parameters
        var meltTechs = GetMeltTechs(machType);
        var techSubject = "New " + machType + " Training Request"
        var techMsg = "Hey " + machType + " techs! <br/> <br/> " + contactInfo[1] + ", a " + contactInfo[2] + ", would like to get trained for " + machType + ". Please respond to this e-mail to arrange a time to get them trained! <br/> <br/> Best regards, <br/> MELT Bot";
        
        //Sending Email to techs
        SendEmail(meltTechs, techSubject, techMsg, (contactInfo[0] + ", " + meltTechs));
        
        //E-mail to student - Parameters
        var studentSubject = "Your " + machType + " Training Request";
        var studentMsg = "Hey " + contactInfo[1] + ", <br/> <br/> We're glad to see that you're interested in getting trained on the " + machType + "! I've notified the " + machType + " techs that you're interested, and one of them will contact you within 48 hours to schedule your training. <br/> <br/> Best regards, <br/> MELT Bot";
        
        //Sending Email to student
        SendEmail(contactInfo[0], studentSubject, studentMsg, meltTechs);
      }
      
    } else if (reqType == "General Inquiry") {
      //E-mail to techs - Parameters
      var meltTechs = GetAllTechs();
      var question = requests.getRange(req_lastRow, 11).getValue();
      var techSubject = "New General Inquiry"
      var techMsg = "Hey techs! <br/> <br/> " + contactInfo[1] + ", a " + contactInfo[2] + ", has a question for you: <br/> <br/>" + question.italics() + "<br/> <br/> Please respond to this question within 48 hours! <br/> <br/> Best regards, <br/> MELT Bot";
      
      //Sending Email to techs
      SendEmail(meltTechs, techSubject, techMsg, (contactInfo[0] + ", " + meltTechs));
      
      //Email to student - Parameters
      var studentSubject = "Your General Inquiry"
      var studentMsg = "Hey " + contactInfo[1] + ", <br/> <br/> Thanks for getting in touch. We've received your general inquiry, and a MELT Tech will respond within 48 hours to answer your questions. <br/> <br/> Best regards, <br/> MELT Bot";
      
      //Sending Email to student
      SendEmail(contactInfo[0], studentSubject, studentMsg, meltTechs);
    }
    
    //Move row to backup
    var backup = spreadsheet.getSheetByName("Project Requests Backup");
    var backup_lastRow = backup.getLastRow();
    
    var response = requests.getRange(req_lastRow, 1, 1, 12).getValues(); //Get entire response
    backup.getRange(backup_lastRow + 1, 1, 1, 12).setValues(response); //Write response to the next empty row in backup
    requests.deleteRow(req_lastRow);
    requests.insertRowAfter(req_lastRow);
  }
}