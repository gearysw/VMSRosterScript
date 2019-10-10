function updateInquirer() {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let newSignUpSheet = spreadsheet.getSheetByName('Sign up form');
    let lastRow = newSignUp.getLastRow();

    if (lastRow > 1) { // execute script 
        let inquirerEmails = [];
        for (let i = 2; i <= lastRow; i++) {
            let email = newSignUpSheet.getRange(lastRow, 2);
            inquirerEmails.push(email);
            Logger.log(inquirerEmails);
        }
    }
}

// function sendEmail(recipients, subject, message, replyto) {
//     MailApp.sendEmail({
        
//     });
// }

function moveEntry() {
    let ss = SpreadsheetApp.getActive();
    let sh0 = ss.getSheetByName('Sign up form');
    let sh1 = ss.getSheetByName('Inquirer');
    
}