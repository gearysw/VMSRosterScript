function updateInquirer() {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let newSignUpSheet = spreadsheet.getSheetByName('Sign up form');
    let lastRow = newSignUp.getLastRow();

    if (lastRow > 1) { // execute script 
        let inquirerEmails = [];
        for (let i = 2; i <= lastRow; i++) {
            let email = newSignUpSheet.getRange(lastRow, 1);
            inquirerEmails.push(email);
            Logger.log(inquirerEmails);
        }
    }
}

// function sendEmail(recipients, subject, message, replyto) {
//     MailApp.sendEmail({
        
//     });
// }