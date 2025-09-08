function sendScheduledEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const email = data[i][1];         // Column B
    const subject = data[i][2];       // Column C
    const body = data[i][3];          // Column D
    const sendAt = new Date(data[i][4]); // Column E
    const status = data[i][5];        // Column F
    
    if (email && subject && body && sendAt <= now && status !== "SENT") {
      GmailApp.sendEmail(email, subject, body);
      sheet.getRange(i + 1, 6).setValue("SENT");
    }
  }
}
