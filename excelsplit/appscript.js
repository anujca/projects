function sendDailyAlerts() {
  // === Configuration ===
  const MAIN_SHEET_NAME = '2022-23'; // Name of your main data sheet
  const CONTACT_SHEET_NAME = 'contact'; // Name of your contact sheet
  const DATE_FORMAT = 'dd/MM/yyyy'; // Default date format in "Next F/Up date" column
  const SUBJECT_TEMPLATE = 'Daily Alert: Tasks Due Today for'; // Email subject template
  const REPORT_EMAIL = 'ak0126002@gmail.com'; // Email for error and completion reports
  const REPORT_SUBJECT = 'Script Execution Report'; // Subject for the report email
  const CC_EMAIL = 'alokwaiting@gmail.com'; // Email for CC in all outgoing emails

  // === Initialize Report Variables ===
  const reportErrors = [];
  let emailsSent = 0;
  let emailsFailed = 0;
  let rowsSelected = 0;

  try {
    Logger.log("Starting script execution...");

    // === Access the Spreadsheet and Sheets ===
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
    const contactSheet = ss.getSheetByName(CONTACT_SHEET_NAME);

    if (!mainSheet || !contactSheet) {
      throw new Error(`One or both sheets ("${MAIN_SHEET_NAME}", "${CONTACT_SHEET_NAME}") not found.`);
    }

    Logger.log("Sheets loaded successfully.");

    // === Retrieve Data from Sheets ===
    const mainData = mainSheet.getDataRange().getValues();
    const contactData = contactSheet.getDataRange().getValues();

    if (mainData.length < 2) {
      throw new Error('Main sheet does not contain data.');
    }

    if (contactData.length < 2) {
      throw new Error('Contact sheet does not contain data.');
    }

    Logger.log(`Main sheet has ${mainData.length - 1} rows.`);
    Logger.log(`Contact sheet has ${contactData.length - 1} rows.`);

    // === Identify Column Indices in Main Sheet ===
    const mainHeaders = mainData[0];
    const dateColIndex = mainHeaders.indexOf('Next F/Up date');
    const handledByColIndex = mainHeaders.indexOf('Handled by');

    if (dateColIndex === -1 || handledByColIndex === -1) {
      throw new Error('Required columns ("Next F/Up date", "Handled by") not found in Main sheet.');
    }

    // === Identify Column Indices in Contact Sheet ===
    const contactHeaders = contactData[0];
    const acronymColIndex = contactHeaders.indexOf('Acronym');
    const fullNameColIndex = contactHeaders.indexOf('Full Name');
    const emailColIndex = contactHeaders.indexOf('Email');

    if (acronymColIndex === -1 || fullNameColIndex === -1 || emailColIndex === -1) {
      throw new Error('Required columns ("Acronym", "Full Name", "Email") not found in Contact sheet.');
    }

    Logger.log("Successfully identified required columns.");

    // === Create Mapping from Acronym to Contact Details ===
    const contactMap = {};
    for (let i = 1; i < contactData.length; i++) {
      const acronym = contactData[i][acronymColIndex].toString().trim();
      const fullName = contactData[i][fullNameColIndex].toString().trim();
      const email = contactData[i][emailColIndex].toString().trim();
      if (acronym) {
        contactMap[acronym] = { fullName, email };
      }
    }

    Logger.log("Created mapping from acronym to contact details.");

    // === Get Today's Date in Specified Format ===
    const today = new Date();
    const todayStr = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), DATE_FORMAT);
    Logger.log(`Today's date: ${todayStr}`);

    // === Function to Handle Multiple Date Formats and Add Leading Zeros ===
    function parseAndFormatDate(dateStr) {
      if (dateStr instanceof Date) return dateStr;

      dateStr = dateStr.replace(/\s+/g, '').trim(); // Remove whitespace

      // Handle the date string format
      const parts = dateStr.split(/[-\/]/);
      if (parts.length === 3) {
        let day = parseInt(parts[0], 10); // Remove leading zeros
        let month = parseInt(parts[1], 10); // Remove leading zeros
        const year = parseInt(parts[2], 10);

        // Add leading zeros where needed
        day = day < 10 ? `0${day}` : day;
        month = month < 10 ? `0${month}` : month;

        // Return formatted date
        return `${day}/${month}/${year}`;
      }
      return null; // Invalid date format
    }

    // === Collect Rows Per Email ===
    const emailRowsMap = {};

    for (let i = 1; i < mainData.length; i++) {
      const rowDate = mainData[i][dateColIndex];
      const rowDateFormatted = parseAndFormatDate(rowDate);

      if (rowDateFormatted === todayStr) {
        rowsSelected++;
        const handledBy = mainData[i][handledByColIndex].toString().trim();
        if (handledBy) {
          const acronyms = handledBy.split('/').map(acronym => acronym.trim());
          acronyms.forEach(acronym => {
            if (contactMap[acronym]) {
              const email = contactMap[acronym].email;
              const fullName = contactMap[acronym].fullName;
              if (email) {
                if (!emailRowsMap[email]) {
                  emailRowsMap[email] = {
                    fullName: fullName,
                    rows: []
                  };
                }
                emailRowsMap[email].rows.push(mainData[i]);
              }
            } else {
              reportErrors.push(`Acronym "${acronym}" not found in contact sheet.`);
              Logger.log(`Acronym "${acronym}" not found in contact sheet.`);
            }
          });
        }
      }
    }

    Logger.log(`Total rows selected for today's date: ${rowsSelected}`);

    if (Object.keys(emailRowsMap).length === 0) {
      Logger.log('No alerts to send today.');
    } else {
      Logger.log('Sending emails to recipients...');
    }

    // === Prepare and Send Emails ===
    const headers = mainHeaders;
    for (const email in emailRowsMap) {
      const recipient = email;
      const name = emailRowsMap[email].fullName;
      const rows = emailRowsMap[email].rows;

      // Add current date in subject
      const subject = `${SUBJECT_TEMPLATE} ${name} - ${todayStr}`;
      const executionDateStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

      let htmlBody = `
        <div style="font-family: Arial, sans-serif; color: #333;">
          <h2 style="color: #4CAF50;">Daily Task Alert</h2>
          <p><strong>Date:</strong> ${executionDateStr}</p>
          <p>Hello ${name},</p>
          <p>The following tasks are due today:</p>
          <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">
            <thead>
              <tr style="background-color:#f2f2f2;">
                ${headers.map(header => `<th style="padding: 8px; text-align: left;">${header}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${rows.map(row => `
                <tr>
                  ${row.map(cell => `<td style="padding: 8px;">${cell}</td>`).join('')}
                </tr>
              `).join('')}
            </tbody>
          </table>
          <p>Please take the necessary actions.</p>
          <p>Best regards,<br>Alok Goel</p>
        </div>
      `;

      try {
        MailApp.sendEmail({
          to: recipient,
          subject: subject,
          htmlBody: htmlBody,
          cc: CC_EMAIL // Add CC here
        });
        emailsSent++;
        Logger.log(`Email sent to ${recipient}`);
      } catch (e) {
        emailsFailed++;
        reportErrors.push(`Failed to send email to ${recipient}: ${e.toString()}`);
        Logger.log(`Failed to send email to ${recipient}: ${e.toString()}`);
      }
    }

    Logger.log(`Emails sent successfully: ${emailsSent}`);
    Logger.log(`Emails failed to send: ${emailsFailed}`);
  } catch (error) {
    reportErrors.push(`General Error: ${error.toString()}`);
    Logger.log(`General Error: ${error.toString()}`);
  } finally {
    // === Prepare and Send Execution Report ===
    let reportBody = `
      <div style="font-family: Arial, sans-serif; color: #333;">
        <h2>Script Execution Report</h2>
        <p><strong>Execution Date:</strong> ${Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss")}</p>
        <p><strong>Total Rows Selected for Processing:</strong> ${rowsSelected}</p>
        <p><strong>Emails Sent Successfully:</strong> ${emailsSent}</p>
        <p><strong>Emails Failed:</strong> ${emailsFailed}</p>
        <p><strong>Errors Encountered:</strong></p>
        ${reportErrors.length > 0 ? `
          <ul>
            ${reportErrors.map(err => `<li>${err}</li>`).join('')}
          </ul>
        ` : `<p>No errors encountered.</p>`}
      </div>
    `;

    MailApp.sendEmail({
      to: REPORT_EMAIL,
      subject: REPORT_SUBJECT,
      htmlBody: reportBody
    });

    Logger.log('Execution report sent.');
  }
}
