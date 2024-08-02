// CONSTANTS
const SHEET_NAME = 'Sheet1'; // Change if your sheet name is different
const PROCESSED_LABEL_NAME = 'Processed';

function processConverge() {
  const threads = GmailApp.search('from:noreply@e-soa.convergeict.com has:attachment');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const processedLabel = getOrCreateLabel(PROCESSED_LABEL_NAME);

  threads.forEach(thread => {
    if (!threadHasLabel(thread, processedLabel)) {
      const messages = thread.getMessages();
      messages.forEach(message => {
        const body = message.getPlainBody();
        
        const accountNumberMatch = body.match(/Account Number:\s*(\d+)/);
        const previousUnpaidBalanceMatch = body.match(/Previous Unpaid Balance:\s*([\d,\.]+)/);
        const currentBillingMatch = body.match(/Current Billing:\s*([\d,\.]+)/);
        const totalAmountDueMatch = body.match(/Total Amount Due:\s*([\d,\.]+)/);
        const dueDateMatch = body.match(/Due Date:\s*(.*)/);

        if (accountNumberMatch && previousUnpaidBalanceMatch && currentBillingMatch && totalAmountDueMatch && dueDateMatch) {
          const accountNumber = "'" + accountNumberMatch[1].replaceAll('*', '');
          const previousUnpaidBalance = previousUnpaidBalanceMatch[1].replaceAll('*', '');
          const currentBilling = currentBillingMatch[1].replaceAll('*', '');
          const totalAmountDue = totalAmountDueMatch[1].replaceAll('*', '');
          const dueDate = formatDate(dueDateMatch[1].replaceAll('*', ''));
       
          // Append the extracted data to the sheet
          sheet.appendRow([accountNumber, previousUnpaidBalance, currentBilling, totalAmountDue, dueDate, new Date()]);

          // Mark the thread as processed by adding the label
          thread.addLabel(processedLabel);
        }
      });
    }
  });
}

function formatDate(dateString) {
  const date = new Date(dateString);
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return formattedDate;
}

function getOrCreateLabel(labelName) {
  const label = GmailApp.getUserLabelByName(labelName);
  return label ? label : GmailApp.createLabel(labelName);
}

function threadHasLabel(thread, label) {
  return thread.getLabels().some(l => l.getName() === label.getName());
}

// To run the script automatically, create a time-driven trigger
function createTrigger() {
  // Delete any existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  ScriptApp.newTrigger('processConverge')
    .timeBased()
    .after(1000) // immediately run after milliseconds. not exact, will vary
    .create();

  ScriptApp.newTrigger('processConverge')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
}
