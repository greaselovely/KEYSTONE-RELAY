function onOpen() {
  // Initialize sheets if needed
  initializeSheets();

  // Create the menu
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Email System")
    .addItem("Compose New Email", "showEmailComposer")
    .addItem("Send Test Email", "showSendTestDialog")
    .addItem("Send to All Subscribers", "confirmBulkSend")
    .addItem("View Analytics Dashboard", "showAnalytics")
    .addSeparator()
    .addItem("System Settings", "showSystemSettings")
    .addToUi();
}

function doGet(e) {
  const action = e.parameter.action;

  if (!action) {
    // If no action specified, return the subscription management page
    // Use template to inject constants
    const template = HtmlService.createTemplateFromFile("Subscription");

    // Pass constants to the template
    template.WEB_APP_URL = WEB_APP_URL;
    template.NEWSLETTER_NAME = NEWSLETTER_NAME;

    return template
      .evaluate()
      .setTitle("Manage Your Subscription")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  switch (action) {
    case "open":
      const emailOpen = e.parameter.email;
      if (!emailOpen) return ContentService.createTextOutput("Invalid request");

      trackOpen(emailOpen, e.parameter.template, e.parameter.hash);
      // Return a 1x1 transparent pixel
      return ContentService.createTextOutput(
        "GIF89a\x01\x00\x01\x00\x80\x00\x00\xff\xff\xff\x00\x00\x00!\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;"
      ).setMimeType(ContentService.MimeType.GIF);

    case "click":
      const emailClick = e.parameter.email;
      if (!emailClick)
        return ContentService.createTextOutput("Invalid request");

      trackClick(emailClick, e.parameter.template, e.parameter.hash);
      // Redirect to the target URL
      return HtmlService.createHtmlOutput(`
          <html>
            <head>
              <meta http-equiv="refresh" content="0;URL='${e.parameter.url}'" />
            </head>
            <body>
              Redirecting...
            </body>
          </html>
        `);

    case "unsubscribe":
      const emailUnsub = e.parameter.email;
      if (!emailUnsub)
        return ContentService.createTextOutput("Invalid request");

      unsubscribeUser(emailUnsub);

      // Return success JSON response for the form
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, message: "Unsubscribed successfully" })
      ).setMimeType(ContentService.MimeType.JSON);

    case "subscribe":
      const emailSub = e.parameter.email;
      const name = e.parameter.name || "";

      if (!emailSub)
        return ContentService.createTextOutput(
          JSON.stringify({
            success: false,
            message: "Email parameter is missing",
          })
        ).setMimeType(ContentService.MimeType.JSON);

      try {
        // Simple direct update approach
        const normalizedEmail = emailSub.toLowerCase().trim();
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SUBSCRIBERS_SHEET);

        Logger.log(`Subscribe request for ${normalizedEmail}`);

        if (!sheet) {
          Logger.log("Sheet not found, creating it");
          const newSheet = ss.insertSheet(SUBSCRIBERS_SHEET);
          newSheet.appendRow([
            "Email",
            "Name",
            "Status",
            "JoinDate",
            "LastEmailSent",
            "LastEmailOpenTime",
          ]);

          const now = new Date();
          newSheet.appendRow([normalizedEmail, name, "Active", now, "", ""]);

          return ContentService.createTextOutput(
            JSON.stringify({
              success: true,
              message: "Subscription successful",
            })
          ).setMimeType(ContentService.MimeType.JSON);
        }

        // Get all data
        const data = sheet.getDataRange().getValues();
        if (data.length === 0) {
          // Empty sheet, add headers
          sheet.appendRow([
            "Email",
            "Name",
            "Status",
            "JoinDate",
            "LastEmailSent",
            "LastEmailOpenTime",
          ]);

          const now = new Date();
          sheet.appendRow([normalizedEmail, name, "Active", now, "", ""]);

          return ContentService.createTextOutput(
            JSON.stringify({
              success: true,
              message: "Subscription successful",
            })
          ).setMimeType(ContentService.MimeType.JSON);
        }

        // Get column positions
        const headers = data[0];
        const emailCol = headers.indexOf("Email");
        const nameCol = headers.indexOf("Name");
        const statusCol = headers.indexOf("Status");
        const joinDateCol = headers.indexOf("JoinDate");

        // Find the email or add new
        let found = false;
        for (let i = 1; i < data.length; i++) {
          const rowEmail = String(data[i][emailCol]).toLowerCase().trim();

          if (rowEmail === normalizedEmail) {
            found = true;
            const rowNum = i + 1;

            // Get current status
            const currentStatus = data[i][statusCol];
            Logger.log(
              `Found existing email at row ${rowNum} with status "${currentStatus}"`
            );

            // Always set to Active regardless of current status
            sheet.getRange(rowNum, statusCol + 1).setValue("Active");
            Logger.log("Status updated to Active");

            // Update name if provided
            if (name && nameCol !== -1) {
              sheet.getRange(rowNum, nameCol + 1).setValue(name);
              Logger.log(`Name updated to "${name}"`);
            }

            return ContentService.createTextOutput(
              JSON.stringify({
                success: true,
                message: "Subscription updated",
                previousStatus: currentStatus,
              })
            ).setMimeType(ContentService.MimeType.JSON);
          }
        }

        // Not found, add new
        if (!found) {
          Logger.log("Email not found, adding new subscriber");
          const now = new Date();
          const newRow = [];

          // Build a complete row with the right number of columns
          for (let i = 0; i < headers.length; i++) {
            if (i === emailCol) newRow[i] = normalizedEmail;
            else if (i === nameCol) newRow[i] = name;
            else if (i === statusCol) newRow[i] = "Active";
            else if (i === joinDateCol) newRow[i] = now;
            else newRow[i] = "";
          }

          // Add the row
          sheet.appendRow(newRow);

          return ContentService.createTextOutput(
            JSON.stringify({
              success: true,
              message: "Subscription successful",
              action: "added",
            })
          ).setMimeType(ContentService.MimeType.JSON);
        }
      } catch (error) {
        Logger.log(`Error in subscribe: ${error.toString()}`);
        return ContentService.createTextOutput(
          JSON.stringify({
            success: false,
            message: `Error: ${error.toString()}`,
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }

    default:
      return ContentService.createTextOutput("Invalid action");
  }
}

function unsubscribeUser(email) {
  try {
    Logger.log("==== UNSUBSCRIBE DEBUGGING ====");
    Logger.log("Unsubscribe request for email: " + email);

    // Update subscriber status
    const subscribers =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUBSCRIBERS_SHEET);
    Logger.log("Subscribers sheet name: " + SUBSCRIBERS_SHEET);
    Logger.log("Sheet found: " + (subscribers !== null));

    if (!subscribers) {
      Logger.log("ERROR: Subscribers sheet not found");
      return false;
    }

    const data = subscribers.getDataRange().getValues();
    Logger.log("Total rows in sheet (including header): " + data.length);

    const headers = data[0];
    Logger.log("Headers: " + headers.join(", "));

    const emailCol = headers.indexOf("Email");
    const statusCol = headers.indexOf("Status");

    Logger.log(
      "Column indexes - Email: " + emailCol + ", Status: " + statusCol
    );

    if (emailCol === -1 || statusCol === -1) {
      Logger.log("Required columns not found");
      return false;
    }

    // Find and update the subscriber
    let found = false;
    for (let i = 1; i < data.length; i++) {
      Logger.log(
        "Row " +
          (i + 1) +
          " - Email: '" +
          data[i][emailCol] +
          "', Status: '" +
          data[i][statusCol] +
          "'"
      );

      // Match is case-insensitive
      if (data[i][emailCol].toString().toLowerCase() === email.toLowerCase()) {
        Logger.log(
          "MATCH FOUND: Email found in row " +
            (i + 1) +
            " with status: '" +
            data[i][statusCol] +
            "'"
        );
        found = true;

        try {
          // Direct update of the cell
          subscribers.getRange(i + 1, statusCol + 1).setValue("Unsubscribed");
          Logger.log("Status updated to 'Unsubscribed'");

          // Verify the update actually happened
          const verifyData = subscribers.getDataRange().getValues();
          const verifyStatus = verifyData[i][statusCol];
          Logger.log(
            "Verification - Status after update: '" + verifyStatus + "'"
          );

          // Update analytics for the most recent template
          const sentHistory =
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
              SENT_HISTORY_SHEET
            );
          const historyData = sentHistory.getDataRange().getValues();
          let mostRecentTemplate = null;
          let mostRecentDate = new Date(0);

          for (let j = 1; j < historyData.length; j++) {
            if (
              historyData[j][0].toString().toLowerCase() ===
                email.toLowerCase() &&
              historyData[j][2] > mostRecentDate
            ) {
              mostRecentTemplate = historyData[j][1];
              mostRecentDate = historyData[j][2];
            }
          }

          if (mostRecentTemplate) {
            Logger.log(
              "Updating analytics for template: " + mostRecentTemplate
            );
            updateAnalytics(mostRecentTemplate, 0, 0, 0, 1);
          } else {
            Logger.log("No recent template found for this email");
          }
        } catch (updateError) {
          Logger.log("ERROR during status update: " + updateError.toString());
        }

        break;
      }
    }

    if (!found) {
      Logger.log("Email not found in subscribers list");
    }

    return found;
  } catch (error) {
    Logger.log("ERROR in unsubscribe process: " + error.toString());
    return false;
  }
}

function showEmailComposer() {
  const html = HtmlService.createHtmlOutputFromFile("EmailComposer")
    .setWidth(600)
    .setHeight(600)
    .setTitle("Compose Email");
  SpreadsheetApp.getUi().showModalDialog(html, "Compose Email");
}

function saveEmailTemplate(subject, body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    EMAIL_TEMPLATES_SHEET
  );
  const templateId = Utilities.getUuid();
  const now = new Date();

  sheet.appendRow([templateId, subject, body, now]);
  return templateId;
}

function showSendTestDialog(subject, body) {
  // Ensure subject and body are strings to prevent errors with replace()
  const safeSubject = subject ? subject.toString() : "";
  const safeBody = body ? body.toString() : "";

  // Escape single quotes to prevent JavaScript errors
  const escapedSubject = safeSubject.replace(/'/g, "\\'");
  const escapedBody = safeBody.replace(/'/g, "\\'");

  const htmlOutput = HtmlService.createHtmlOutput(
    `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <!-- Bootstrap CSS -->
          <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
          <!-- Bootstrap Icons -->
          <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.8.1/font/bootstrap-icons.css">
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
              padding: 20px;
            }
            .form-container {
              max-width: 100%;
            }
            .btn-send {
              min-width: 120px;
            }
          </style>
        </head>
        <body>
          <div class="container-fluid p-0">
            <h3 class="mb-3">
              <i class="bi bi-send me-2"></i>
              Send Test Email
            </h3>
            
            <div class="alert alert-info">
              <i class="bi bi-info-circle-fill me-2"></i>
              Test emails include all tracking features and unsubscribe links.
            </div>
            
            <div class="mb-3">
              <label for="testEmail" class="form-label">Recipient Email Address:</label>
              <input type="email" class="form-control" id="testEmail" placeholder="email@example.com" required>
            </div>
            
            <div class="card mb-3">
              <div class="card-header">
                <strong>Email Preview</strong>
              </div>
              <div class="card-body">
                <h6 class="card-subtitle mb-2 text-muted">Subject:</h6>
                <p class="card-text">${escapedSubject}</p>
              </div>
            </div>
            
            <div class="d-flex justify-content-between mt-4">
              <button class="btn btn-outline-secondary" onclick="google.script.host.close()">
                <i class="bi bi-x-circle me-1"></i>
                Cancel
              </button>
              <button class="btn btn-primary btn-send" onclick="sendTest()">
                <i class="bi bi-send-fill me-1"></i>
                Send Test
              </button>
            </div>
          </div>
          
          <!-- Bootstrap JS -->
          <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.bundle.min.js"></script>
          
          <script>
            function sendTest() {
              const email = document.getElementById('testEmail').value;
              if (!email) {
                showAlert('Please enter an email address', 'warning');
                return;
              }
              
              const sendBtn = document.querySelector('.btn-send');
              sendBtn.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Sending...';
              sendBtn.disabled = true;
              
              google.script.run
                .withSuccessHandler(() => {
                  showAlert('Test email sent successfully!', 'success');
                  setTimeout(() => google.script.host.close(), 1500);
                })
                .withFailureHandler(error => {
                  showAlert('Error: ' + error.message, 'danger');
                  sendBtn.innerHTML = '<i class="bi bi-send-fill me-1"></i>Send Test';
                  sendBtn.disabled = false;
                })
                .sendTestEmail(email, '${escapedSubject}', '${escapedBody}');
            }
            
            function showAlert(message, type) {
              // Remove any existing alerts
              const existingAlerts = document.querySelectorAll('.alert:not(.alert-info)');
              existingAlerts.forEach(alert => alert.remove());
              
              // Create new alert
              const alertDiv = document.createElement('div');
              alertDiv.className = 'alert alert-' + type + ' alert-dismissible fade show mt-3';
              alertDiv.innerHTML = message + 
                '<button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>';
              
              // Add to page
              document.querySelector('.container-fluid').appendChild(alertDiv);
            }
          </script>
        </body>
      </html>
    `
  )
    .setWidth(800)
    .setHeight(800);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Send Test Email");
}

function sendTestEmail(email, subject, body) {
  // Add tracking and unsubscribe links
  const finalBody = addTracking(body, email, "test-template");

  // Send the email
  GmailApp.sendEmail(
    email,
    subject,
    "This email requires HTML to view properly",
    {
      htmlBody: finalBody,
      name: NEWSLETTER_NAME,
    }
  );
}

function sendBulkEmail(templateId) {
  // Get the template
  const templates = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    EMAIL_TEMPLATES_SHEET
  );
  const templateData = templates.getDataRange().getValues();
  let subject, body;

  for (let i = 0; i < templateData.length; i++) {
    if (templateData[i][0] === templateId) {
      subject = templateData[i][1];
      body = templateData[i][2];
      break;
    }
  }

  if (!subject || !body) {
    throw new Error("Template not found");
  }

  // Get active subscribers
  const subscribers =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUBSCRIBERS_SHEET);
  const subscriberData = subscribers.getDataRange().getValues();
  const headers = subscriberData.shift(); // Remove header row

  // Find column indexes
  const emailCol = headers.indexOf("Email");
  const statusCol = headers.indexOf("Status");
  const lastSentCol = headers.indexOf("LastEmailSent");

  if (emailCol === -1 || statusCol === -1 || lastSentCol === -1) {
    throw new Error("Required columns not found in subscriber sheet");
  }

  // Prepare for bulk sending
  const now = new Date();
  const sentHistory =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SENT_HISTORY_SHEET);
  let sentCount = 0;

  // Process each subscriber
  for (let i = 0; i < subscriberData.length; i++) {
    const row = subscriberData[i];
    const email = row[emailCol];
    const status = row[statusCol];

    // Skip if not active
    if (status !== "Active") continue;

    // Create a hash of the message for this specific user
    const messageHash = Utilities.base64Encode(
      Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        email + templateId + now.getTime()
      )
    );

    // Add tracking and unsubscribe links
    const customBody = addTracking(body, email, templateId, messageHash);

    try {
      // Send the email
      GmailApp.sendEmail(
        email,
        subject,
        "This email requires HTML to view properly",
        {
          htmlBody: customBody,
          name: NEWSLETTER_NAME,
        }
      );

      // Update sent history
      sentHistory.appendRow([email, templateId, now, messageHash, "", 0]);

      // Update last sent date in subscriber sheet
      const rowNum = i + 2; // +2 because we removed headers and indexes start at 1
      subscribers.getRange(rowNum, lastSentCol + 1).setValue(now);

      sentCount++;
    } catch (error) {
      Logger.log("Error sending to " + email + ": " + error.toString());
      // Could add error handling here (e.g., mark as bounced)
    }

    // Sleep briefly to avoid rate limits
    if (i % 10 === 0) Utilities.sleep(1000);
  }

  // Update analytics
  updateAnalytics(templateId, sentCount, 0, 0, 0);

  return sentCount;
}

function confirmBulkSend() {
  // Get templates for selection
  const templates = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    EMAIL_TEMPLATES_SHEET
  );
  const templateData = templates.getDataRange().getValues();
  const headers = templateData.shift();

  let options = "";
  for (let i = 0; i < templateData.length; i++) {
    options += `<option value="${templateData[i][0]}">${templateData[i][1]}</option>`;
  }

  const html = `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <!-- Bootstrap CSS -->
          <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
          <!-- Bootstrap Icons -->
          <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.8.1/font/bootstrap-icons.css">
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
              padding: 20px;
            }
          </style>
        </head>
        <body>
          <div class="container-fluid p-0">
            <h3 class="mb-3">
              <i class="bi bi-send-check me-2"></i>
              Send to All Subscribers
            </h3>
            
            <div class="alert alert-warning">
              <i class="bi bi-exclamation-triangle-fill me-2"></i>
              This will send an email to <strong>all active subscribers</strong> in your database.
            </div>
            
            <div class="mb-4">
              <label for="templateSelect" class="form-label">Select Email Template:</label>
              <select id="templateSelect" class="form-select" required>
                <option value="" selected disabled>-- Select a template --</option>
                ${options}
              </select>
            </div>
            <div class="d-flex justify-content-between">
              <button class="btn btn-outline-secondary" onclick="google.script.host.close()">
                <i class="bi bi-x-circle me-1"></i>
                Cancel
              </button>
              <button class="btn btn-primary" onclick="sendToAll()">
                <i class="bi bi-send-fill me-1"></i>
                Confirm & Send
              </button>
            </div>
          </div>
          
          <!-- Bootstrap JS -->
          <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.bundle.min.js"></script>
          
          <script>
            function sendToAll() {
              const templateId = document.getElementById('templateSelect').value;
              if (!templateId) {
                showAlert('Please select a template', 'danger');
                return;
              }
              
              const sendBtn = document.querySelector('.btn-primary');
              const originalText = sendBtn.innerHTML;
              sendBtn.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Sending...';
              sendBtn.disabled = true;
              
              google.script.run
                .withSuccessHandler(result => {
                  showAlert('Email sent to ' + result + ' subscribers', 'success');
                  setTimeout(() => google.script.host.close(), 2000);
                })
                .withFailureHandler(error => {
                  showAlert('Error: ' + error.message, 'danger');
                  sendBtn.innerHTML = originalText;
                  sendBtn.disabled = false;
                })
                .sendBulkEmail(templateId);
            }
            
            function showAlert(message, type) {
              // Remove any existing alerts except the warning
              const existingAlerts = document.querySelectorAll('.alert:not(.alert-warning)');
              existingAlerts.forEach(alert => alert.remove());
              
              // Create new alert
              const alertDiv = document.createElement('div');
              alertDiv.className = 'alert alert-' + type + ' alert-dismissible fade show my-3';
              alertDiv.innerHTML = message + 
                '<button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>';
              
              // Add to page after the select
              const selectElement = document.getElementById('templateSelect');
              selectElement.parentNode.after(alertDiv);
            }
          </script>
        </body>
      </html>
    `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Send to All Subscribers");
}

function addTracking(body, email, templateId, messageHash = "") {
  // Add unsubscribe link
  const unsubscribeLink = `${WEB_APP_URL}?action=unsubscribe&email=${encodeURIComponent(
    email
  )}`;
  const unsubscribeText = `<div style="margin-top: 20px; padding-top: 10px; border-top: 1px solid #eee; font-size: 12px; color: #777;">
    If you no longer wish to receive these emails, <a href="${unsubscribeLink}">click here to unsubscribe</a>.
  </div>`;

  // Add tracking pixel
  const trackingPixel = `<img src="${WEB_APP_URL}?action=open&email=${encodeURIComponent(
    email
  )}&template=${templateId}&hash=${messageHash}" width="1" height="1" alt="" style="display:none">`;

  // Replace any links with tracking links
  let trackedBody = body.replace(
    /<a\s+(?:[^>]*?\s+)?href=(["'])(.*?)\1/gi,
    function (match, quote, url) {
      const trackingUrl = `${WEB_APP_URL}?action=click&email=${encodeURIComponent(
        email
      )}&template=${templateId}&hash=${messageHash}&url=${encodeURIComponent(
        url
      )}`;
      return `<a href="${trackingUrl}" target="_blank"`;
    }
  );

  // Add pixel and unsubscribe text to the end
  return trackedBody + trackingPixel + unsubscribeText;
}

function updateAnalytics(
  templateId,
  sent = 0,
  opens = 0,
  clicks = 0,
  unsubs = 0
) {
  const analytics =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANALYTICS_SHEET);
  const data = analytics.getDataRange().getValues();

  // Find if this template already has an entry
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === templateId) {
      rowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }

  if (rowIndex > 0) {
    // Update existing row
    const currentSent = data[rowIndex - 1][1] || 0;
    const currentOpens = data[rowIndex - 1][2] || 0;
    const currentClicks = data[rowIndex - 1][3] || 0;
    const currentUnsubs = data[rowIndex - 1][4] || 0;

    analytics.getRange(rowIndex, 2).setValue(currentSent + sent);
    analytics.getRange(rowIndex, 3).setValue(currentOpens + opens);
    analytics.getRange(rowIndex, 4).setValue(currentClicks + clicks);
    analytics.getRange(rowIndex, 5).setValue(currentUnsubs + unsubs);
  } else {
    // Add new row
    analytics.appendRow([templateId, sent, opens, clicks, unsubs]);
  }
}

function trackOpen(email, templateId, messageHash) {
  try {
    // Update sent history
    const sentHistory =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SENT_HISTORY_SHEET);
    const data = sentHistory.getDataRange().getValues();
    const now = new Date();

    // Find the right row to update
    for (let i = 1; i < data.length; i++) {
      if (
        data[i][0] === email &&
        data[i][1] === templateId &&
        data[i][3] === messageHash
      ) {
        // Only update if not already opened
        if (!data[i][4]) {
          sentHistory.getRange(i + 1, 5).setValue(now); // Set open time

          // Update subscriber's last open time
          updateSubscriberLastOpen(email, now);

          // Update analytics
          updateAnalytics(templateId, 0, 1, 0, 0);
        }
        break;
      }
    }
  } catch (error) {
    Logger.log("Error tracking open: " + error.toString());
  }
}

function trackClick(email, templateId, messageHash) {
  try {
    // Update sent history
    const sentHistory =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SENT_HISTORY_SHEET);
    const data = sentHistory.getDataRange().getValues();

    // Find the right row to update
    for (let i = 1; i < data.length; i++) {
      if (
        data[i][0] === email &&
        data[i][1] === templateId &&
        data[i][3] === messageHash
      ) {
        // Increment click count
        const currentClicks = data[i][5] || 0;
        sentHistory.getRange(i + 1, 6).setValue(currentClicks + 1);

        // Update analytics
        updateAnalytics(templateId, 0, 0, 1, 0);
        break;
      }
    }
  } catch (error) {
    Logger.log("Error tracking click: " + error.toString());
  }
}

function unsubscribeUser(email) {
  try {
    // Update subscriber status
    const subscribers =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUBSCRIBERS_SHEET);
    const data = subscribers.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf("Email");
    const statusCol = headers.indexOf("Status");

    if (emailCol === -1 || statusCol === -1) {
      Logger.log("Required columns not found");
      return;
    }

    // Find and update the subscriber
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email) {
        subscribers.getRange(i + 1, statusCol + 1).setValue("Unsubscribed");

        // Update analytics for the most recent template
        const sentHistory =
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
            SENT_HISTORY_SHEET
          );
        const historyData = sentHistory.getDataRange().getValues();
        let mostRecentTemplate = null;
        let mostRecentDate = new Date(0);

        for (let j = 1; j < historyData.length; j++) {
          if (
            historyData[j][0] === email &&
            historyData[j][2] > mostRecentDate
          ) {
            mostRecentTemplate = historyData[j][1];
            mostRecentDate = historyData[j][2];
          }
        }

        if (mostRecentTemplate) {
          updateAnalytics(mostRecentTemplate, 0, 0, 0, 1);
        }

        break;
      }
    }
  } catch (error) {
    Logger.log("Error unsubscribing user: " + error.toString());
  }
}

function updateSubscriberLastOpen(email, time) {
  const subscribers =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUBSCRIBERS_SHEET);
  const data = subscribers.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf("Email");
  const lastOpenCol = headers.indexOf("LastEmailOpenTime");

  if (emailCol === -1 || lastOpenCol === -1) {
    Logger.log("Required columns not found");
    return;
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] === email) {
      subscribers.getRange(i + 1, lastOpenCol + 1).setValue(time);
      break;
    }
  }
}

function showAnalytics() {
  // Create the HTML output from the Analytics HTML file
  const htmlOutput = HtmlService.createHtmlOutputFromFile("Analytics")
    .setWidth(1000) // Maximum possible width
    .setHeight(800) // Maximum possible height
    .setTitle("Email Analytics");

  // Use showModalDialog for larger size
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Email Analytics");
}

function getAnalyticsData() {
  // Get subscriber data
  const subscribers =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUBSCRIBERS_SHEET);
  const subData = subscribers.getDataRange().getValues();
  const subHeaders = subData.shift();
  const statusCol = subHeaders.indexOf("Status");

  let totalSubscribers = subData.length;
  let activeSubscribers = 0;

  for (let i = 0; i < subData.length; i++) {
    if (subData[i][statusCol] === "Active") {
      activeSubscribers++;
    }
  }

  // Get analytics data
  const analytics =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANALYTICS_SHEET);
  const analyticsData = analytics.getDataRange().getValues();
  const analyticsHeaders = analyticsData.shift();

  // Get template data for names
  const templates = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    EMAIL_TEMPLATES_SHEET
  );
  const templateData = templates.getDataRange().getValues();
  const templateHeaders = templateData.shift();

  const templateMap = {};
  for (let i = 0; i < templateData.length; i++) {
    templateMap[templateData[i][0]] = templateData[i][1]; // Map ID to subject
  }

  // Calculate metrics
  let totalSent = 0;
  let totalOpens = 0;
  let totalClicks = 0;
  const campaigns = [];

  for (let i = 0; i < analyticsData.length; i++) {
    const templateId = analyticsData[i][0];
    const sent = analyticsData[i][1] || 0;
    const opens = analyticsData[i][2] || 0;
    const clicks = analyticsData[i][3] || 0;
    const unsubs = analyticsData[i][4] || 0;

    totalSent += sent;
    totalOpens += opens;
    totalClicks += clicks;

    campaigns.push({
      id: templateId,
      name: templateMap[templateId] || "Unknown Template",
      sent: sent,
      opens: opens,
      clicks: clicks,
      unsubscribes: unsubs,
    });
  }

  // Calculate averages
  const averageOpenRate = totalSent > 0 ? totalOpens / totalSent : 0;
  const averageClickRate = totalOpens > 0 ? totalClicks / totalOpens : 0;

  return {
    totalSubscribers: totalSubscribers,
    activeSubscribers: activeSubscribers,
    averageOpenRate: averageOpenRate,
    averageClickRate: averageClickRate,
    campaigns: campaigns,
  };
}

function addSubscriber(email, name) {
  try {
    Logger.log("Adding subscriber - Email: " + email + ", Name: " + name);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SUBSCRIBERS_SHEET);

    if (!sheet) {
      Logger.log("Sheet not found: " + SUBSCRIBERS_SHEET);
      throw new Error("Subscribers sheet not found");
    }

    const now = new Date();

    // Verify we can write to the sheet by getting the sheet's last row
    const lastRow = sheet.getLastRow();
    Logger.log("Last row in sheet: " + lastRow);

    // Try to append the new row
    sheet.appendRow([email, name, "Active", now, "", ""]);

    Logger.log("Subscriber added successfully");
    return true;
  } catch (error) {
    Logger.log("Error in addSubscriber: " + error.toString());
    throw error;
  }
}

function createSignupForm() {
  // This would create a Google Form linked to your sheet
  // Left as an extension point
}

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    {
      name: SUBSCRIBERS_SHEET,
      headers: [
        "Email",
        "Name",
        "Status",
        "JoinDate",
        "LastEmailSent",
        "LastEmailOpenTime",
      ],
    },
    {
      name: EMAIL_TEMPLATES_SHEET,
      headers: ["TemplateID", "Subject", "Body", "CreatedDate"],
    },
    {
      name: SENT_HISTORY_SHEET,
      headers: [
        "Email",
        "TemplateID",
        "SentDate",
        "MessageHash",
        "OpenTime",
        "Clicks",
      ],
    },
    {
      name: ANALYTICS_SHEET,
      headers: [
        "TemplateID",
        "SentCount",
        "OpenCount",
        "ClickCount",
        "UnsubCount",
      ],
    },
  ];

  // Track if we created any sheets
  let sheetsCreated = false;

  // Check and create each required sheet
  requiredSheets.forEach((sheet) => {
    // Try to get the sheet by name
    let existingSheet = ss.getSheetByName(sheet.name);

    if (!existingSheet) {
      // Create the sheet if it doesn't exist
      existingSheet = ss.insertSheet(sheet.name);

      // Add headers
      existingSheet
        .getRange(1, 1, 1, sheet.headers.length)
        .setValues([sheet.headers]);

      // Format headers
      existingSheet
        .getRange(1, 1, 1, sheet.headers.length)
        .setFontWeight("bold")
        .setBackground("#f3f3f3");

      // Freeze header row
      existingSheet.setFrozenRows(1);

      // Auto-resize columns
      existingSheet.autoResizeColumns(1, sheet.headers.length);

      sheetsCreated = true;
    }
  });

  // Show success message if sheets were created
  if (sheetsCreated) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      NEWSLETTER_NAME + " Setup",
      "Required sheets have been created and initialized.",
      ui.ButtonSet.OK
    );
  }

  return sheetsCreated;
}

function showSystemSettings() {
  const htmlOutput = HtmlService.createHtmlOutput(
    `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <!-- Bootstrap CSS -->
          <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
          <!-- Bootstrap Icons -->
          <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.8.1/font/bootstrap-icons.css">
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
              padding: 20px;
            }
          </style>
        </head>
        <body>
          <div class="container-fluid p-0">
            <h3 class="mb-3">
              <i class="bi bi-gear-fill me-2"></i>
              System Settings
            </h3>
            
            <div class="card mb-3">
              <div class="card-body">
                <h5 class="card-title">Web App URL</h5>
                <p class="card-text text-muted">Used for tracking and unsubscribe functionality</p>
                <div class="input-group mb-3">
                  <input type="text" class="form-control" id="webAppUrl" value="${WEB_APP_URL}">
                  <button class="btn btn-outline-primary" type="button" onclick="updateWebAppUrl()">
                    Update
                  </button>
                </div>
                <small class="text-muted">
                  <i class="bi bi-info-circle me-1"></i>
                  To get your Web App URL, deploy this script as a web app and paste the URL here.
                </small>
              </div>
            </div>
            
            <div class="card mb-3">
              <div class="card-body">
                <h5 class="card-title">Sender Settings</h5>
                <div class="mb-3">
                  <label for="senderName" class="form-label">Sender Name</label>
                  <input type="text" class="form-control" id="senderName" value="Your Newsletter Name">
                </div>
                <button class="btn btn-outline-primary" onclick="updateSenderName()">
                  Save Sender Name
                </button>
              </div>
            </div>
            
            <div class="mt-4 text-center">
              <p class="text-muted">
                <strong>Email System</strong> v1.0<br>
                <small>Powered by Google Apps Script</small>
              </p>
            </div>
          </div>
          
          <!-- Bootstrap JS -->
          <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/js/bootstrap.bundle.min.js"></script>
          
          <script>
            function updateWebAppUrl() {
              const url = document.getElementById('webAppUrl').value;
              if (!url || !url.startsWith('https://')) {
                showAlert('Please enter a valid HTTPS URL', 'danger');
                return;
              }
              
              google.script.run
                .withSuccessHandler(() => showAlert('Web App URL updated successfully', 'success'))
                .withFailureHandler(error => showAlert('Error: ' + error.message, 'danger'))
                .updateWebAppUrl(url);
            }
            
            function updateSenderName() {
              const name = document.getElementById('senderName').value;
              if (!name) {
                showAlert('Please enter a sender name', 'danger');
                return;
              }
              
              google.script.run
                .withSuccessHandler(() => showAlert('Sender name updated successfully', 'success'))
                .withFailureHandler(error => showAlert('Error: ' + error.message, 'danger'))
                .updateSenderName(name);
            }
            
            function showAlert(message, type) {
              // Remove any existing alerts
              const existingAlerts = document.querySelectorAll('.alert');
              existingAlerts.forEach(alert => alert.remove());
              
              // Create new alert
              const alertDiv = document.createElement('div');
              alertDiv.className = 'alert alert-' + type + ' alert-dismissible fade show mt-3';
              alertDiv.innerHTML = message + 
                '<button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>';
              
              // Add to page
              document.querySelector('.container-fluid').appendChild(alertDiv);
            }
          </script>
        </body>
      </html>
    `
  )
    .setWidth(500)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "System Settings");
}
