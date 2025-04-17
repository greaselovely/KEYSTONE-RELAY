# Email System with Google Sheets

A comprehensive email marketing system built with Google Sheets and Apps Script that provides subscriber management, email sending, and engagement analytics.

## Overview

This system allows you to manage email subscribers, send HTML emails, track opens and clicks, and view analytics - all using Google Sheets as the database and Google Apps Script as the backend.

To implement this email system, users must have access to Google Workspace (formerly G Suite) or a personal Google account with access to Google Sheets and Google Apps Script. The system relies on Google's infrastructure, including Sheets for database storage, Apps Script for backend processing, and Gmail API for sending emails. Note that free Gmail accounts have stricter daily sending limits (100 emails per day), while Google Workspace accounts typically have higher quotas based on your subscription tier. Additionally, the web app deployment requires administrative permissions to publish web apps that can be accessed by others. For organizations with strict security policies, an administrator may need to approve the script's access to Google services. While the core functionality works with free Google accounts, organizations with larger subscriber lists or higher security requirements should consider using a paid Google Workspace subscription for better performance and support.

## Features

- **Subscriber Management**
  - Add and remove subscribers
  - Track subscriber status (Active/Unsubscribed)
  - Public web page for subscription management

- **Email Composition**
  - HTML email creation with rich formatting
  - Preview emails before sending
  - Test emails to specific addresses

- **Email Sending**
  - Send to all active subscribers
  - Rate-limited sending to avoid quota issues
  - Custom sender name

- **Tracking and Analytics**
  - Open tracking via tracking pixels
  - Click tracking via link redirection
  - Unsubscribe tracking
  - Analytics dashboard with key metrics

## Setup Instructions

### 1. Create a Google Sheet with these tabs:
- "Subscribers" (columns: Email, Name, Status, JoinDate, LastEmailSent, LastEmailOpenTime)
- "EmailTemplates" (columns: TemplateID, Subject, Body, CreatedDate)
- "SentHistory" (columns: Email, TemplateID, SentDate, MessageHash, OpenTime, Clicks)
- "Analytics" (columns: TemplateID, SentCount, OpenCount, ClickCount, UnsubCount)

### 2. Set Up Google Apps Script
1. In your Google Sheet, go to Extensions > Apps Script
2. Paste the main script (main.gs) as well as constants.gs. 
3. Create additional HTML files as needed:
   - Create Subscription.html
   - Create EmailComposer.html
   - Create Analytics.html

### 3. Create a Constants File
1. Create a new script file named constants.gs
2. Add the constants for sheet names and configuration:
```javascript
// Sheet names
const SUBSCRIBERS_SHEET = "Subscribers";
const EMAIL_TEMPLATES_SHEET = "EmailTemplates";
const SENT_HISTORY_SHEET = "SentHistory";
const ANALYTICS_SHEET = "Analytics";
const NEWSLETTER_NAME = "Your Newsletter Name";
// Base URL for your deployed web app
const WEB_APP_URL = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";
```

### 4. Deploy the Web App
1. Click Deploy > New Deployment
2. Select type as "Web App"
3. Set "Execute as" to "Me" (important for proper spreadsheet access)
4. Set "Who has access" to "Anyone" or "Anyone with Google Account"
5. Click Deploy
6. Copy the deployment URL and update your constants.gs file with this URL

### 5. Set Up the Subscription Page
1. Ensure your Subscription.html file is properly set up
2. Access the subscription page at your web app URL
3. Users can subscribe and unsubscribe through this page

## Usage

### Managing Subscribers
- Subscribers are automatically added when they use the subscription page
- You can manually add subscribers directly to the Subscribers sheet
- Status is automatically updated when users subscribe or unsubscribe

### Creating Emails
1. From your Google Sheet, use the Email System menu
2. Select "Compose New Email"
3. Enter a subject and compose the email body using HTML
4. Save the template when done

### Sending Emails
1. To test an email, select "Send Test Email" from the menu
2. Enter a test recipient email and send
3. To send to all subscribers, select "Send to All Subscribers"
4. Choose a template and confirm sending

### Viewing Analytics
1. Select "View Analytics Dashboard" from the menu
2. View open rates, click rates, and other metrics
3. Analytics are automatically updated as emails are opened and links clicked

## Limitations

- Google Apps Script has a daily limit of 100 emails on free accounts
- The script uses MailApp/GmailApp which has sending limits
- For larger lists, consider integrating with an external email service API

## Troubleshooting

### Common Issues
- **Subscribers not receiving emails**: Check their status in the sheet, ensure it's "Active"
- **Tracking not working**: Verify your web app URL is correct in constants.gs
- **Permission errors**: Make sure the web app is deployed with "Execute as: Me"

### For Developers
- Check the Apps Script logs (View > Logs) for debugging information
- The script includes extensive logging to help diagnose issues
- For critical issues, use the emergency fix functions included in the script

## System Architecture
- Google Sheet serves as the database
- Apps Script provides the backend processing
- HTML templates provide the user interface
- Web app endpoints handle subscription and tracking

