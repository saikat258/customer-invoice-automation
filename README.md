# customer-invoice-automation
This repo contains code in google form automation to take order from customer and then generate the pdf invoice in a mail to the user and also gives the seller a master data sheet that will contain the order details.
# Customer Information to Invoice Automation

This project automates the customer onboarding and invoicing workflow.
The customer has to fill this form
https://forms.gle/vxCJmKxCWSXZ6B2m9

## Overview

When a customer submits the form:
1. Their details are saved in the master Google Sheet.
- Google Sheet: stores the master customer list and invoice status
- Google Docs template: used to generate invoice content
- Google Apps Script: connects the form, sheet, docs, drive, and Gmail
Create a form with fields such as:
- Customer Name
- Email

### 3. Prepare the Sheet Headers
Use these column names in the first row:
- Status
- Invoice Sent Date

- `{{date}}`
- `{{email}}`
- `{{amount}}`
- `{{unitPrice}}`
- `{{total}}`

### 5. Add Apps Script

## Apps Script Function

The script runs on form submission and performs the following:
- emails the PDF to the customer
- updates the sheet with invoice details

## Apps Script Trigger
- Event type: `On form submit`


The admin should have access to:
- the Google Form
- the Google Sheet
- the Google Docs invoice template
- the Apps Script project

Recommended permissions:
- Editor access for the Sheet, Form, Doc, and Apps Script project

## Testing

4. Check that the customer receives the email

## Troubleshooting

- confirm the Google Docs template ID is correct
- ensure the script has permission to access Drive, Docs, Sheets, and Gmail

## Notes

This setup is ideal for:
- automatic customer information collection
- creating invoices without manual work
- storing a master record for future use
