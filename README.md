# COI Compliance Automation System

Excel/VBA automation system for insurance compliance tracking — batched Outlook email drafting, multi-folder document scanning, and live ERP data population via ADODB/ODBC SQL Server connection.

---

## Overview

In property services, every work order requires up-to-date insurance certificates (COI, Workers' Comp, and Hold Harmless agreements) before a job can proceed. Managing these manually — copying data into email templates, hunting for files across multiple folders, and tracking document status — is a significant time sink at high volume.

This system was built from scratch, without direction, to automate that entire workflow inside Excel. It handles email drafting, document attachment, folder scanning, status tracking, and ERP data retrieval — reducing per-work-order processing time from ~10 minutes to ~1-2 minutes.

At current production volume (~145 work orders per week), the system saves an estimated 15-20 hours of manual processing time per week.

---

## What It Does

### Email Dispatch (`email_dispatch.bas`)
Batch-processes selected work orders to draft or send two types of emails via Outlook:
- **Broker request emails** — requests COI and Workers' Comp certificates from the appropriate broker, dynamically routing to a secondary broker if no Workers' Comp file is found locally
- **Client delivery emails** — sends the collected insurance documents to the client, building management, and sales rep, with conditional logic for Hold Harmless agreement handling and management CC behavior

Both email types auto-populate all fields (recipient, subject, body, attachments) from the spreadsheet row with no manual copy-pasting.

### Folder Scanner (`folder_scan.bas`)
Scans local document folders for COI, Workers' Comp, and Hold Harmless files, deduplicates against existing tracker tables using `Application.Match`, and logs received dates. Companion `Update` functions cross-reference the main work order table and automatically stamp status fields when matching documents are found.

### ERP Auto-Populate (`odbc_populate.bas`)
On work order number entry, fires a `Worksheet_Change` event that opens a live ADODB connection to a SQL Server ERP database (SAMPro) via ODBC and executes a 7-table JOIN to retrieve appointment date, site address, company, client name, client email, and assigned technician — populating the row instantly with no manual lookup.

---

## Tech Stack

- **Excel VBA** — automation, event handling, UI logic
- **ADODB / ODBC** — live SQL Server connection
- **SQL** — 7-table JOIN query against ERP database
- **Outlook Object Model** — programmatic email drafting and sending
- **Windows FileSystem (Dir, FileDateTime)** — folder scanning and file matching

---

## Architecture

The system is built as three cooperating modules inside a single Excel workbook. The main table (`COIRequest`) on Sheet 1 is the source of truth — rows represent active work orders. Module 3 populates rows automatically on WO# entry via ERP lookup. Module 2 scans local folders and updates document status columns. Module 1 reads completed rows to batch-draft and send emails with the correct attachments. Sheets 4, 5, and 6 hold secondary tracking tables for COI, Workers' Comp, and Hold Harmless documents respectively.

---

## Notes

- All credentials, email addresses, file paths, and company-identifying information have been replaced with placeholder constants for privacy
- The `archive_deprecated.bas` module preserves earlier single-row, active-cell-based iterations of the email functions, included to document the evolution toward the current batch-selection pattern
- Built for and deployed in a live production environment