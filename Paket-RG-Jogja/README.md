# Paket RG Jogja – Package Management System

Paket RG Jogja is an internal web-based package management system built using 
Google Apps Script and Google Sheets to manage incoming and outgoing packages 
at Ruangguru Jogja (HQ Jogja).

The system replaces manual package tracking with a centralized dashboard, 
real-time status monitoring, SLA tracking, and photo-based documentation.

## Features
- Real-time package dashboard
- Auto-generated Package ID
- Package receiving and delivery tracking
- SLA monitoring (over 3 days)
- Photo documentation (receive & deliver)
- Search, filter, and pagination
- Responsive and mobile-friendly UI

## Package ID Format
RG-FirstName-Courier-Number

## System Flow
1. Package is received and registered via web form
2. System automatically generates a unique Package ID
3. Receive documentation is uploaded and stored in Google Drive
4. Package status is set to **Pending**
5. Package delivery is recorded with delivery date and photo
6. Status automatically changes to **Completed**
7. Dashboard updates automatically

## Tech Stack
- Google Apps Script (Backend)
- HTML, CSS, JavaScript (Frontend)
- Google Sheets (Database)
- Google Drive (Photo Storage)

## Use Case
- IT Operations
- Office Support
- Front Office Administration
- Internal Operational Monitoring

## Author
Eko Rian
