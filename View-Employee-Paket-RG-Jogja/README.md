# View Employee Paket RG Jogja – Package Status Viewer

View Employee Paket RG Jogja is a read-only web module built with Google Apps Script 
to allow employees to view the status of packages addressed to them.

This module is part of the Paket RG Jogja ecosystem and focuses on transparency, 
self-service access, and reducing repetitive inquiries to IT and office support.

---

## 🎯 Purpose
- Provide employees with self-service access to package status
- Reduce manual package status inquiries to IT / Security
- Improve transparency of package delivery at HQ Jogja
- Highlight pending packages that require immediate pickup

---

## 🚀 Features

### Employee Package View
- Displays package data sourced from the **`View Only`** sheet
- Columns displayed:
  - Recipient Name
  - Receive Date
  - Courier
  - Status (Pending / Done)

### Status Visualization
- Pending packages are visually highlighted
- Status badge indicators (Pending / Done)
- Warning message for packages that need to be picked up

### Search & Pagination
- Search by:
  - Recipient name
  - Courier
- Adjustable page size (10 / 20 / 50)
- Client-side pagination for performance

### Data Handling
- Data is automatically sorted so the **latest packages appear first**
- Date formatting handled dynamically
- Graceful fallback for invalid or empty data

### Auto Refresh
- Package data automatically refreshes every **2 minutes**
- Ensures near real-time visibility without manual reload

---

## 📊 Data Source

### Spreadsheet
- Google Sheet ID is defined in the script
- Sheet name used:
  - `View Only`

### Required Columns (View Only)
| Column | Description |
|------|------------|
| A | Recipient Name |
| B | Receive Date |
| C | Courier |
| D | Status (Done / empty) |

---

## 🔄 System Flow
1. Employee opens the View Employee web page
2. System fetches data from the `View Only` sheet
3. Data is processed and sorted (latest first)
4. Package status is rendered in the table
5. Pending packages are highlighted automatically
6. Data refreshes periodically without user action

---

## 🛠 Tech Stack
- Google Apps Script (Backend)
- HTML, CSS, JavaScript
- Google Sheets (Data Source)

---

## 🔐 Access & Security
- View-only access (no data mutation)
- No authentication logic in frontend
- Intended for internal network usage

---

## 👤 Author
Eko Rian  
IT Operations & Automation Specialist

