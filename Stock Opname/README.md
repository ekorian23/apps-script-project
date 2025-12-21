# Stock Opname – Employee Asset Self Reporting System

Stock Opname is an internal web-based self-reporting system built with 
Google Apps Script and Google Sheets to support annual employee asset verification.

This system allows employees to submit their assigned asset information 
directly through a web form, ensuring asset data accuracy and completeness.

---

## 🎯 Purpose
- Support annual stock opname activities
- Allow employees to self-report assigned assets
- Ensure asset data consistency between employees and IT records
- Reduce manual follow-ups by IT Operations

---

## 🚀 Features

### Employee Data Integration
- Fetch employee names dynamically from `data_employee` sheet
- Display employee details automatically:
  - NIP
  - Position
  - Department
  - Work Location
  - Email

### Asset Selection
- Fetch asset codes from `raw_data_asset` sheet
- Searchable asset dropdown using Select2
- Manual asset code input if asset is not listed
- Support multiple additional asset codes

### Asset Identification
- Input fields for:
  - Laptop serial number
  - MAC address
- Includes guidance for checking serial number and MAC address
  (Windows & macOS instructions provided)

### Data Submission
- Store submitted data into `Form_Response` sheet
- Automatically creates response sheet if not exists
- Timestamped submission for audit tracking
- Loading indicator and form validation handling

---

## 📊 Data Structure

### Source Sheets
- `data_employee`
- `raw_data_asset`

### Response Sheet
- `Form_Response`

Columns:
- Timestamp  
- Nama Lengkap  
- NIP  
- Posisi  
- Departemen  
- Work Location  
- Email  
- Kode Asset  
- Kode Asset Manual  
- Mac Address  
- Serial Number  
- Kode Asset Lainnya  

---

## 🔄 System Flow

1. Employee opens Stock Opname web form
2. Employee selects their name from dropdown
3. System automatically displays employee details
4. Employee selects assigned asset code
5. Employee inputs serial number and MAC address
6. Optional: input additional asset codes
7. Data is submitted and stored in Google Sheets
8. Confirmation message is shown to the employee

---

## 🛠 Tech Stack
- Google Apps Script (Backend)
- HTML, CSS, JavaScript
- Bootstrap 5
- Select2 (Searchable dropdown)
- Google Sheets (Data Storage)

---

## 📌 Deployment
1. Open project in Google Apps Script
2. Ensure required sheets exist:
   - `data_employee`
   - `raw_data_asset`
3. Deploy as **Web App**
4. Set access permissions for employees

---

## 👤 Author
Eko Rian  
IT Operations & Automation Specialist
