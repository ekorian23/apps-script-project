# Mutasi IT to Snipe-IT – Asset & User Synchronization

Mutasi IT to Snipe-IT is a set of Google Apps Script automation scripts used to 
synchronize user and asset data from Google Sheets into the Snipe-IT asset 
management system via REST API.

This project is designed to support IT Operations by reducing manual work during 
user onboarding, asset mutation, check-in, and check-out processes.

---

## 🎯 Purpose
- Automate user creation in Snipe-IT
- Automate asset status updates from Google Sheets
- Handle asset check-in and check-out logic
- Reduce manual errors in IT asset management
- Support bulk operations using spreadsheet-driven workflows

---

## 📂 Included Scripts

### 1️⃣ insertNewUsers.gs
Creates new users in Snipe-IT based on data stored in Google Sheets.

#### Key Features
- Reads user data from `Result Push User` sheet
- Validates employee number format  
  (`NIP.xx.xx.xx.xxxxx` or `NIP2.xx.xx.xx.xxxxx`)
- Skips rows that are:
  - Not marked with `yes`
  - Already successfully created
- Sends user data to Snipe-IT `/users` endpoint
- Writes status results back to Google Sheets
- Color-coded status:
  - Green → success
  - Red → failed
- Handles API validation errors gracefully
- Built-in delay to prevent API rate limiting

#### Required Columns (Result Push User)
| Column | Description |
|------|------------|
| A | First Name |
| B | Last Name |
| C | Username |
| D | Employee Number |
| E | Email |
| G | Password |
| H | Notes |
| I | Update Flag (`yes`) |
| J | Status Message |
| K | Location ID (optional) |

---

### 2️⃣ updateAssets.gs
Updates asset status, ownership, and location in Snipe-IT using spreadsheet data.

#### Key Features
- Reads asset mutation data from `Result Push Asset` sheet
- Supports:
  - Asset status update
  - Asset check-in
  - Asset check-out to user or location
- Automatically:
  - Fetches Status ID by status name
  - Fetches User ID by username
  - Fetches Location ID by location name
- Prevents double updates by checking status column
- Writes success or failure results back to sheet
- Automatically checks in assets before checkout if already assigned

#### Required Columns (Result Push Asset)
| Column | Description |
|------|------------|
| A | Asset ID |
| B | Asset Tag |
| C | Status Name |
| D | Username |
| E | Location Name |
| F | Notes |
| H | Update Flag (`yes`) |
| I | Status Message |

---

## 🔄 System Flow

### User Creation Flow
1. IT fills user data in `Result Push User`
2. Mark update flag as `yes`
3. Script validates employee number
4. User data is sent to Snipe-IT API
5. Status is written back to the sheet

### Asset Mutation Flow
1. IT updates asset data in `Result Push Asset`
2. Mark update flag as `yes`
3. Script resolves status, user, and location IDs
4. Asset is checked in or checked out if required
5. Asset status and location are updated
6. Result is logged back to the sheet

---

## 🛠 Tech Stack
- Google Apps Script
- Google Sheets
- Snipe-IT REST API
- JSON-based API communication

---

## 🔐 Configuration
Before running the scripts, configure:

```js
var api_baseurl = 'https://your-snipeit-url/api/v1';
var api_bearer_token = 'Bearer YOUR_API_TOKEN';
