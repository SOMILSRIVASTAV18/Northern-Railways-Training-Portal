# Northern Railways Training Portal

A **web-based trainee registration and management system** built for the **Northern Railways Training Programme**.  
The portal automates trainee onboarding, email verification, batch & project allocation, payment tracking, and report generation — all powered by **Google Apps Script** integrated with **Google Sheets**.

---

## 📌 Features

- **User Registration & Email Verification**
  - Secure sign-up with email verification link.
  - Login with session management.

- **Batch & Project Management**
  - Automatic project code & batch code generation.
  - Project capacity limit handling with auto-next-batch allocation.

- **Payment Management**
  - Upload and verify payment transactions.
  - Admin approval system.

- **Dashboard & Reporting**
  - Export reports directly from Google Sheets.
  - Real-time trainee & payment data.

- **Document Generation**
  - Auto-generate confirmation PDFs and reports.
  - Google Docs & Sheets integration.

---

## 🛠️ Tech Stack

**Frontend:**
- HTML5, CSS3, JavaScript
- Bootstrap 5

**Backend:**
- Google Apps Script (JavaScript)
- HTMLService templates

**Database:**
- Google Sheets (via SpreadsheetApp API)

**Other Integrations:**
- Google Drive API (for storing trainee photos)
- Google MailApp (for sending verification & notification emails)
- Google Calendar (for adding batch schedules)

---

## 📂 Project Structure

```
/Code
 ├── code.gs                  # Main Google Apps Script backend
 ├── index.html               # Login page
 ├── dashboard.html           # User/Student dashboard
 ├── userForm.html            # Trainee details form
 ├── paymentVerification.html # Payment upload form
 ├── adminPage.html           # Admin reporting panel
 └── assets/                  # CSS, JS, and image files
```

---

## 🚀 Deployment

1. **Create Google Spreadsheet**  
   - Add required sheets: `Verifications`, `Users`, `UserInfo`, `Payments`.

2. **Set Up Google Apps Script**  
   - Open **Extensions → Apps Script** in Google Sheets.
   - Paste the `code.js` contents into the editor.
   - Update `SPREADSHEET_ID` with your sheet’s ID.

3. **Deploy Web App**  
   - Click **Deploy → New Deployment**.
   - Select **Web app**:
     - Execute as: **Me**
     - Access: **Anyone with the link**
   - Copy the deployment URL.

4. **Update HTML Files**  
   - Modify any paths or branding in the HTML files as needed.
   - Upload them to Apps Script via `HtmlService.createTemplateFromFile`.

---

## 📸 Screenshots 
- Login Page
  <img width="1781" height="850" alt="Screenshot 2025-08-13 233328" src="https://github.com/user-attachments/assets/90e49b7d-89e1-4a21-81e9-f7d5ba4fdbf3" />
  
- Dashboard View
  <img width="899" height="529" alt="Screenshot 2025-08-13 235342" src="https://github.com/user-attachments/assets/653ab06a-9d5d-40b8-a437-b61e86d40f90" />

- Batch & Project Allocation
  <img width="1508" height="337" alt="Screenshot 2025-08-13 234143" src="https://github.com/user-attachments/assets/9a12014d-6b4e-4883-a1b6-d865a890a9c3" />

- Payment Verification
  <img width="1364" height="756" alt="Screenshot 2025-08-13 235059" src="https://github.com/user-attachments/assets/050ac53e-da87-44a6-a4a4-70622fe9c686" />

- Admin Reports
  <img width="1049" height="591" alt="Screenshot 2025-08-13 235852" src="https://github.com/user-attachments/assets/c75a25ed-0dc9-4e00-aa97-ddf50c81adf2" />

  <img width="1836" height="886" alt="Screenshot 2025-08-14 000259" src="https://github.com/user-attachments/assets/bfaa0a6a-a599-4c37-ba2c-6878c6a0e640" />

---

## 📜 License
This project is created for **Northern Railways Training Programme** and is intended for internal or academic use.
