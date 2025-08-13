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
- Dashboard View  
- Batch & Project Allocation  
- Payment Verification  
- Admin Reports

---

## 📜 License
This project is created for **Northern Railways Training Programme** and is intended for internal or academic use.
