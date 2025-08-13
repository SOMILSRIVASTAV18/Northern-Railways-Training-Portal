// Global variables
const SPREADSHEET_ID = '1J5cZcCjtX1DAp8nxiUcUtahK8XMtl5gOdFnZQSqMr0Q'; // Add your Google Sheet ID here
const VERIFICATION_SHEET = 'Verifications';
const USERS_SHEET = 'Users';
const USER_INFO_SHEET = 'UserInfo';
const PAYMENT_SHEET = 'Payments';
const MAX_STUDENTS_PER_PROJECT_BATCH = 15; // Define the maximum number of students per project per batch

// This template ID is now only used by createConfirmationPdf and createPaymentRequestPdfFromTemplate
// The generateBatchProjectDocReport function will now create a document from scratch. gene send
//const BATCH_PROJECT_DOC_TEMPLATE_ID = 'YOUR_GOOGLE_DOCS_TEMPLATE_ID_HERE'; // Keep for other functions if

// Initialize the spreadsheet and sheets
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Create sheets if they don't exist
  if (!ss.getSheetByName(VERIFICATION_SHEET)) {
    ss.insertSheet(VERIFICATION_SHEET)
      .appendRow(['Email', 'VerificationCode', 'Verified', 'Timestamp']);
  }

  if (!ss.getSheetByName(USERS_SHEET)) {
    ss.insertSheet(USERS_SHEET)
      .appendRow(['Email', 'Name', 'PasswordHash', 'Verified', 'Timestamp','IsAdmin']); // Added 'IsAdmin'
  }

  if (!ss.getSheetByName(USER_INFO_SHEET)) {
    // Added 'BatchCode', 'ProjectCode', 'StartDateOfBatch'
    ss.insertSheet(USER_INFO_SHEET)
      .appendRow(['Email', 'FullName', 'PhoneNumber', 'College', 'Address', 'PhotoUrl', 'UniqueId',
                  'DateOfBirth', 'Gender', 'Course', 'Branch', 'GuardianName', 'TrainingDuration',
                  'Project', 'CurrentYear', 'BatchCode', 'ProjectCode', 'StartDateOfBatch', 'Timestamp']);
  }

  if (!ss.getSheetByName(PAYMENT_SHEET)) {
    ss.insertSheet(PAYMENT_SHEET)
      .appendRow(['UniqueId', 'Email', 'TransactionId', 'Verified', 'Timestamp']);
  }

  return ss;
}

function doGet(e) {

  const userEmail = PropertiesService.getUserProperties().getProperty('userEmail');
  const isLoggedIn = userEmail && isUserVerified(userEmail);

  // Handle one-time actions like email verification first.
  if (e.parameter.page === 'verify') {
    const result = verifyEmail(e.parameter.email, e.parameter.code);
    const template = HtmlService.createTemplateFromFile('verification_result');
    template.result = result;
    return template.evaluate().setTitle('Email Verification');
  }

  if (e.parameter.page === 'adminLogin') {
    return HtmlService.createTemplateFromFile('adminLogin')
      .evaluate()
      .setTitle('Admin Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // --- NEW: Admin Page Route with robust authentication ---
  if (e.parameter.page === 'admin') {
    if (userEmail && isAdminUser(userEmail)) { // Check if user is logged in AND is an admin
      return HtmlService.createTemplateFromFile('adminPage')
        .evaluate()
        .setTitle('Admin Report Generator')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      // Redirect to admin login if not authenticated as admin
      return HtmlService.createTemplateFromFile('adminLogin') // Redirect to the new admin login page
        .evaluate()
        .setTitle('Admin Login Required')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }


  if (isLoggedIn) {
    // --- START: NEW ROUTING LOGIC for logged-in users ---
    let template;
    switch (e.parameter.page) {
      case 'userForm':
        template = HtmlService.createTemplateFromFile('userForm'); // Assuming your user form is named userForm.html
        template.userEmail = userEmail; // Pass email to pre-fill if needed
        return template.evaluate()
          .setTitle('Student Information Form')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'paymentVerification':
        template = HtmlService.createTemplateFromFile('paymentVerification'); // Assuming your payment form is named paymentVerification.html
        template.userEmail = userEmail; // Pass email to pre-fill if needed
        return template.evaluate()
          .setTitle('Payment Verification')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      default:
        // If no page is specified, or it's an unknown page, show the dashboard.
        template = HtmlService.createTemplateFromFile('dashboard');
        return template.evaluate()
          .setTitle('Student Dashboard')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    // --- END: NEW ROUTING LOGIC ---

  } else {
    // ALL non-logged-in users go to the login page.
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Northern Railways Training Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function getUserStatus() {
  const userEmail = PropertiesService.getUserProperties().getProperty('userEmail');
  if (!userEmail) {
    return { error: "User not logged in." };
  }

  const userInfoSubmitted = isUserInfoSubmitted(userEmail);
  const paymentVerified = isPaymentVerified(userEmail);

  let u = {};
  if(userInfoSubmitted){
    u = getu(userEmail); // Your existing function to get user details
  }

  return {
    isLoggedIn: true,
    userInfoSubmitted: userInfoSubmitted,
    paymentVerified: paymentVerified,
    u: u
  };
}


// Helper function to get user data from sheets
function getu(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);

    // Get user info
    const userInfoData = userInfoSheet.getDataRange().getValues();
    let u = {
      email: email,
      fullName: "",
      phoneNumber: "",
      college: "",
      address: "",
      duration: "",
      guardianName: "",
      course: "",
      project:"",
      uniqueId: "",
      photoUrl: "",
      batchCode: "", // Added
      projectCode: "", // Added
      startDateOfBatch: "" // Added
    };

    for (let i = 1; i < userInfoData.length; i++) {
      if (userInfoData[i][0] === email) {
        u.fullName = userInfoData[i][1];
        u.phoneNumber = userInfoData[i][2];
        u.college = userInfoData[i][3];
        u.address = userInfoData[i][4];
        u.guardianName = userInfoData[i][11];
        u.course = userInfoData[i][9];
        u.duration = userInfoData[i][12];
        u.project = userInfoData[i][13];
        u.photoUrl = userInfoData[i][5];
        u.uniqueId = userInfoData[i][6];
        u.batchCode = userInfoData[i][15];       // Column P
        u.projectCode = userInfoData[i][16];     // Column Q
        u.startDateOfBatch = userInfoData[i][17]; // Column R
        break;
      }
    }

    return u;
  } catch (error) {
    Logger.log('Error in getu: ' + error);
    return null;
  }
}

// Calculate end date based on start date and weeks
function getEndDate(startDate, weeks) {
  const endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + (weeks * 7 - 1));
  return endDate;
}

// Generate Google Calendar link
function generateCalendarLink(u, courseData) {
  const startDate = new Date(courseData.startDate);
  const endDate = new Date(courseData.endDate);

  // Format dates for Google Calendar
  const startFormatted = startDate.toISOString().replace(/-|:|\.\d\d\d/g, "");
  const endFormatted = endDate.toISOString().replace(/-|:|\.\d\d\d/g, "");

  const eventTitle = encodeURIComponent("Northern Railways Training: " + u.project);
  const eventDetails = encodeURIComponent("Registration ID: " + u.uniqueId + "\nLocation: " + courseData.location);

  return "https://www.google.com/calendar/render?action=TEMPLATE&text=" + eventTitle + "&dates=" + startFormatted + "/" + endFormatted + "&details=" + eventDetails + "&location=" + encodeURIComponent(courseData.location);
}

// Process user registration
function processRegistration(form) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(USERS_SHEET);
    const verificationSheet = ss.getSheetByName(VERIFICATION_SHEET);

    const email = form.email;
    const name = form.name;
    const password = form.password;

    // Check if user already exists
    const u = usersSheet.getDataRange().getValues();
    for (let i = 1; i < u.length; i++) {
      if (u[i][0] === email) {
        return { success: false, message: 'Email already registered' };
      }
    }

    // Generate password hash (simple hash for demo)
    const passwordHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));

    // Generate verification code
    const verificationCode = Utilities.getUuid();

    // Add user to Users sheet
    usersSheet.appendRow([email, name, passwordHash, false, new Date()]);

    // Add verification record
    verificationSheet.appendRow([email, verificationCode, false, new Date()]);

    // Send verification email
    sendVerificationEmail(email, name, verificationCode);

    return { success: true };
  } catch (error) {
    console.error('Registration error:', error);
    return { success: false, message: 'Registration failed: ' + error.toString() };
  }
}

// Send verification email
function sendVerificationEmail(email, name, verificationCode) {
  const subject = 'Verify Your Course Registration Account';
  const verificationUrl = ScriptApp.getService().getUrl() + '?page=verify&code=' + verificationCode + '&email=' + encodeURIComponent(email);

  const htmlBody = `
    <h2>Hello ${name},</h2>
    <p>Thank you for registering with our Northern Railways Training Portal.</p>
    <p>Please click the link below to verify your email address:</p>
    <p><a href="${verificationUrl}">Verify Email Address</a></p>
    <p>If you did not request this registration, please ignore this email.</p>
    <p>Regards,<br>Northern Railways</p>
  `;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody
  });
}


// Verify email with verification code
function verifyEmail(email, code) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const verificationSheet = ss.getSheetByName(VERIFICATION_SHEET);
    const usersSheet = ss.getSheetByName(USERS_SHEET);

    // Find verification record
    const verificationData = verificationSheet.getDataRange().getValues();
    let verificationRow = -1;

    for (let i = 1; i < verificationData.length; i++) {
      if (verificationData[i][0] === email && verificationData[i][1] === code) {
        verificationRow = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }

    if (verificationRow === -1) {
      return { success: false, message: 'Invalid verification link' };
    }

    // Update verification status
    verificationSheet.getRange(verificationRow, 3).setValue(true);

    // Update user verification status
    const u = usersSheet.getDataRange().getValues();
    for (let i = 1; i < u.length; i++) {
      if (u[i][0] === email) {
        usersSheet.getRange(i + 1, 4).setValue(true);
        break;
      }
    }

    return { success: true, message: 'Email verified successfully! You can now log in.' };
  } catch (error) {
    console.error('Verification error:', error);
    return { success: false, message: 'Verification failed: ' + error.toString() };
  }
}

// Process login
function processLogin(form) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(USERS_SHEET);

    const email = form.email;
    const password = form.password;

    const passwordHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));

    const u = usersSheet.getDataRange().getValues();
    let userFound = false;
    let isVerified = false;
    let isAdmin = false; // Initialize isAdmin flag

    for (let i = 1; i < u.length; i++) {
      if (u[i][0] === email) {
        userFound = true;
        isVerified = u[i][3];
        isAdmin = u[i][5]; // Get IsAdmin status from the 6th column (index 5)

        if (u[i][2] !== passwordHash) {
          return { success: false, message: 'Invalid password' };
        }

        if (!isVerified) {
          return { success: false, message: 'Email not verified. Please check your email for verification link.' };
        }

        break;
      }
    }

    if (!userFound) {
      return { success: false, message: 'Email not registered' };
    }

    PropertiesService.getUserProperties().setProperty('userEmail', email);
    CacheService.getUserCache().put('userEmail', email, 1800); // Cache for 0.5 hours

    console.log('Login successful for: ' + email);

    return { success: true, isAdmin: isAdmin }; // Return isAdmin status
  } catch (error) {
    console.error('Login error:', error);
    return { success: false, message: 'Login failed: ' + error.toString() };
  }
}

// Check if user is verified
function isUserVerified(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(USERS_SHEET);

    // Make sure we have a valid sheet
    if (!usersSheet) {
      console.error('Users sheet not found');
      return false;
    }

    // Define u here before using it
    const u = usersSheet.getDataRange().getValues();

    for (let i = 1; i < u.length; i++) {
      if (u[i][0] === email) {
        return u[i][3]; // Return verified status
      }
    }

    return false;
  } catch (error) {
    console.error('Verification check error:', error);
    return false;
  }
}

function testConnection() {
  return {
    success: true,
    message: "Server connection working",
    timestamp: new Date().toString()
  };
}

/**
 * Formats a Date object into a YYYYMMDD string.
 * @param {Date} date The date object to format.
 * @returns {string} The formatted date string.
 */
function formatDateToYYYYMMDD(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${day}-${month}`;
}

/**
 * Generates a batch code based on the start date.
 * Format: STVT/YYYYMMDD
 * @param {Date} startDate The start date of the batch.
 * @returns {string} The generated batch code.
 */
function generateBatchCode(startDate) {
  return `STVT/${formatDateToYYYYMMDD(startDate)}`;
}

/**
 * Generates a short version of the project name.
 * @param {string} projectName The full project name.
 * @returns {string} A shortened, sanitized project name.
 */
function getProjectShortName(projectName) {
  // Remove common suffixes like (4w), (6w), (8w)
  let shortName = projectName.replace(/\s*\(\d+w\)\s*$/, '');
  // Take first 3 words, or first 20 characters, remove non-alphanumeric
  shortName = shortName.split(' ').slice(0, 3).join(' ');
  shortName = shortName.replace(/[^a-zA-Z0-9]/g, '').substring(0, 20).toUpperCase();
  return shortName;
}

/**
 * Generates a project code based on the branch and project name.
 * Format: BranchShort/ProjectShortName
 * @param {string} branch The branch name (e.g., Computer Science).
 * @param {string} projectName The full project name.
 * @returns {string} The generated project code.
 */
function generateProjectCode(branch, projectName) {
  const branchShort = branch.replace(/[^a-zA-Z0-9]/g, '').substring(0, 5).toUpperCase(); // e.g., COMPU
  const projectShortName = getProjectShortName(projectName);
  return `${branchShort}/${projectShortName}`;
}


/**
 * Counts the number of students assigned to a specific project within a given batch.
 * @param {string} project The project name.
 * @param {string} batchCode The batch code.
 * @returns {number} The count of students.
 */
function getProjectCountInBatch(project, batchCode) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);
  if (!userInfoSheet) {
    Logger.log("getProjectCountInBatch failed: UserInfo sheet not found.");
    return 0;
  }
  const data = userInfoSheet.getDataRange().getValues();
  let count = 0;
  // Columns: Project (13), BatchCode (15)
  for (let i = 1; i < data.length; i++) { // Start from 1 to skip headers
    if (data[i][13] === project && data[i][15] === batchCode) {
      count++;
    }
  }
  return count;
}


/**
 * Returns a list of available batch start dates and project counts for each.
 * This function is called by the frontend to populate dropdowns and manage project availability.
 * @returns {Object} An object containing potential future Mondays and current project counts.
 */
function getAvailableBatchesAndProjects() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);
  const projectCounts = {}; // { 'YYYYMMDD': { 'Project Name': count, ... } }

  if (userInfoSheet) {
    const data = userInfoSheet.getDataRange().getValues();
    // Columns: Project (13), StartDateOfBatch (17)
    for (let i = 1; i < data.length; i++) {
      const project = data[i][13];
      const startDateRaw = data[i][17]; // This might be a Date object or string from sheet

      if (project && startDateRaw) {
        let startDate;
        // Ensure startDate is a Date object for consistent formatting
        if (startDateRaw instanceof Date) {
          startDate = startDateRaw;
        } else {
          // Attempt to parse string dates (e.g., "Mon Jul 15 2024 00:00:00 GMT+0530 (India Standard Time)")
          try {
            startDate = new Date(startDateRaw);
            // Basic validation to ensure it's a valid date
            if (isNaN(startDate.getTime())) {
              Logger.log(`Invalid date string in sheet: ${startDateRaw}`);
              continue;
            }
          } catch (e) {
            Logger.log(`Error parsing date string: ${startDateRaw}, Error: ${e}`);
            continue;
          }
        }

        const formattedDate = formatDateToYYYYMMDD(startDate);

        if (!projectCounts[formattedDate]) {
          projectCounts[formattedDate] = { date: startDate.toDateString(), projects: {} };
        }
        if (!projectCounts[formattedDate].projects[project]) {
          projectCounts[formattedDate].projects[project] = 0;
        }
        projectCounts[formattedDate].projects[project]++;
      }
    }
  }

  // Generate future Mondays for selection
  const futureMondays = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize to start of day

  // Find the next Monday from today
  let nextMonday = new Date(today);
  const day = nextMonday.getDay(); // 0 = Sunday, 1 = Monday, etc.
  const daysUntilNextMonday = day === 0 ? 1 : 8 - day;
  nextMonday.setDate(nextMonday.getDate() + daysUntilNextMonday);

  for (let i = 0; i < 12; i++) { // Generate for the next 12 weeks
    const currentMonday = new Date(nextMonday);
    futureMondays.push({
      dateString: currentMonday.toDateString(), // e.g., "Mon Jul 15 2024"
      formattedDate: formatDateToYYYYMMDD(currentMonday) // e.g., "20240715"
    });
    nextMonday.setDate(nextMonday.getDate() + 7); // Move to the next Monday
  }

  return { futureMondays: futureMondays, projectCounts: projectCounts };
}

// Process user information and photo upload
function processUserInfo(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);

    // Get user email from session
    const userEmail = PropertiesService.getUserProperties().getProperty('userEmail');
    if (!userEmail) {
      return { success: false, message: 'User not logged in' };
    }

    // Generate unique ID
    const uniqueId = generateUniqueId();

    // Save photo to Drive
    const photoUrl = savePhotoToDrive(formData.photoData, formData.photoName, formData.photoType, userEmail);

    // Parse the selected start date from the form
    let assignedStartDate = new Date(formData.startDateOfBatch);
    assignedStartDate.setHours(0, 0, 0, 0); // Normalize to start of day

    let assignedBatchCode = generateBatchCode(assignedStartDate);
    let assignedProjectCode = generateProjectCode(formData.branch, formData.project);

    // --- Project Capacity Logic ---
    let projectCount = getProjectCountInBatch(formData.project, assignedBatchCode);
    let attempts = 0;
    const MAX_ATTEMPTS = 15; // Limit attempts to find a new batch

    // If the selected project in the chosen batch is full, find the next available batch
    while (projectCount >= MAX_STUDENTS_PER_PROJECT_BATCH && attempts < MAX_ATTEMPTS) {
      Logger.log(`Project "${formData.project}" in batch "${assignedBatchCode}" is full (${projectCount} students). Trying next week.`);
      assignedStartDate.setDate(assignedStartDate.getDate() + 7); // Move to the next week
      assignedBatchCode = generateBatchCode(assignedStartDate);
      assignedProjectCode = generateProjectCode(formData.branch, formData.project); // Regenerate project code for new batch
      projectCount = getProjectCountInBatch(formData.project, assignedBatchCode); // Check new batch
      attempts++;
    }

    if (attempts >= MAX_ATTEMPTS) {
      Logger.log(`Could not find an available slot for project "${formData.project}" after ${MAX_ATTEMPTS} attempts.`);
      return { success: false, message: 'No available slots for this project in upcoming batches. Please try a different project or contact support.' };
    }

    Logger.log(`Assigned to batch "${assignedBatchCode}" with start date ${assignedStartDate.toDateString()}. Project count for this slot: ${projectCount}.`);

    // Add user info to sheet with new fields
    userInfoSheet.appendRow([
      userEmail,
      formData.fullName,
      formData.phoneNumber,
      formData.college,
      formData.address,
      photoUrl,
      uniqueId,
      formData.dateOfBirth,
      formData.gender,
      formData.course,
      formData.branch,
      formData.guardianName,
      formData.trainingDuration,
      formData.project,
      formData.currentYear,
      assignedBatchCode,         // Automatically generated BatchCode
      assignedProjectCode,       // Automatically generated ProjectCode
      assignedStartDate,         // Determined StartDateOfBatch (Date object)
      new Date() // Timestamp
    ]);

    // Send payment request email
    //sendPaymentRequestEmail(userEmail, formData.fullName, uniqueId);

    return { success: true, uniqueId: uniqueId, batchCode: assignedBatchCode, projectCode: assignedProjectCode, startDateOfBatch: assignedStartDate.toDateString() };
  } catch (error) {
    console.error('User info processing error:', error);
    return { success: false, message: 'Failed to process user information: ' + error.toString() };
  }
}
function isUserInfoSubmitted(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);
    const userInfoData = userInfoSheet.getDataRange().getValues();

    for (let i = 1; i < userInfoData.length; i++) {
      if (userInfoData[i][0] === email) {
        return true;
      }
    }

    return false;
  } catch (error) {
    console.error('User info check error:', error);
    return false;
  }
}
// Check if the user's payment is verified
function isPaymentVerified(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const paymentSheet = ss.getSheetByName(PAYMENT_SHEET);
    if (!paymentSheet) return false;

    const paymentData = paymentSheet.getDataRange().getValues();
    for (let i = 1; i < paymentData.length; i++) {
      // Corrected condition: Check for boolean TRUE.
      if (paymentData[i][1] === email && paymentData[i][3] === true) {
        return true;
      }
    }
    return false;
  } catch (error) {
    console.error('isPaymentVerified Error: ' + error);
    return false;
  }
}



// Logout function
function logout() {
  try {
    console.log('Starting logout process');
    PropertiesService.getUserProperties().deleteProperty('userEmail');
    CacheService.getUserCache().remove('userEmail');
    console.log('Logout successful');
    return { success: true };
  } catch (error) {
    console.error('Logout error:', error);
    return { success: false, message: error.toString() };
  }
}
function checkLoginStatus() {
  const userEmail = PropertiesService.getUserProperties().getProperty('userEmail');
  const cacheEmail = CacheService.getUserCache().get('userEmail');

  return {
    propertiesEmail: userEmail,
    cacheEmail: cacheEmail,
    isVerified: userEmail ? isUserVerified(userEmail) : false
  };
}
function getCurrentUserEmail() {
  return PropertiesService.getUserProperties().getProperty('userEmail');
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// Generate a unique ID
// Remove or comment out this global declaration: let currentSequentialNumber = 0;

function getLatestSequentialNumber() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // Open the spreadsheet by ID
  const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET); // Get the 'UserInfo' sheet

  if (!userInfoSheet) {
    Logger.log('UserInfo sheet not found. Returning 0 for sequential number.');
    return 0;
  }

  const uniqueIdColumnIndex = 6; // Column G (0-indexed)
  const data = userInfoSheet.getDataRange().getValues(); // Get all data from the sheet
  let maxSequentialNumber = 0; // Initialize with 0

  // Start from 1 to skip header row
  for (let i = 1; i < data.length; i++) {
    const uniqueId = data[i][uniqueIdColumnIndex]; // Get the UniqueId from column G

    if (uniqueId && typeof uniqueId === 'string' && uniqueId.startsWith('STVT/')) { // Check if it's a valid ID string
      // Extract the sequential part (e.g., "001" from "STVT/25/001")
      const parts = uniqueId.split('/'); // Split the ID by '/'
      if (parts.length === 3) { // Ensure it has all three parts
        const sequentialNumStr = parts[2]; // Get the last part
        const sequentialNum = parseInt(sequentialNumStr, 10); // Convert to integer

        if (!isNaN(sequentialNum)) { // Check if it's a valid number
          if (sequentialNum > maxSequentialNumber) { // If it's greater than the current max, update
            maxSequentialNumber = sequentialNum;
          }
        }
      }
    }
  }
  Logger.log(`Latest sequential number found: ${maxSequentialNumber}`);
  return maxSequentialNumber;
}

function generateUniqueId() {
  const now = new Date(); // Get current date and time
  const yearCode = now.getFullYear().toString().slice(-2); // Get last two digits of the year

  // Initialize currentSequentialNumber with the latest value from the sheet
  // This ensures it's always starting from the correct base.
  let currentSequentialNumber = getLatestSequentialNumber();

  // Increment the sequential number.
  currentSequentialNumber++;
  // Pad with leading zeros to 3 digits (e.g., 1 becomes '001')
  const sequentialNumber = currentSequentialNumber.toString().padStart(3, '0');

  return `STVT/${yearCode}/${sequentialNumber}`;
}

// Save photo to Google Drive
function savePhotoToDrive(base64Data, fileName, mimeType, userEmail) {
  // Create a folder for user photos if it doesn't exist
  const folderName = 'CourseRegistrationPhotos';
  let folder;
  const folderIterator = DriveApp.getFoldersByName(folderName);

  if (folderIterator.hasNext()) {
    folder = folderIterator.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  // Create blob from base64 data
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);

  // Save file with user email in name to avoid duplicates
  const file = folder.createFile(blob);
  file.setName(userEmail + '_' + new Date().getTime() + '_' + fileName);

  // Return the URL to the file
  return file.getUrl();
}

// Send payment request email with PDF attachment
// function sendPaymentRequestEmail(email, name, uniqueId) {
//   //////////////////
//   const pdfBlob = createPaymentRequestPdfFromTemplate(name, uniqueId);

//   const subject = 'Payment Request for Course Registration';
//   const htmlBody = `
//     <h2>Hello ${name},</h2>
//     <p>Thank you for submitting your information for course registration.</p>
//     <p>Your unique ID is: <strong>${uniqueId}</strong></p>
//     <p>Please make a payment of ₹250 to complete your registration.</p>
//     <p>Payment details are attached in the PDF.</p>
//     <p>After making the payment, please log back in to verify your payment and select your course.</p>
//     <p>Regards,<br>Northern Railways </p>
//   `;

//   MailApp.sendEmail({
//     to: email,
//     subject: subject,
//     htmlBody: htmlBody,
//     attachments: [pdfBlob]
//   });
// }

// Create payment request PDF
function createPaymentRequestPdfFromTemplate(name, uniqueId) {
  const TEMPLATE_ID = '1I2Sm9igfMke5Gfj_qDy7nCwFxIYpNYdEAprcA-hEedI'; // Replace with your template's ID

  // 1. Make a copy of the template
  const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy('Payment Request for ' + name);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // 2. Replace placeholders
  body.replaceText('{{NAME}}', name);
  body.replaceText('{{UNIQUE_ID}}', uniqueId);
  body.replaceText('{{DATE}}', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
  // Add more replacements as needed

  doc.saveAndClose();

  // 3. Export as PDF
  const pdf = copy.getAs('application/pdf').setName('Payment_Request_' + uniqueId + '.pdf');

  // 4. (Optional) Delete the copy to keep Drive clean
  DriveApp.getFileById(copy.getId()).setTrashed(true);

  return pdf;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////
function verifyPaymentByAdmin(uniqueId, verified) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const paymentSheet = ss.getSheetByName(PAYMENT_SHEET);
    const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);

    if (!paymentSheet || !userInfoSheet) {
      return { success: false, message: 'Required sheets not found' };
    }

    // Find the payment records for the unique ID
    const paymentData = paymentSheet.getDataRange().getValues();
    let paymentRows = [];
    let userEmail = '';
    for (let i = 1; i < paymentData.length; i++) {
      if (paymentData[i][0] === uniqueId) {
        paymentRows.push(i);
        userEmail = paymentData[i][1];
      }
    }

    if (paymentRows.length === 0) {
      return { success: false, message: 'No payment records found for the given Unique ID' };
    }

    // Update the payment status to "true" for all relevant records
    paymentRows.forEach(row => {
      paymentSheet.getRange(row + 1, 4).setValue(verified);
    });

    // Get user info for confirmation email
    const userInfoData = userInfoSheet.getDataRange().getValues();
    let userRow = -1;
    let name = '', project = '', photoUrl = '', trainingDuration = '',
        guardianName = '', currentYear = '', college = '', course = '',
        address = '', branch = '', batchCode = '', projectCode = '', startDateOfBatch = '', phoneNumber = ''; // Added new variables

    for (let i = 1; i < userInfoData.length; i++) {
      if (userInfoData[i][0] === userEmail) {
        userRow = i;
        name = userInfoData[i][1];
        phoneNumber = userInfoData[i][2]; // Get phoneNumber from column C (index 2)
        address = userInfoData[i][4];
        college = userInfoData[i][3];
        course = userInfoData[i][9];
        uniqueId = userInfoData[i][6];
        currentYear = userInfoData[i][14];
        branch = userInfoData[i][10];
        project = userInfoData[i][13];
        photoUrl = userInfoData[i][5];
        guardianName = userInfoData[i][11];
        trainingDuration = userInfoData[i][12];
        batchCode = userInfoData[i][15];       // Get BatchCode
        projectCode = userInfoData[i][16];     // Get ProjectCode
        startDateOfBatch = userInfoData[i][17]; // Get StartDateOfBatch
        break;
      }
    }

    if (userRow === -1) {
      return { success: false, message: 'User information not found' };
    }

    // Create confirmation PDF with all necessary data
    const pdfBlob = createConfirmationPdf(
      name, uniqueId, guardianName, currentYear, college, course,
      address, photoUrl, branch, project, trainingDuration,
      batchCode, projectCode, startDateOfBatch, phoneNumber // Pass phoneNumber
    );



    return { success: true, message: 'Payment verified and confirmation email sent' };
  } catch (error) {
    Logger.log('Error in verifyPaymentByAdmin: ' + error);
    return { success: false, message: 'Error verifying payment: ' + error.message };
  }
}


// function sendConfirmationEmail(userEmail, name, uniqueId, project, photoUrl) {
//   const subject = 'Course Payment Confirmation';
//   const body = `Dear ${name},

//   Your payment for the course "${project}" has been received and is pending verification by the admin.

//   Unique ID: ${uniqueId}
//   Course: ${project}

//   Thank you for your patience.

//   Best regards,
//   Course Team`;

//   MailApp.sendEmail(userEmail, subject, body);
// }



// Send confirmation email with PDF attachment
// function sendConfirmationfinalEmail(email, name, uniqueId, project, photoUrl, trainingDuration, batchCode, projectCode, startDateOfBatch) {
//   // Create PDF with confirmation details
//   const pdfBlob = createConfirmationPdf(name, uniqueId, project, photoUrl, trainingDuration, batchCode, projectCode, startDateOfBatch);

//   const subject = 'Training Confirmation';
//   const htmlBody = `
//     <h2>Hello ${name},</h2>
//     <p>Congratulations! Your payment has been verified and your course registration is complete.</p>
//     <p>Your unique ID is: <strong>${uniqueId}</strong></p>
//     <p>You have registered for the <strong>${project}</strong> course.</p>
//     <p>Please find attached your course registration confirmation certificate.</p>
//     <p>Classes will begin on ${startDateOfBatch instanceof Date ? startDateOfBatch.toDateString() : startDateOfBatch || getNextMonday().toDateString()}. You will receive further instructions via email.</p>
//     <p>Regards,<br>Northern Railways</p>
//   `;

//   MailApp.sendEmail({
//     to: email,
//     subject: subject,
//     htmlBody: htmlBody,
//     attachments: [pdfBlob]
//   });
// }


// Get next Monday's date
function getNextMonday() {
  const today = new Date();
  const day = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
  const daysUntilNextMonday = day === 0 ? 1 : 8 - day;

  const nextMonday = new Date(today);
  nextMonday.setDate(today.getDate() + daysUntilNextMonday);

  return nextMonday;
}

// Create confirmation PDF with user photo
function createConfirmationPdf(name, uniqueId, guardianName, currentYear, college, course, address, photoUrl, branch, project, trainingDuration, batchCode, projectCode, startDateOfBatch, phoneNumber) {
  const TEMPLATE_ID = '1BeePP3z-getWJD3CP-UHEnjY0c_R4C5InQMC9o9991k'; // Replace with your template's Doc ID

  // 1. Make a copy of the template
  const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy('Certificate for ' + name);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // 2. Replace placeholders
  body.replaceText('{{Name}}', name);
  body.replaceText('{{College}}', college);
  body.replaceText('{{Branch}}', branch);
  body.replaceText('{{Address}}', address);
  body.replaceText('{{Father}}', guardianName);
  body.replaceText('{{Unique ID}}', uniqueId);
  body.replaceText('{{Course}}', course);
  body.replaceText('{{Year}}', currentYear);
  body.replaceText('{{Project}}', project);
  // Ensure startDateOfBatch is formatted correctly if it's a Date object
  body.replaceText('{{START_DATE}}', startDateOfBatch instanceof Date ? Utilities.formatDate(startDateOfBatch, Session.getScriptTimeZone(), "dd/MM/yyyy") : startDateOfBatch || Utilities.formatDate(getNextMonday(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
  body.replaceText('{{DURATION}}', trainingDuration);
  body.replaceText('{{BATCH_CODE}}', batchCode || 'N/A'); // New placeholder
  body.replaceText('{{PROJECT_CODE}}', projectCode || 'N/A'); // New placeholder
  body.replaceText('{{PHONE_NUMBER}}', phoneNumber || 'N/A'); // NEW: Add phone number placeholder

  // 3. (Optional) Insert photo if photoUrl is provided
  if (photoUrl) {
    try {
      const fileIdMatch = photoUrl.match(/[-\w]{25,}/);
      if (fileIdMatch && fileIdMatch[0]) {
        const photoBlob = DriveApp.getFileById(fileIdMatch[0]).getBlob();
        // Insert at a specific placeholder, e.g., {{PHOTO}}
        const found = body.findText('{{PHOTO}}');
        if (found) {
          const element = found.getElement();
          element.asText().setText(''); // Remove the placeholder
          
          body.insertImage(body.getChildIndex(element.getParent()), photoBlob);
        
        }
      }
    } catch (e) {
      Logger.log('Photo insert error: ' + e);
    }
  }

  doc.saveAndClose();

  // 4. Export as PDF
  const pdf = copy.getAs('application/pdf').setName('Course_Registration_' + uniqueId + '.pdf');

  // 5. (Optional) Delete the copy to keep Drive clean
  copy.setTrashed(true);

  return pdf;
}


// Function to verify payment and course selection
function verifyPaymentAndCourse(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const paymentSheet = ss.getSheetByName(PAYMENT_SHEET);
  const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);

  if (!paymentSheet || !userInfoSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const uniqueId = formData.uniqueId;
  const transactionId = formData.transactionId;

  // Check if the unique ID exists in the UserInfo sheet
  const userInfoData = userInfoSheet.getDataRange().getValues();
  let userEmail = '';
  for (let i = 1; i < userInfoData.length; i++) {
    if (userInfoData[i][6] === uniqueId) {
      userEmail = userInfoData[i][0];
      break;
    }
  }

  if (!userEmail) {
    return { success: false, message: 'Invalid Unique ID' };
  }
  const paymentData = paymentSheet.getDataRange().getValues();
    let paymentVerified = false;

    for (let i = 1; i < paymentData.length; i++) {
      console.log(`Checking payment row ${i}: ${paymentData[i][0]} (uniqueId), ${paymentData[i][4]} (verified)`);
      if (paymentData[i][0] === formData.uniqueId) {
        if (paymentData[i][3] === true) {
          paymentVerified = true;
          break;
        } else {
          return { success: false, message: 'Payment not verified by admin yet' };
        }
      }
    }



  // Add payment information to the Payment sheet
  paymentSheet.appendRow([uniqueId, userEmail, transactionId, 'No', new Date()]);

  if (paymentVerified) {
     console.log("Payment verified, sending redirect to confirmation page");
     return {
       success: true,
       message: 'Your payment has been verified! Redirecting to confirmation page...',
       redirect: 'https://script.google.com/macros/s/AKfycbxrUDVoI4_yFwPxPaKdHxqln4r0JIlc77py9v-RhJjgO8nQ-6GYeiwIvu94FvR_sCY/exec?page=confirmation'
      };
    }

  // Send a response back to the client
  return { success: true, message: 'Payment information submitted successfully. Please wait for verification.', redirect: getScriptUrl() + '?page=confirmation' };
}

// Function to get the script URL
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Function to log out the user
function logout() {
  PropertiesService.getUserProperties().deleteProperty('userEmail');
  CacheService.getUserCache().remove('userEmail');
  return { success: true };
}


// Generate Google Calendar link
function generateCalendarLink(u, courseData) {
  const startDate = new Date(courseData.startDate);
  const endDate = new Date(courseData.endDate);

  // Format dates for Google Calendar
  const startFormatted = startDate.toISOString().replace(/-|:|\.\d\d\d/g, "");
  const endFormatted = endDate.toISOString().replace(/-|:|\.\d\d\d/g, "");

  const eventTitle = encodeURIComponent("Northern Railways Training: " + u.project);
  const eventDetails = encodeURIComponent("Registration ID: " + u.uniqueId + "\nLocation: " + courseData.location);

  return "https://www.google.com/calendar/render?action=TEMPLATE&text=" + eventTitle + "&dates=" + startFormatted + "/" + endFormatted + "&details=" + eventDetails + "&location=" + encodeURIComponent(courseData.location);
}
// Returns user info and status for dashboard
// In code.gs.js

/**
 * CORRECTED: Returns user info and status for the dashboard.
 * Fixes the "Cannot read properties of undefined" error by correctly fetching
 * and iterating over the data from the Users, UserInfo, and Payments sheets.
 */
function getCurrentUserDashboardInfo() {
  const email = getCurrentUserEmail();
  if (!email) {
    // Return a default object if the user is not found, to prevent client-side errors.
    return { email: null, fullName: '', uniqueId: '', userInfoSubmitted: false, paymentVerified: false };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(USERS_SHEET);
    const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);
    const paymentSheet = ss.getSheetByName(PAYMENT_SHEET);

    // Initialize the user object with the current user's email.
    let user = {
      email: email,
      fullName: '',
      uniqueId: '',
      userInfoSubmitted: false,
      paymentVerified: false
    };

    // 1. Get Full Name from the 'Users' sheet.
    if (usersSheet) {
      const usersData = usersSheet.getDataRange().getValues();
      for (let i = 1; i < usersData.length; i++) {
        if (usersData[i] && usersData[i][0] === email) {
          user.fullName = usersData[i][1]; // Get name from column B
          break;
        }
      }
    }

    // 2. Get Unique ID and check if user info has been submitted from the 'UserInfo' sheet.
    if (userInfoSheet) {
      const userInfoData = userInfoSheet.getDataRange().getValues();
      for (let i = 1; i < userInfoData.length; i++) {
        if (userInfoData[i] && userInfoData[i][0] === email) {
          user.userInfoSubmitted = true;
          user.uniqueId = userInfoData[i][6]; // Get Unique ID from column G
          break;
        }
      }
    }

    // 3. If user info is submitted, check for payment verification.
    if (user.userInfoSubmitted && paymentSheet) {
      const paymentData = paymentSheet.getDataRange().getValues();
      for (let i = 1; i < paymentData.length; i++) {
        // Check if the row exists, the email matches, and the verification status is boolean `true`.
        if (paymentData[i] && paymentData[i][1] === email && paymentData[i][3] === true) {
          user.paymentVerified = true;
          break;
        }
      }
    }

    return user;

  } catch (e) {
    Logger.log("Error in getCurrentUserDashboardInfo: " + e.message);
    // Return a safe object in case of any unexpected errors
    return { email: email, fullName: 'Error loading data', uniqueId: 'Error', userInfoSubmitted: false, paymentVerified: false };
  }
}

// --- NEW ADMIN FUNCTIONS ---

/**
 * Fetches unique batch codes and project names from the UserInfo sheet
 * for populating admin dropdowns.
 * @returns {Object} An object containing unique batch codes and all batch-project combinations.
 */
function getBatchAndProjectOptionsForAdmin() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);

    if (!userInfoSheet) {
      return { success: false, message: 'UserInfo sheet not found.', combinations: [], uniqueBatchCodes: [] };
    }

    const data = userInfoSheet.getDataRange().getValues();
    if (data.length <= 1) { // Only headers or no data
      return { success: true, message: 'No student data found.', combinations: [], uniqueBatchCodes: [] };
    }

    const combinations = [];
    const uniqueBatchCodes = new Set();

    // Iterate from the second row (index 1) to skip headers
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Ensure row has enough columns for BatchCode (index 15) and Project (index 13)
      if (row.length > 15 && row[13] && row[15]) {
        const batchCode = row[15];
        const projectName = row[13];
        combinations.push({ batchCode: batchCode, projectName: projectName });
        uniqueBatchCodes.add(batchCode);
      }
    }

    return {
      success: true,
      combinations: combinations,
      uniqueBatchCodes: Array.from(uniqueBatchCodes).sort() // Convert Set to Array and sort
    };

  } catch (error) {
    Logger.log('Error in getBatchAndProjectOptionsForAdmin: ' + error.message);
    return { success: false, message: 'Error fetching batch/project data: ' + error.message, combinations: [], uniqueBatchCodes: [] };
  }
}

/**
 * Generates a PDF report for a specific batch and project using the Google Docs template.
 * This function is called by the admin page.
 * @param {string} batchCode The selected batch code.
 * @param {string} projectName The selected project name.
 * @returns {Object} An object with success status and the URL of the generated PDF.
 */
function generateAndGetBatchProjectPdfUrl(batchCode, projectName) {
  // Now, generateBatchProjectDocReport no longer uses a templateDocId parameter
  // It creates the document from scratch.
  const result = generateBatchProjectDocReport(batchCode, projectName);

  if (result.success) {
    return { success: true, url: result.url };
  } else {
    return { success: false, message: result.message };
  }
}


// Returns a temporary Drive URL to the payment request PDF
function getPaymentRequestPdfUrl() {
  const email = getCurrentUserEmail();
  if (!email) return { url: null };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);
  const data = userInfoSheet.getDataRange().getValues();
  let name = "", uniqueId = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      name = data[i][1];
      uniqueId = data[i][6];
      break;
    }
  }
  if (!name || !uniqueId) return { url: null };
  const pdfBlob = createPaymentRequestPdfFromTemplate(name, uniqueId);
  const tempFolder = DriveApp.getRootFolder();
  const file = tempFolder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: file.getUrl() };
}

// Returns a temporary Drive URL to the confirmation PDF
function getConfirmationPdfUrl() {
  const email = getCurrentUserEmail();
  if (!email) {
    Logger.log("getConfirmationPdfUrl failed: User not logged in.");
    return { url: null };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);
  if (!userInfoSheet) {
    Logger.log("getConfirmationPdfUrl failed: UserInfo sheet not found.");
    return { url: null };
  }

  const data = userInfoSheet.getDataRange().getValues();

  // 1. Declare ALL variables needed for the PDF
  let name = "", uniqueId = "", guardianName = "", currentYear = "", college = "",
      course = "", address = "", photoUrl = "", branch = "", project = "", trainingDuration = "",
      batchCode = "", projectCode = "", startDateOfBatch = "", phoneNumber = ""; // Added new variable here

  // 2. Find the user and populate ALL variables from the correct columns
  for (let i = 1; i < data.length; i++) {
    if (data[i] && data[i][0] === email) {
      name = data[i][1];             // FullName
      phoneNumber = data[i][2];      // Phone Number (Column C, index 2)
      address = data[i][4];          // Address
      college = data[i][3];          // College
      course = data[i][9];           // Course
      uniqueId = data[i][6];         // UniqueId
      currentYear = data[i][14];     // CurrentYear
      branch = data[i][10];          // Branch
      project = data[i][13];         // Project
      photoUrl = data[i][5];         // PhotoUrl
      guardianName = data[i][11];    // GuardianName
      trainingDuration = data[i][12];// TrainingDuration
      batchCode = data[i][15];       // BatchCode (assuming column P, index 15)
      projectCode = data[i][16];     // ProjectCode (assuming column Q, index 16)
      startDateOfBatch = data[i][17]; // StartDateOfBatch (assuming column R, index 17)
      break;
    }
  }

  if (!name || !uniqueId) {
    Logger.log(`getConfirmationPdfUrl failed: Could not find user data for ${email}`);
    return { url: null };
  }

  // 3. Call the PDF creation function with all the data
  // AND capture the returned PDF blob in a variable.
  const pdfBlob = createConfirmationPdf(
    name, uniqueId, guardianName, currentYear, college, course,
    address, photoUrl, branch, project, trainingDuration,
    batchCode, projectCode, startDateOfBatch, phoneNumber // Pass phoneNumber
  );

  // 4. Create a temporary file and return its URL
  const tempFolder = DriveApp.getRootFolder();
  const file = tempFolder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = file.getUrl();

  // Clean up the temporary file after a delay to allow for download
  Utilities.sleep(100); // 1 seconds

  return { url: url };
}

/**
 * Helper function to calculate end date based on start date and duration string.
 * @param {Date} startDate The start date of the training.
 * @param {string} durationString The duration string (e.g., "4 Weeks", "6 Weeks", "8 Weeks").
 * @returns {Date|null} The calculated end date, or null if duration is invalid.
 */
function calculateEndDate(startDate, durationString) {
  if (!(startDate instanceof Date) || isNaN(startDate.getTime())) {
    Logger.log(`calculateEndDate: Invalid startDate: ${startDate}`);
    return null; // Invalid start date
  }

  const weeksMatch = durationString.match(/(\d+)\s*Weeks/i);
  if (!weeksMatch || !weeksMatch[1]) {
    Logger.log(`calculateEndDate: Could not parse weeks from duration string: "${durationString}"`);
    return null; // Could not parse weeks from duration string
  }

  const numberOfWeeks = parseInt(weeksMatch[1], 10);
  if (isNaN(numberOfWeeks)) {
    Logger.log(`calculateEndDate: Invalid number of weeks parsed: ${weeksMatch[1]}`);
    return null; // Invalid number of weeks
  }

  const endDate = new Date(startDate);
  // Add weeks, then subtract one day to make it inclusive (e.g., 4 weeks from Mon to Fri is 27 days)
  endDate.setDate(startDate.getDate() + (numberOfWeeks * 7) - 1);
  return endDate;
}

/**
 * Helper function to extract the number of weeks from a duration string.
 * @param {string} durationString The duration string (e.g., "4 Weeks", "6 Weeks").
 * @returns {string} The number of weeks as a string, or 'N/A' if not found.
 */
function getWeeksFromDurationString(durationString) {
  const weeksMatch = durationString.match(/(\d+)\s*Weeks/i);
  return weeksMatch && weeksMatch[1] ? weeksMatch[1] : 'N/A';
}


/**
 * Generates a Google Docs report of students belonging to a specific batch and project.
 * This function now creates the document structure entirely in code, without a template document.
 *
 * @param {string} targetBatchCode The batch code to filter students by (e.g., "STVT/20250714").
 * @param {string} targetProjectName The full project name to filter students by (e.g., "Mobile App for Live Train Tracking (6w)").
 * @returns {Object} An object indicating success and the URL of the generated report document, or an error message.
 */
function generateBatchProjectDocReport(targetBatchCode, targetProjectName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userInfoSheet = ss.getSheetByName(USER_INFO_SHEET);

    if (!userInfoSheet) {
      Logger.log('ERROR: UserInfo sheet not found in the main spreadsheet.');
      return { success: false, message: 'UserInfo sheet not found in the main spreadsheet.' };
    }

    const userInfoData = userInfoSheet.getDataRange().getValues();
    const studentsInBatchProject = [];
    for (let i = 1; i < userInfoData.length; i++) {
      const row = userInfoData[i];
      if (row.length > 17) { // Ensure row has enough columns
        const studentBatchCode = row[15];
        const studentProject = row[13];

        if (studentBatchCode === targetBatchCode && studentProject === targetProjectName) {
          studentsInBatchProject.push({
            email: row[0], fullName: row[1], phoneNumber: row[2], college: row[3], address: row[4],
            photoUrl: row[5], uniqueId: row[6], dateOfBirth: row[7], gender: row[8], course: row[9],
            branch: row[10], guardianName: row[11], trainingDuration: row[12], project: row[13],
            currentYear: row[14], batchCode: row[15], projectCode: row[16], startDateOfBatch: row[17]
          });
        }
      } else {
        Logger.log(`WARNING: Row ${i+1} in UserInfo sheet has insufficient columns. Skipping.`);
      }
    }

    if (studentsInBatchProject.length === 0) {
      Logger.log(`INFO: No students found for Batch Code: "${targetBatchCode}" and Project: "${targetProjectName}".`);
      return { success: false, message: `No students found for Batch Code: "${targetBatchCode}" and Project: "${targetProjectName}".` };
    }

    // Get common data for summary (from the first student, assuming consistent for the batch/project)
    const firstStudent = studentsInBatchProject[0];
    const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy"); // Format as dd.MM.yyyy

    // Ensure startDateOfBatch is a Date object for calculation
    const studentStartDateForCalc = firstStudent.startDateOfBatch instanceof Date ?
    firstStudent.startDateOfBatch : new Date(firstStudent.startDateOfBatch);



    const batchStartDateSummary = Utilities.formatDate(studentStartDateForCalc, Session.getScriptTimeZone(), "dd.MM.yy");

    const trainingWeeks = getWeeksFromDurationString(firstStudent.trainingDuration || '');

    // Calculate the end date for the batch summary
    const batchEndDate = calculateEndDate(studentStartDateForCalc, firstStudent.trainingDuration || '');
    const batchEndDateSummary = batchEndDate instanceof Date ? Utilities.formatDate(batchEndDate, Session.getScriptTimeZone(), "dd.MM.yy") : 'N/A';


    const docTitle = `Batch_Project_Report_${targetBatchCode.replace(/\//g, '_')}_${getProjectShortName(targetProjectName)}`;
    // Create a new blank Google Doc
    const doc = DocumentApp.create(docTitle);
    const body = doc.getBody();
    Logger.log(`New document created with title: ${docTitle}`);

    // --- Build the document content dynamically ---

    // Header section
    body.appendParagraph(`${targetBatchCode}/${firstStudent.projectCode || 'N/A'}`)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph(`Date:${reportDate}`)
        .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendParagraph('Automation/STC')
        .setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    body.appendParagraph('N. Rly, Loco Workshop, Charbagh, Lucknow.')
        .setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    body.appendParagraph(''); // Blank line

    body.appendParagraph(`Sub: - ${trainingWeeks} Week Training Projects Vocational Training to BE/Diploma Students.`)
        .setBold(true)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendParagraph(''); // Blank line

    // Main body text
    const mainParagraph = body.appendParagraph('In Subject to above the undernoted trainee is directed to yours end for vocational training from ');
    mainParagraph.appendText(`${batchStartDateSummary || 'N/A'}`).setBold(true); // Dynamic Start Date
    mainParagraph.appendText(` to ${batchEndDateSummary || 'N/A'}. The project allotted to the trainee is mentioned as under.`); // Dynamic End Date

    body.appendParagraph(`Project title & training area: ${firstStudent.project || 'N/A'}`)
        .setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    body.appendParagraph(''); // Blank line

    // Student Details Table
    const tableHeaders = [
      'S. No', 'Unique ID', 'Name', 'College', 'Branch', 'Year', 'Course'
    ];
    const table = body.appendTable();

    // Append header row
    const headerRow = table.appendTableRow();
    tableHeaders.forEach(headerText => {
      headerRow.appendTableCell(headerText).setBold(true);
    });

    // Append student data rows
    studentsInBatchProject.forEach((student, index) => {
      const row = table.appendTableRow();

      row.appendTableCell((index + 1).toString());
      row.appendTableCell(student.uniqueId || 'N/A');
      row.appendTableCell(student.fullName || 'N/A');
      row.appendTableCell(student.college || 'N/A');
      row.appendTableCell(student.branch || 'N/A');
      row.appendTableCell(student.currentYear ? student.currentYear.toString() : 'N/A');
      row.appendTableCell(student.course || 'N/A');
    });

    // Add extra blank rows to match the visual length of your template if needed
    const numBlankRows = 15; // Approximate number of blank rows seen in the image
    for (let i = 0; i < numBlankRows; i++) {
        table.appendTableRow(); // Appends an empty row
    }


    body.appendParagraph('Kindly release him and send his attendance along with the project report. Necessary assistance may be provided if any.');
    body.appendParagraph(''); // Blank line

    // Conditions section
    body.appendParagraph('The vocational training is allowed with following conditions:-')
        .setBold(false)
        .setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    const conditions = [
      'The trainee will have to follow all rules and regulations & will maintain high discipline all the time during training.',
      'The trainee will not have access to any confidential document.',
      'Camera Mobiles are strictly prohibited in the training area. If found, Camera Mobiles will be seized and will be returned after producing Original Bill of Mobile on completion of training.',
      'The trainee will have to make his/her own arrangement for boarding lodging and transportation',
      'No remuneration in the shape of stipend, salary or allowances of any kind will be paid to trainees by the Railway Administration. No passes or PTO will be issued to them.',
      'The trainees will indemnify Railway Administration concerned for any loss or damage to equipment and fittings that may be caused by the trainee during the course of training in workshop area.',
      'The trainee shall not be the employees of the Railway Administration and as such will not be entitled to any compensation or damages from the railway Administration under any circumstances and the trainee will not make any claim thereof.',
      'The trainee shall be under the control and discipline of Dy. CME concerned',
      'This vocational training is purely on the risk and cost of the trainee and can be disallowed at any time without assigning any reason/notice.'
    ];

    conditions.forEach((condition, index) => {
      body.appendParagraph(`${index + 1}. ${condition}`);
    });

    body.appendParagraph(''); // Blank line
    body.appendParagraph('Course Coordinator,')
        .setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    doc.saveAndClose();
    Logger.log('Document content built, saved, and closed.');

    // Get the newly created file from Drive to set sharing and get URL
    const file = DriveApp.getFileById(doc.getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('Sharing permissions set.');

    return { success: true, url: file.getUrl(), docName: docTitle, studentCount: studentsInBatchProject.length };

  } catch (error) {
    Logger.log('FATAL ERROR in generateBatchProjectDocReport: ' + error.message);
    return { success: false, message: 'Error generating Docs report: ' + error.message };
  }
}

function isAdminUser(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(USERS_SHEET);

    if (!usersSheet) {
      console.error('Users sheet not found');
      return false;
    }

    const userData = usersSheet.getDataRange().getValues();
    // Assuming 'IsAdmin' is the 5th column (index 4)
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] === email) {
        return userData[i][5] === true; // Return true if IsAdmin is true
      }
    }
    return false; // User not found or not marked as admin
  } catch (error) {
    console.error('Error in isAdminUser:', error);
    return false;
  }
}

// Helper function (already exists in your code, but added here for completeness if copied standalone)
function getProjectShortName(projectName) {
  let shortName = projectName.replace(/\s*\(\d+w\)\s*$/, '');
  shortName = shortName.split(' ').slice(0, 3).join(' ');
  shortName = shortName.replace(/[^a-zA-Z0-9]/g, '').substring(0, 20).toUpperCase();
  return shortName;
}