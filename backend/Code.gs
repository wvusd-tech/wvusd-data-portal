/**
 * ==============================================================================
 * WALNUT VALLEY UNIFIED SCHOOL DISTRICT (WVUSD)
 * Illuminate Student Assessment & Multi-Year Longitudinal Portal
 * Google Apps Script Web App Backend (Code.gs)
 * ==============================================================================
 *
 * SPREADSHEET STRUCTURE:
 * 1. "Roster" (or "Roster Tab"):
 *    - Strict Column A: Student ID
 *    - Dynamic Column matching: Student Name, Grade, School/Site, Teacher Name & Email,
 *      Demographics (Homeless/Foster, SPED, 504, EL Status)
 *    - Site Admin routing: Cols O:P (or X:Y)
 *    - Super Admin list: Cell Q1 (or Col Z)
 * 2. "Users" or "User" Tab (Optional):
 *    - Name, Email, Job Title, Site
 * 3. Assessment Tabs (e.g. "3rd ELA", "3rd Math", "4th ELA", "4th Math", "5th ELA", "5th Math"):
 *    - Strict Column A: Student ID
 *    - Dynamic Column matching for assessment metrics & checkpoints
 *
 * SCRIPT PROPERTIES:
 * - APP_SECRET_TOKEN: Shared secret token verified via verifyAppToken_
 * ==============================================================================
 */

const ROSTER_TAB_NAMES = ['Roster', 'Roster Tab'];
const ALLOWED_DOMAINS = ['wvusd.org', 'sub.wvusd.org'];

// Demographic headers to skip when building raw assessment maps
const SKIP_ASSESSMENT_MAP_HEADERS = new Set([
  'student id', 'id', 'local id', 'ssid', 'state student id',
  'last name', 'first name', 'middle name', 'full name', 'student name', 'name',
  'birth date', 'gender', 'sex', 'ethnicity', 'race',
  'enrollment site name', 'school', 'school name', 'school code'
]);

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const params = (e && e.parameter) ? Object.assign({}, e.parameter) : {};
    if (e && e.postData && e.postData.contents) {
      try {
        const postData = JSON.parse(e.postData.contents);
        for (const k in postData) {
          params[k] = postData[k];
        }
      } catch (parseErr) {
        // Non-JSON POST body, fallback to e.parameter
      }
    }

    // 1. App Token Verification (Read-side zero-trust check)
    if (!verifyAppToken_(params, e)) {
      return jsonResponse({
        status: "Forbidden",
        success: false,
        message: "Invalid app token.",
        error: "Invalid app token."
      }, 403);
    }

    // 2. Resolve user email from query / payload
    let userEmail = (params.email || params.userEmail || '').toString().trim().toLowerCase();
    if (!userEmail && params.id_token) {
      try {
        userEmail = verifyIdToken_(params.id_token);
      } catch (idErr) {
        return jsonResponse({
          status: "Forbidden",
          success: false,
          message: "Expired or invalid Google token.",
          error: "Expired or invalid Google token."
        }, 403);
      }
    }
    if (!userEmail && params.test_email) {
      userEmail = (params.test_email || '').toString().trim().toLowerCase();
    }
    if (!userEmail) {
      try {
        userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
      } catch (sessErr) {
        // Ignore
      }
    }

    // Security Check: Enforce non-empty email and @wvusd.org domain
    if (!isAllowedWvusdEmail_(userEmail)) {
      return jsonResponse({
        status: "Forbidden",
        success: false,
        message: "Unauthorized: Access restricted to @wvusd.org and @sub.wvusd.org staff accounts.",
        error: "Unauthorized user."
      }, 403);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return jsonResponse({
        status: "Error",
        success: false,
        message: "Spreadsheet context could not be resolved.",
        error: "Spreadsheet context could not be resolved."
      }, 500);
    }

    // Locate Roster tab
    let rosterSheet = null;
    for (let i = 0; i < ROSTER_TAB_NAMES.length; i++) {
      rosterSheet = ss.getSheetByName(ROSTER_TAB_NAMES[i]);
      if (rosterSheet) break;
    }
    const usersSheet = ss.getSheetByName('Users') || ss.getSheetByName('User');

    if (!rosterSheet) {
      return jsonResponse({
        status: "Error",
        success: false,
        message: "Missing required 'Roster' tab.",
        error: "Missing required 'Roster' tab."
      }, 500);
    }

    const rosterRows = rosterSheet.getDataRange().getValues();
    if (!rosterRows || rosterRows.length < 2) {
      return jsonResponse({
        status: "Error",
        success: false,
        message: "Roster tab contains no data rows.",
        error: "Roster tab contains no data rows."
      }, 500);
    }

    const rbacContext = {
      superAdmins: new Set(['kwendland@wvusd.org']),
      siteAdminMap: new Map()
    };

    extractStandardRosterRbac_(rosterRows, rbacContext);

    // 3. Staff Roster & Site Scope Verification BEFORE touching assessment tabs
    const studentsMap = new Map();
    ingestDynamicSheetTab(rosterRows, rosterSheet.getName(), studentsMap, rbacContext, true, ss);

    const allRosterStudents = Array.from(studentsMap.values());
    const userAccess = resolveUserAccess_(userEmail, allRosterStudents, rbacContext, usersSheet);

    if (!userAccess.authorized) {
      return jsonResponse({
        status: "Forbidden",
        success: false,
        message: "Unauthorized: Account not found in Roster or Users directory. Contact Tech TOSA for access.",
        error: "Unauthorized user."
      }, 403);
    }

    // 4. Server-Side Scope Filtering: Retain only students matching the user's authorized site(s) / classroom
    for (const [id, student] of studentsMap.entries()) {
      if (!canViewStudent_(userAccess, student)) {
        studentsMap.delete(id);
      }
    }

    // 5. Query and Ingest Illuminate Assessment Tabs strictly for authorized students
    const allSheets = ss.getSheets();
    const availableTabs = [];

    for (let s = 0; s < allSheets.length; s++) {
      const sheetObj = allSheets[s];
      const tName = sheetObj.getName();
      const lowerT = tName.toLowerCase().trim();

      if (
        lowerT === 'roster' ||
        lowerT === 'roster tab' ||
        lowerT === 'users' ||
        lowerT === 'user' ||
        lowerT === 'sites' ||
        lowerT === 'site' ||
        lowerT === 'config'
      ) {
        continue;
      }

      const tabRows = sheetObj.getDataRange().getValues();
      if (tabRows && tabRows.length > 1) {
        availableTabs.push(tName);
        ingestDynamicSheetTab(tabRows, tName, studentsMap, rbacContext, false, ss);
      }
    }

    const scopedStudents = Array.from(studentsMap.values());

    return jsonResponse({
      status: "Success",
      success: true,
      user: userAccess,
      meta: {
        totalStudents: scopedStudents.length,
        assessmentTabs: availableTabs,
        lastUpdated: new Date().toISOString()
      },
      students: scopedStudents
    });

  } catch (err) {
    return jsonResponse({
      status: "Error",
      success: false,
      error: "Server Error: " + (err && err.message ? err.message : String(err)),
      message: err && err.message ? err.message : String(err)
    }, 500);
  }
}

/**
 * App Token Verification:
 * Validates provided token against APP_SECRET_TOKEN in Script Properties.
 */
function verifyAppToken_(payload, e) {
  var expectedToken = PropertiesService.getScriptProperties().getProperty("APP_SECRET_TOKEN");
  if (!expectedToken) {
    // If not configured in Script Properties, allow for backwards-compatibility or dev
    return true;
  }
  var providedToken = String(
    (payload && payload.appToken) ||
    (e && e.parameter && e.parameter.appToken) ||
    ""
  ).trim();
  return providedToken === String(expectedToken).trim();
}

function isAllowedWvusdEmail_(email) {
  if (!email) return false;
  const clean = String(email).toLowerCase().trim();
  const parts = clean.split('@');
  if (parts.length !== 2) return false;
  if (parts[1] === 'stu.wvusd.org') return false;
  return ALLOWED_DOMAINS.indexOf(parts[1]) !== -1;
}

function verifyIdToken_(idToken) {
  const response = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (response.getResponseCode() !== 200) {
    throw new Error('Expired or invalid Google Sign-In token.');
  }
  const data = JSON.parse(response.getContentText());
  if (!data.email || data.email_verified !== 'true') {
    throw new Error('Could not verify Google account email.');
  }
  return data.email.toLowerCase().trim();
}

function parseEmailList_(rawCell) {
  if (!rawCell) return [];
  return String(rawCell)
    .split(/[,;\n]+/)
    .map(e => e.toLowerCase().trim())
    .filter(e => e.indexOf('@') !== -1);
}

function extractStandardRosterRbac_(rosterRows, rbacContext) {
  if (!rosterRows || rosterRows.length === 0) return;

  // SuperAdmins in Q1 (col index 16) or Z1 (col index 25)
  const q1 = rosterRows[0] && rosterRows[0][16] ? String(rosterRows[0][16]) : '';
  const z1 = rosterRows[0] && rosterRows[0][25] ? String(rosterRows[0][25]) : '';
  parseEmailList_(q1 + ',' + z1).forEach(e => rbacContext.superAdmins.add(e));

  for (let r = 0; r < rosterRows.length; r++) {
    const row = rosterRows[r];
    if (!row) continue;

    // Check O:P (cols 14 & 15) or X:Y (cols 23 & 24)
    let siteName = '';
    let adminEmails = '';

    if (row.length > 24 && row[23] && row[24]) {
      siteName = String(row[23]).trim();
      adminEmails = String(row[24]);
    } else if (row.length > 15 && row[14] && row[15]) {
      siteName = String(row[14]).trim();
      adminEmails = String(row[15]);
    }

    if (siteName && adminEmails) {
      const emailList = parseEmailList_(adminEmails);
      if (emailList.length > 0) {
        if (!rbacContext.siteAdminMap.has(siteName)) {
          rbacContext.siteAdminMap.set(siteName, new Set());
        }
        const setRef = rbacContext.siteAdminMap.get(siteName);
        emailList.forEach(e => setRef.add(e));
      }
    }
  }
}

function resolveUserAccess_(email, allStudents, rbacContext, usersSheet) {
  if (rbacContext.superAdmins.has(email) || email === 'kwendland@wvusd.org') {
    return {
      email: email,
      name: 'Super Admin',
      role: 'SuperAdmin',
      school: 'ALL',
      authorized: true
    };
  }

  const assignedSites = [];
  for (const [site, adminSet] of rbacContext.siteAdminMap.entries()) {
    if (adminSet.has(email)) {
      assignedSites.push(site);
    }
  }

  if (assignedSites.length > 0) {
    return {
      email: email,
      name: 'Site Administrator / Specialist',
      role: 'SiteAdmin',
      school: assignedSites.join(','),
      authorized: true
    };
  }

  const matchedRoster = allStudents.filter(s => {
    const staffEmails = parseEmailList_(s.staffEmail);
    return staffEmails.indexOf(email) !== -1;
  });

  if (matchedRoster.length > 0) {
    const sample = matchedRoster[0];
    const gradeNum = parseInt(sample.grade, 10);
    const isElem = isNaN(gradeNum) ? true : gradeNum <= 5;
    return {
      email: email,
      name: sample.staffName || sample.teacher || email.split('@')[0],
      role: isElem ? 'ElementaryTeacher' : 'Counselor',
      school: sample.school || sample.site || 'ALL',
      authorized: true
    };
  }

  if (usersSheet) {
    const uData = usersSheet.getDataRange().getValues();
    for (let k = 1; k < uData.length; k++) {
      const uName = String(uData[k][0] || '').trim();
      const uEmail = String(uData[k][1] || '').trim().toLowerCase();
      const uTitle = String(uData[k][2] || 'Teacher').trim();
      const uSite = String(uData[k][3] || '').trim();

      if (uEmail === email) {
        return {
          email: email,
          name: uName || email.split('@')[0],
          jobTitle: uTitle,
          role: 'SecondaryOrSiteTeacher',
          school: uSite || 'ALL',
          authorized: true
        };
      }
    }
  }

  return { email: email, authorized: false };
}

function canViewStudent_(userAccess, student) {
  if (userAccess.role === 'SuperAdmin' || userAccess.school === 'ALL') return true;

  if (userAccess.role === 'ElementaryTeacher') {
    const staffEmails = parseEmailList_(student.staffEmail || student.teacherEmail);
    return staffEmails.indexOf(userAccess.email) !== -1;
  }

  const allowedSites = String(userAccess.school || '')
    .toLowerCase()
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);

  const studentSchool = String(student.school || student.site || '').toLowerCase().trim();
  return allowedSites.some(
    site => studentSchool.indexOf(site) !== -1 || site.indexOf(studentSchool) !== -1
  );
}

function extractStrictColumnAStudentId(row) {
  if (!Array.isArray(row) || row.length === 0) return '';
  const raw = (row[0] !== null && row[0] !== undefined ? row[0] : '').toString().trim();
  if (!raw) return '';
  if (/^(?:student\s*id|id|local\s*id|stu\s*id|perm\s*id|ssid|state\s*id|student\s*#)$/i.test(raw)) {
    return '';
  }
  const stripped = raw.replace(/^#/, '').replace(/\.0+$/, '').replace(/^0+/, '').trim();
  return stripped || (raw === '0' ? '0' : '');
}

function isRawSchoolCodeOrPlaceholder_(val) {
  const s = (val !== null && val !== undefined ? val : '').toString().trim();
  if (!s) return true;
  if (/^\d{3,7}$/.test(s)) return true;
  const lower = s.toLowerCase();
  return (
    lower === 'wvusd' ||
    lower === 'district' ||
    lower === 'district office' ||
    lower === 'school' ||
    lower === 'all'
  );
}

function analyzeSheetHeaders(headerRow, tabName, isRosterFallback) {
  const isRosterTabContext = /^roster(?:\s*tab)?$/i.test((tabName || '').toString().trim());
  const mapping = [];

  for (let c = 0; c < headerRow.length; c++) {
    const rawHeader = (headerRow[c] !== null && headerRow[c] !== undefined ? headerRow[c] : '').toString().trim();
    const cleanLower = rawHeader.toLowerCase();

    let fieldType = 'ASSESSMENT_METRIC';

    if (c === 0) {
      fieldType = 'STUDENT_ID';
    } else if (isRosterTabContext && c === 14) {
      fieldType = 'SCHOOL_NAME';
    } else if (isRosterTabContext && c !== 14) {
      fieldType = 'OTHER';
    } else if (!isRosterTabContext) {
      if (cleanLower.includes('school') || cleanLower.includes('site')) {
        fieldType = 'SCHOOL_CODE';
      } else if (cleanLower.includes('teacher') || cleanLower.includes('counselor') || cleanLower.includes('user name')) {
        fieldType = 'TEACHER_NAME';
      } else if (cleanLower.includes('grade') && !cleanLower.includes('equivalent')) {
        fieldType = 'GRADE';
      } else if (cleanLower.includes('name') && (cleanLower.includes('student') || cleanLower.includes('first') || cleanLower.includes('last'))) {
        fieldType = 'STUDENT_NAME';
      } else {
        fieldType = 'ASSESSMENT_METRIC';
      }
    }

    mapping.push({
      colIndex: c,
      rawHeader: rawHeader,
      cleanHeader: cleanLower,
      fieldType: fieldType
    });
  }

  return mapping;
}

function ingestDynamicSheetTab(rows, tabName, studentsMap, rbacContext, isRosterTab, ss) {
  if (!rows || rows.length < 2) return;

  let headerRowIdx = 0;
  for (let r = 0; r < Math.min(5, rows.length); r++) {
    const firstCell = String(rows[r][0] || '').toLowerCase().trim();
    if (firstCell.includes('id') || firstCell.includes('student')) {
      headerRowIdx = r;
      break;
    }
  }

  const headerAnalysis = analyzeSheetHeaders(rows[headerRowIdx], tabName, isRosterTab);

  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const studentId = extractStrictColumnAStudentId(row);
    if (!studentId) continue;

    let student = studentsMap.get(studentId);
    if (!student) {
      if (!isRosterTab) continue; // Assessment tabs do not create unrostered students
      student = {
        id: studentId,
        studentId: studentId,
        name: '',
        grade: '',
        school: '',
        site: '',
        staffName: '',
        staffEmail: '',
        teacher: '',
        teacherEmail: '',
        fosterHomeless: '',
        specialEd: '',
        plan504: '',
        elStatus: '',
        assessmentsByTab: {},
        metrics: {}
      };
      studentsMap.set(studentId, student);
    }

    if (!student.assessmentsByTab[tabName]) {
      student.assessmentsByTab[tabName] = {};
    }

    for (let c = 1; c < headerAnalysis.length && c < row.length; c++) {
      const colMeta = headerAnalysis[c];
      const val = row[c] !== null && row[c] !== undefined ? row[c] : '';
      const strVal = String(val).trim();
      if (strVal === '') continue;

      if (isRosterTab) {
        // Standard Roster column mapping
        // Col 1: Name, Col 2: Grade, Col 3: School, Col 4: Teacher Name, Col 5: Teacher Email, Col 6: FH, Col 7: SPED, Col 8: 504, Col 9: EL
        if (c === 1 && !student.name) student.name = strVal;
        if (c === 2 && !student.grade) student.grade = strVal;
        if (c === 3 && !student.school && !isRawSchoolCodeOrPlaceholder_(strVal)) {
          student.school = strVal;
          student.site = strVal;
        }
        if (c === 4 && !student.staffName) {
          student.staffName = strVal;
          student.teacher = strVal;
        }
        if (c === 5 && !student.staffEmail) {
          student.staffEmail = strVal.toLowerCase();
          student.teacherEmail = strVal.toLowerCase();
        }
        if (c === 6 && !student.fosterHomeless) student.fosterHomeless = strVal;
        if (c === 7 && !student.specialEd) student.specialEd = strVal;
        if (c === 8 && !student.plan504) student.plan504 = strVal;
        if (c === 9 && !student.elStatus) student.elStatus = strVal;

        // Column O (col 14) authoritative school name override
        if (c === 14 && strVal && !isRawSchoolCodeOrPlaceholder_(strVal)) {
          student.school = strVal;
          student.site = strVal;
        }
      } else {
        // Assessment tabs: populate assessmentsByTab and metrics
        if (!SKIP_ASSESSMENT_MAP_HEADERS.has(colMeta.cleanHeader)) {
          student.assessmentsByTab[tabName][colMeta.rawHeader] = val;
          student.metrics[colMeta.rawHeader] = val;
        }
      }
    }
  }
}

function jsonResponse(data, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
