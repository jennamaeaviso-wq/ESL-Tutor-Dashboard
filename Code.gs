function doGet() {

  return HtmlService.createHtmlOutputFromFile('Index')

    .setTitle('ESL Teacher Dashboard')

    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}





/* =========================================================

   STUDENTS

\========================================================= */



function getStudents() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Students');



  if (!sheet) {

    createStudentsSheet_();

    return [];

  }



  ensureStudentPhotoColumn_();



  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];



  const lastColumn = Math.max(sheet.getLastColumn(), 11);



  const values = sheet

    .getRange(2, 1, lastRow - 1, lastColumn)

    .getValues();



  return values

    .filter(row => row[1] !== '')

    .map(row => ({

      id: row[0],

      name: row[1],

      age: row[2],

      level: row[3],

      enrolledDate: formatDateForClient_(row[4]),

      materials: row[5] || '',

      packageName: row[6] || '',

      packageMinutes: row[7] || '',

      paymentStatus: row[8] || '',

      studentStatus: row[9] || '',

      photoUrl: convertStudentPhotoUrl_(row[10] || '')

    }));

}





function addStudent(studentData) {

  if (!studentData) {

    throw new Error('No student information was provided.');

  }



  const name = String(studentData.name || '').trim();



  if (!name) {

    throw new Error('Please enter the student name.');

  }



  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Students');



  if (!sheet) {

    createStudentsSheet_();

    sheet = ss.getSheetByName('Students');

  }



  ensureStudentPhotoColumn_();



  const lastRow = sheet.getLastRow();



  if (lastRow > 1) {

    const names = sheet

      .getRange(2, 2, lastRow - 1, 1)

      .getValues()

      .flat()

      .map(v => String(v).trim().toLowerCase());



    if (names.includes(name.toLowerCase())) {

      throw new Error('This student is already enrolled.');

    }

  }



  const id =

    'STU-' +

    Utilities.getUuid().substring(0, 8).toUpperCase();



  let photoUrl = '';



  if (studentData.photoData) {

    photoUrl = saveStudentPhoto_(

      studentData.photoData,

      studentData.photoName || name

    );

  }



  let materials = '';



  if (Array.isArray(studentData.materials)) {

    materials = studentData.materials

      .map(v => String(v).trim())

      .filter(Boolean)

      .join('\n');

  }



  sheet.appendRow([

    id,

    name,

    studentData.age || '',

    studentData.level || '',

    studentData.enrolledDate || '',

    materials,

    studentData.packageName || '',

    studentData.packageMinutes || '',

    studentData.paymentStatus || '',

    studentData.studentStatus || '',

    photoUrl

  ]);



  return {

    id: id,

    name: name,

    age: studentData.age || '',

    level: studentData.level || '',

    enrolledDate: studentData.enrolledDate || '',

    materials: materials,

    packageName: studentData.packageName || '',

    packageMinutes: studentData.packageMinutes || '',

    paymentStatus: studentData.paymentStatus || '',

    studentStatus: studentData.studentStatus || '',

    photoUrl: photoUrl

  };

}





function updateStudent(studentData) {
  if (!studentData || !studentData.id) {
    throw new Error('Student ID is required.');
  }

  const name = String(studentData.name || '').trim();
  if (!name) {
    throw new Error('Please enter the student name.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Students');
  if (!sheet) {
    throw new Error('Students sheet was not found.');
  }

  ensureStudentPhotoColumn_();

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    throw new Error('No students found.');
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const rowIndex = rows.findIndex(row => String(row[0]) === String(studentData.id));

  if (rowIndex === -1) {
    throw new Error('Student not found.');
  }

  const duplicateName = rows.some((row, i) =>
    i !== rowIndex &&
    String(row[1] || '').trim().toLowerCase() === name.toLowerCase()
  );

  if (duplicateName) {
    throw new Error('This student is already enrolled.');
  }

  const rowNumber = rowIndex + 2;
  let photoUrl = rows[rowIndex][10] || '';

  if (studentData.photoData) {
    photoUrl = saveStudentPhoto_(
      studentData.photoData,
      studentData.photoName || name
    );
  }

  let materials = '';
  if (Array.isArray(studentData.materials)) {
    materials = studentData.materials
      .map(v => String(v).trim())
      .filter(Boolean)
      .join('\n');
  }

  sheet.getRange(rowNumber, 1, 1, 11).setValues([[
    studentData.id,
    name,
    studentData.age || '',
    studentData.level || '',
    studentData.enrolledDate || '',
    materials,
    studentData.packageName || '',
    studentData.packageMinutes || '',
    studentData.paymentStatus || '',
    studentData.studentStatus || '',
    photoUrl
  ]]);

  return {
    success: true,
    id: studentData.id
  };
}


function deleteStudent(studentId) {

  if (!studentId) {
    throw new Error('Student ID is required.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Students');

  if (!sheet) {
    throw new Error('Students sheet was not found.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    throw new Error('No students found.');
  }

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();

  const index = ids.findIndex(
    id => String(id) === String(studentId)
  );

  if (index === -1) {
    throw new Error('Student not found.');
  }

  const rowNumber = index + 2;

  sheet.deleteRow(rowNumber);

  return {
    success: true,
    id: studentId
  };
}


function createStudentsSheet_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Students');



  if (!sheet) {

    sheet = ss.insertSheet('Students');

  }



  if (sheet.getLastRow() === 0) {

    sheet.getRange(1, 1, 1, 11).setValues([[

      'ID',

      'Student Name',

      'Age',

      'Level',

      'Student Enrolled Date',

      'Course Materials',

      'Package Name',

      'Package Minutes',

      'Payment Status',

      'Student Status',

      'Photo URL'

    ]]);



    styleHeader_(sheet, 11);



    sheet.setFrozenRows(1);



    sheet.setColumnWidth(1, 125);

    sheet.setColumnWidth(2, 180);

    sheet.setColumnWidth(3, 70);

    sheet.setColumnWidth(4, 150);

    sheet.setColumnWidth(5, 150);

    sheet.setColumnWidth(6, 240);

    sheet.setColumnWidth(7, 170);

    sheet.setColumnWidth(8, 130);

    sheet.setColumnWidth(9, 130);

    sheet.setColumnWidth(10, 130);

    sheet.setColumnWidth(11, 300);



    sheet.getRange('F:F').setWrap(true);

  }

}





/* =========================================================

   STUDENT PHOTO

\========================================================= */



function ensureStudentPhotoColumn_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Students');



  if (!sheet) {

    createStudentsSheet_();

    sheet = ss.getSheetByName('Students');

  }



  const lastColumn = Math.max(sheet.getLastColumn(), 1);



  const headers = sheet

    .getRange(1, 1, 1, lastColumn)

    .getValues()[0]

    .map(v => String(v).trim());



  const photoIndex = headers.findIndex(

    h => h.toLowerCase() === 'photo url'

  );



  if (photoIndex === -1) {

    sheet.getRange(1, 11).setValue('Photo URL');



    sheet.getRange(1, 11)

      .setFontWeight('bold')

      .setBackground('#DDEBD7')

      .setFontColor('#52796F');



    sheet.setColumnWidth(11, 300);

  }

}





/*

 * Saves the uploaded student photo to Google Drive

 * and returns a direct thumbnail URL that can be

 * displayed by the web app.

 */

function saveStudentPhoto_(dataUrl, fileName) {

  try {

    const match = String(dataUrl).match(

      /^data:(image\/[^;]+);base64,(.+)$/

    );



    if (!match) {

      throw new Error('Invalid image data.');

    }



    const mimeType = match[1];

    const base64 = match[2];



    const bytes = Utilities.base64Decode(base64);



    let extension = '.jpg';



    if (mimeType === 'image/png') {

      extension = '.png';

    } else if (mimeType === 'image/webp') {

      extension = '.webp';

    } else if (mimeType === 'image/gif') {

      extension = '.gif';

    }



    const safeName =

      String(fileName)

        .replace(/[^\w\s-]/g, '')

        .trim() || 'student-photo';



    const blob = Utilities.newBlob(

      bytes,

      mimeType,

      safeName + extension

    );



    const folders =

      DriveApp.getFoldersByName('ESL Student Photos');



    let folder;



    if (folders.hasNext()) {

      folder = folders.next();

    } else {

      folder =

        DriveApp.createFolder('ESL Student Photos');

    }



    const file = folder.createFile(blob);



    /*
     * Make the saved photo viewable by the web app.
     * If the Google Workspace account does not allow
     * public link sharing, the file remains private and
     * the app will still use the thumbnail URL while the
     * teacher is signed into the same Google account.
     */
    try {
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
    } catch (sharingError) {
      // Some Google Workspace domains disable public sharing.
      // Do not stop the student from being saved.
    }

    const fileId = file.getId();

    return (
      'https://drive.google.com/thumbnail?id=' +
      encodeURIComponent(fileId) +
      '&sz=w400'
    );



  } catch (error) {

    throw new Error(

      'The student photo could not be uploaded: ' +

      error.message

    );

  }

}





/*

 * Converts old saved Google Drive URLs into

 * thumbnail URLs so previously uploaded photos

 * can also display correctly.

 */

function convertStudentPhotoUrl_(url) {

  if (!url) return '';



  const value = String(url).trim();



  if (

    value.indexOf(

      'https://drive.google.com/thumbnail'

    ) === 0

  ) {

    return value;

  }



  /*

   * Old Drive file URL:

   * https://drive.google.com/file/d/FILE_ID/view

   */

  const fileMatch =

    value.match(/\/file\/d\/([^/]+)/);



  if (fileMatch) {

    return (

      'https://drive.google.com/thumbnail?id=' +

      encodeURIComponent(fileMatch[1]) +

      '&sz=w400'

    );

  }



  /*

   * If the cell contains only a Drive file ID.

   */

  if (/^[a-zA-Z0-9_-]{20,}$/.test(value)) {

    return (

      'https://drive.google.com/thumbnail?id=' +

      encodeURIComponent(value) +

      '&sz=w400'

    );

  }



  return value;

}





/* =========================================================

   DRIVE AUTHORIZATION

\========================================================= */



function authorizeDrive() {

  const folders =

    DriveApp.getFoldersByName('ESL Student Photos');



  if (folders.hasNext()) {

    return 'ESL Student Photos folder already exists.';

  }



  const folder =

    DriveApp.createFolder('ESL Student Photos');



  return 'ESL Student Photos folder created successfully.';

}





/* =========================================================

   CLASSES

\========================================================= */



/* =========================================================
   PRIMARY CLASS STORE
   The web app's Script Properties store is authoritative.
   The Google Sheet is kept as a backup and migration source only.
========================================================= */

const ESL_CLASSES_PRIMARY_INITIALIZED_KEY_ = 'ESL_CLASSES_PRIMARY_INITIALIZED_V1';
const ESL_CLASSES_PRIMARY_CHUNK_COUNT_KEY_ = 'ESL_CLASSES_PRIMARY_CHUNK_COUNT_V1';
const ESL_CLASSES_PRIMARY_CHUNK_PREFIX_ = 'ESL_CLASSES_PRIMARY_CHUNK_V1_';
const ESL_CLASSES_PRIMARY_CHUNK_SIZE_ = 1800;

function ESL_readClassesBackup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Classes');
  if (!sheet) {
    createClassesSheet_();
    sheet = ss.getSheetByName('Classes');
  }
  ensureClassesStatusColumn_();
  ensureClassLearningColumns_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 14).getValues()
    .filter(row => String(row[1] || '').trim() !== '')
    .map(row => ({
      id: String(row[0] || ''),
      student: String(row[1] || ''),
      classDateStarts: formatDateForClient_(row[2]),
      totalMinutes: Number(row[3] || 0),
      recurringDays: row[4] == null ? '' : String(row[4]),
      timeType: String(row[5] || ''),
      startTime: formatTimeForClient_(row[6]),
      endTime: formatTimeForClient_(row[7]),
      createdAt: row[8] instanceof Date ? row[8].toISOString() : (row[8] || ''),
      status: String(row[9] || 'Scheduled'),
      courseMaterial: String(row[10] || ''),
      lastTopicConducted: String(row[11] || ''),
      minutesConducted: row[12] === '' || row[12] == null ? '' : Number(row[12] || 0),
      sessionDetails: (() => {
        try { return row[13] ? JSON.parse(String(row[13])) : {}; }
        catch (e) { return {}; }
      })()
    }));
}

function ESL_readPrimaryClasses_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(ESL_CLASSES_PRIMARY_INITIALIZED_KEY_) !== '1') {
    // One-time migration: import existing schedules before treating the sheet as backup only.
    const legacyRecords = ESL_readClassesBackup_();
    ESL_writePrimaryClasses_(legacyRecords);
    props.setProperty(ESL_CLASSES_PRIMARY_INITIALIZED_KEY_, '1');
    return legacyRecords;
  }

  const count = Number(props.getProperty(ESL_CLASSES_PRIMARY_CHUNK_COUNT_KEY_) || 0);
  if (!count) return [];
  let json = '';
  for (let i = 0; i < count; i++) {
    const chunk = props.getProperty(ESL_CLASSES_PRIMARY_CHUNK_PREFIX_ + i);
    if (chunk == null) throw new Error('The primary class data is incomplete. Please do not edit the Classes backup sheet.');
    json += chunk;
  }
  const records = JSON.parse(json);
  if (!Array.isArray(records)) throw new Error('The primary class data is invalid.');
  return records;
}

function ESL_writePrimaryClasses_(records) {
  const props = PropertiesService.getScriptProperties();
  const json = JSON.stringify(Array.isArray(records) ? records : []);
  const chunks = [];
  for (let i = 0; i < json.length; i += ESL_CLASSES_PRIMARY_CHUNK_SIZE_) {
    chunks.push(json.substring(i, i + ESL_CLASSES_PRIMARY_CHUNK_SIZE_));
  }
  const previousCount = Number(props.getProperty(ESL_CLASSES_PRIMARY_CHUNK_COUNT_KEY_) || 0);
  chunks.forEach((chunk, index) => props.setProperty(ESL_CLASSES_PRIMARY_CHUNK_PREFIX_ + index, chunk));
  for (let i = chunks.length; i < previousCount; i++) props.deleteProperty(ESL_CLASSES_PRIMARY_CHUNK_PREFIX_ + i);
  props.setProperty(ESL_CLASSES_PRIMARY_CHUNK_COUNT_KEY_, String(chunks.length));
  props.setProperty(ESL_CLASSES_PRIMARY_INITIALIZED_KEY_, '1');
}

function ESL_syncClassesBackup_(records) {
  // Backup failures must not undo a successful primary web-app save.
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Classes');
    if (!sheet) {
      createClassesSheet_();
      sheet = ss.getSheetByName('Classes');
    }
    ensureClassesStatusColumn_();
    ensureClassLearningColumns_();
    sheet.getRange(1, 1, 1, 14).setValues([[
      'ID', 'Student', 'Class Date Starts', 'Total Minutes', 'Recurring Days',
      'Time Type', 'Start Time', 'End Time', 'Created At', 'Status',
      'Course Material', 'Last Topic Conducted', 'Minutes Conducted', 'Session Details JSON'
    ]]);
    styleHeader_(sheet, 14);
    const existingRows = Math.max(0, sheet.getLastRow() - 1);
    if (existingRows) sheet.getRange(2, 1, existingRows, 14).clearContent();
    const rows = (records || []).map(cls => [
      cls.id || '', cls.student || '', cls.classDateStarts || '', Number(cls.totalMinutes || 0),
      Array.isArray(cls.recurringDays) ? cls.recurringDays.join(', ') : (cls.recurringDays || ''),
      cls.timeType || '', cls.startTime || '', cls.endTime || '', cls.createdAt || '',
      cls.status || 'Scheduled', cls.courseMaterial || '', cls.lastTopicConducted || '',
      cls.minutesConducted === '' || cls.minutesConducted == null ? '' : Number(cls.minutesConducted || 0),
      JSON.stringify(cls.sessionDetails || {})
    ]);
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, 14).setValues(rows);
      sheet.getRange(2, 3, rows.length, 1).setNumberFormat('@');
      sheet.getRange(2, 4, rows.length, 1).setNumberFormat('0');
      sheet.getRange(2, 7, rows.length, 2).setNumberFormat('@');
      sheet.getRange(2, 9, rows.length, 1).setNumberFormat('@');
      sheet.getRange(2, 13, rows.length, 1).setNumberFormat('0');
      sheet.getRange(2, 14, rows.length, 1).setNumberFormat('@');
    }
    SpreadsheetApp.flush();
  } catch (error) {
    console.error('Class backup sync failed: ' + (error && error.message ? error.message : error));
  }
}

function ESL_savePrimaryClassesAndBackup_(records) {
  ESL_writePrimaryClasses_(records);
  ESL_syncClassesBackup_(records);
  return records;
}

function getClasses() {
  return ESL_readPrimaryClasses_().map(cls => ({
    id: cls.id || '', student: cls.student || '',
    classDateStarts: formatDateForClient_(cls.classDateStarts),
    totalMinutes: Number(cls.totalMinutes || 0),
    recurringDays: Array.isArray(cls.recurringDays) ? cls.recurringDays.join(', ') : (cls.recurringDays || ''),
    timeType: cls.timeType || '', startTime: formatTimeForClient_(cls.startTime),
    endTime: formatTimeForClient_(cls.endTime), createdAt: cls.createdAt || '',
    status: cls.status || 'Scheduled', courseMaterial: cls.courseMaterial || '',
    lastTopicConducted: cls.lastTopicConducted || '',
    minutesConducted: cls.minutesConducted == null ? '' : cls.minutesConducted,
    sessionDetails: cls.sessionDetails && typeof cls.sessionDetails === 'object' ? cls.sessionDetails : {}
  }));
}

function addClass(classData) {
  if (!classData) throw new Error('No class information was provided.');

  const student = String(classData.student || '').trim();
  const classDateStarts = String(classData.classDateStarts || '').trim();

  if (!student) throw new Error('Please select a student.');
  if (!classDateStarts) throw new Error('Please select the class start date.');

  /*
   * Students > Package Minutes is the authoritative source for the
   * Total Minutes value when a new class is saved.
   */
  let totalMinutes = Number(classData.totalMinutes || 0);

  try {
    const students = getStudents();
    const selectedStudent = students.find(s => String(s.name) === student);

    if (selectedStudent && Number(selectedStudent.packageMinutes || 0) > 0) {
      totalMinutes = Number(selectedStudent.packageMinutes);
    }
  } catch (studentLookupError) {
    // Use the submitted value as a fallback if the lookup cannot run.
  }

  if (!Number.isFinite(totalMinutes) || totalMinutes <= 0) {
    throw new Error('The selected student does not have valid Package Minutes in the Students tab.');
  }

  const records = ESL_readPrimaryClasses_();

  const recurringDays = Array.isArray(classData.recurringDays)
    ? classData.recurringDays.join(', ')
    : String(classData.recurringDays || '');

  const timeType = String(classData.timeType || '');

  const dayTimes = classData.dayTimes && typeof classData.dayTimes === 'object'
    ? classData.dayTimes
    : {};

  const record = {
    id: 'CLS-' + Utilities.getUuid().substring(0, 8).toUpperCase(),
    student: student,
    classDateStarts: classDateStarts,
    totalMinutes: totalMinutes,
    recurringDays: recurringDays,
    timeType: timeType,
    startTime: String(classData.startTime || ''),
    endTime: String(classData.endTime || ''),
    createdAt: new Date().toISOString(),
    status: 'Scheduled',
    courseMaterial: '',
    lastTopicConducted: '',
    minutesConducted: '',
    sessionDetails: timeType === 'Flexible Time Per Day'
      ? { scheduleTimes: dayTimes }
      : {}
  };

  records.push(record);
  ESL_savePrimaryClassesAndBackup_(records);

  return {
    success: true,
    id: record.id,
    class: record
  };
}


function ensureClassesStatusColumn_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Classes');



  if (!sheet) {

    createClassesSheet_();

    return;

  }



  const lastColumn =

    Math.max(sheet.getLastColumn(), 1);



  const headers = sheet

    .getRange(1, 1, 1, lastColumn)

    .getValues()[0]

    .map(v => String(v).trim());



  const statusIndex = headers.findIndex(

    h => h.toLowerCase() === 'status'

  );



  if (statusIndex === -1) {

    sheet.getRange(1, 10).setValue('Status');



    sheet.getRange(1, 10)

      .setFontWeight('bold')

      .setBackground('#DDEBD7')

      .setFontColor('#52796F');



    sheet.setColumnWidth(10, 140);

  }



  const lastRow = sheet.getLastRow();



  if (lastRow > 1) {

    const statusRange =

      sheet.getRange(2, 10, lastRow - 1, 1);



    const statuses = statusRange.getValues();



    let changed = false;



    statuses.forEach(row => {

      if (!row[0]) {

        row[0] = 'Scheduled';

        changed = true;

      }

    });



    if (changed) {

      statusRange.setValues(statuses);

    }

  }

}





function ensureClassLearningColumns_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Classes');

  if (!sheet) {
    createClassesSheet_();
    return;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(v => String(v).trim().toLowerCase());

  const columns = [
    ['Course Material', 11, 150],
    ['Last Topic Conducted', 12, 220],
    ['Minutes Conducted', 13, 150],
    ['Session Details JSON', 14, 300]
  ];

  columns.forEach(([label, column, width]) => {
    if (!headers.includes(label.toLowerCase())) {
      sheet.getRange(1, column).setValue(label);
      sheet.getRange(1, column)
        .setFontWeight('bold')
        .setBackground('#DDEBD7')
        .setFontColor('#52796F');
      sheet.setColumnWidth(column, width);
    }
  });
}


function saveClassLearningDetails(classId, courseMaterial, lastTopicConducted, minutesConducted, eventDate) {
  if (!classId) throw new Error('Class ID is required.');
  const records = ESL_readPrimaryClasses_();
  const index = records.findIndex(cls => String(cls.id) === String(classId));
  if (index < 0) throw new Error('Class not found. Please reload the Dashboard.');

  const scheduledMinutes = Number(records[index].totalMinutes || 0);
  const conducted = Number(minutesConducted || 0);
  if (!Number.isFinite(conducted) || conducted < 0 || conducted > scheduledMinutes) {
    throw new Error('Minutes conducted cannot be greater than the scheduled minutes.');
  }

  const material = String(courseMaterial || '').trim();
  const topic = String(lastTopicConducted || '').trim();
  const dateKey = String(eventDate || '').trim();

  if (dateKey) {
    // Save each occurrence separately so recurring classes can have different
    // topics and minutes for each date.
    if (!records[index].sessionDetails || typeof records[index].sessionDetails !== 'object') {
      records[index].sessionDetails = {};
    }
    records[index].sessionDetails[dateKey] = {
      courseMaterial: material,
      lastTopicConducted: topic,
      minutesConducted: conducted
    };
  } else {
    // Backward compatibility for callers that save class-level details.
    records[index].courseMaterial = material;
    records[index].lastTopicConducted = topic;
    records[index].minutesConducted = conducted;
  }

  ESL_savePrimaryClassesAndBackup_(records);
  return {
    success: true, id: classId, eventDate: dateKey,
    courseMaterial: material, lastTopicConducted: topic, minutesConducted: conducted
  };
}

function saveCompletedClassMinutes(classId, eventDate, minutesConducted) {
  if (!classId) throw new Error('Class ID is required.');
  const dateKey = String(eventDate || '').trim();
  if (!dateKey) throw new Error('Class date is required.');

  const records = ESL_readPrimaryClasses_();
  const index = records.findIndex(cls => String(cls.id) === String(classId));
  if (index < 0) throw new Error('Class not found. Please reload the app.');

  const scheduledMinutes = Number(records[index].totalMinutes || 0);
  const conducted = Number(minutesConducted);
  if (!Number.isFinite(conducted) || conducted < 0 || !Number.isInteger(conducted) || conducted > scheduledMinutes) {
    throw new Error('Please enter a whole number of minutes between 0 and the scheduled class duration.');
  }

  if (!records[index].sessionDetails || typeof records[index].sessionDetails !== 'object') {
    records[index].sessionDetails = {};
  }
  const previous = records[index].sessionDetails[dateKey] || {};
  records[index].sessionDetails[dateKey] = Object.assign({}, previous, {
    minutesConducted: conducted,
    status: 'Completed / Finished'
  });

  ESL_savePrimaryClassesAndBackup_(records);
  return { success: true, id: classId, eventDate: dateKey, status: 'Completed / Finished', minutesConducted: conducted };
}


function updateClassStatus(classId, newStatus) {
  const allowedStatuses = ['Scheduled', 'Completed / Finished', 'Cancelled', 'Reschedule'];
  if (!allowedStatuses.includes(newStatus)) throw new Error('Invalid class status.');
  if (!classId) throw new Error('Class ID is required.');
  const records = ESL_readPrimaryClasses_();
  const index = records.findIndex(cls => String(cls.id) === String(classId));
  if (index < 0) throw new Error('Class not found. Please reload the Dashboard.');
  records[index].status = newStatus;
  ESL_savePrimaryClassesAndBackup_(records);
  return { success: true, id: classId, status: newStatus };
}


function deleteClass(classId) {
  if (!classId) throw new Error('Class ID is required.');

  const records = ESL_readPrimaryClasses_();
  const index = records.findIndex(cls => String(cls.id) === String(classId));

  if (index < 0) {
    throw new Error('Class not found. Please reload the Dashboard.');
  }

  records.splice(index, 1);
  ESL_savePrimaryClassesAndBackup_(records);

  return {
    success: true,
    id: classId
  };
}


function createClassesSheet_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Classes');



  if (!sheet) {

    sheet = ss.insertSheet('Classes');

  }



  if (sheet.getLastRow() === 0) {

    sheet.getRange(1, 1, 1, 14).setValues([[
      'ID',
      'Student',
      'Class Date Starts',
      'Total Minutes',
      'Recurring Days',
      'Time Type',
      'Start Time',
      'End Time',
      'Created At',
      'Status',
      'Course Material',
      'Last Topic Conducted',
      'Minutes Conducted',
      'Session Details JSON'
    ]]);

    styleHeader_(sheet, 14);



    sheet.setFrozenRows(1);



    sheet.setColumnWidth(1, 125);

    sheet.setColumnWidth(2, 180);

    sheet.setColumnWidth(3, 140);

    sheet.setColumnWidth(4, 110);

    sheet.setColumnWidth(5, 180);

    sheet.setColumnWidth(6, 120);

    sheet.setColumnWidth(7, 100);

    sheet.setColumnWidth(8, 100);

    sheet.setColumnWidth(9, 150);

    sheet.setColumnWidth(10, 150);

  } else {

    ensureClassesStatusColumn_();

  }

}





/* =========================================================

   HELPERS

\========================================================= */



function styleHeader_(sheet, columns) {

  sheet.getRange(1, 1, 1, columns)

    .setFontWeight('bold')

    .setBackground('#DDEBD7')

    .setFontColor('#52796F');

}





function formatDateForClient_(value) {

  if (!value) return '';



  if (

    Object.prototype.toString.call(value) ===

    '[object Date]'

  ) {

    return Utilities.formatDate(

      value,

      Session.getScriptTimeZone(),

      'yyyy-MM-dd'

    );

  }



  return String(value);

}





function formatTimeForClient_(value) {

  if (!value) return '';



  if (

    Object.prototype.toString.call(value) ===

    '[object Date]'

  ) {

    return Utilities.formatDate(

      value,

      Session.getScriptTimeZone(),

      'HH:mm'

    );

  }



  return String(value);

}

/** Deletes a class from the primary web-app store and then refreshes the backup sheet. */
function deleteClass(classId) {
  if (!classId) throw new Error('No class ID was provided.');
  const records = ESL_readPrimaryClasses_();
  const index = records.findIndex(cls => String(cls.id) === String(classId));
  if (index < 0) throw new Error('Class not found. Please reload the Dashboard.');
  records.splice(index, 1);
  ESL_savePrimaryClassesAndBackup_(records);
  return { success: true, id: classId };
}
