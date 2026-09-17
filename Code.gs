/**
 * ระบบทำข้อสอบวิทยาศาสตร์ ม.ปลาย (Science Examination Web App)
 * Google Apps Script Backend (Code.gs)
 * 
 * รองรับ:
 * - แยกหมวดวิชา ม.ปลาย: ชีววิทยา (Biology), เคมี (Chemistry), ฟิสิกส์ (Physics)
 * - ระบบกู้คืนรหัสผ่านด้วย PIN 4 หลัก และคำถามกันลืม (Privacy/PDPA Compliant)
 * - ระบบแยกบัญชีครูออกจากกันโดยเด็ดขาด (Strict Teacher Isolation)
 * - การสลับตัวเลือก ก-ข-ค-ง และสลับข้อ (Choice & Question Shuffling)
 * - นาฬิกาจับเวลา และระบบรีเซ็ตข้อสอบเมื่อสลับหน้าจอ
 */

const SHEET_NAMES = {
  TEACHERS: 'Teachers',
  ROOMS: 'Rooms',
  QUESTIONS: 'Questions',
  SESSIONS: 'Sessions',
  RESULTS: 'Results'
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบทำข้อสอบวิทยาศาสตร์ ม.ปลาย')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * กำหนดค่าและเตรียมชีตเริ่มต้น
 */
function setupRoomSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Teachers
  const teacherSheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
  if (teacherSheet.getLastRow() === 0) {
    teacherSheet.appendRow(['TeacherID', 'FullName', 'Username', 'PasswordHash', 'RecoveryPin', 'SecurityQuestion', 'SecurityAnswerHash', 'CreatedAt']);
  }

  // 2. Rooms
  const roomSheet = getOrCreateSheet(SHEET_NAMES.ROOMS);
  if (roomSheet.getLastRow() === 0) {
    roomSheet.appendRow(['RoomID', 'TeacherID', 'RoomCode', 'RoomName', 'Subject', 'Classroom', 'EasyCount', 'MediumCount', 'HardCount', 'TimeLimit', 'IsActive', 'CreatedAt']);
  }

  // 3. Questions (Question Bank ม.ปลาย)
  const questionSheet = getOrCreateSheet(SHEET_NAMES.QUESTIONS);
  if (questionSheet.getLastRow() === 0) {
    questionSheet.appendRow(['QuestionID', 'Subject', 'Difficulty', 'Question', 'ChoiceA', 'ChoiceB', 'ChoiceC', 'ChoiceD', 'CorrectAnswer', 'Explanation']);
    seedDefaultHighSchoolQuestions(questionSheet);
  }

  // 4. Sessions
  const sessionSheet = getOrCreateSheet(SHEET_NAMES.SESSIONS);
  if (sessionSheet.getLastRow() === 0) {
    sessionSheet.appendRow(['Token', 'Type', 'UserID', 'RoomCode', 'DataJSON', 'CreatedAt']);
  }

  // 5. Results
  const resultSheet = getOrCreateSheet(SHEET_NAMES.RESULTS);
  if (resultSheet.getLastRow() === 0) {
    resultSheet.appendRow(['ResultID', 'RoomID', 'StudentName', 'Classroom', 'Number', 'Score', 'Total', 'Percentage', 'SwitchCount', 'Timestamp']);
  }

  return 'Setup completed successfully!';
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function hashPassword(text) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''), Utilities.Charset.UTF_8);
  return digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function normalizeText(text) {
  return String(text || '').trim();
}

function shuffle(array) {
  const arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// -------------------------------------------------------------
// Teacher Authentication & Password Recovery (No Email Required)
// -------------------------------------------------------------

function registerTeacher(fullName, username, password, recoveryPin, securityQuestion, securityAnswer) {
  fullName = normalizeText(fullName);
  username = normalizeText(username).toLowerCase();
  password = normalizeText(password);
  recoveryPin = normalizeText(recoveryPin);
  securityQuestion = normalizeText(securityQuestion);
  securityAnswer = normalizeText(securityAnswer).toLowerCase();

  if (!fullName || !username || !password || !recoveryPin || !securityQuestion || !securityAnswer) {
    throw new Error('กรุณากรอกข้อมูลให้ครบถ้วนทุกช่อง');
  }
  if (password.length < 6) {
    throw new Error('รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร');
  }
  if (!/^\d{4}$/.test(recoveryPin)) {
    throw new Error('PIN กู้คืนต้องเป็นตัวเลข 4 หลักเท่านั้น (เช่น 1234)');
  }

  const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === username) {
      throw new Error('ชื่อผู้ใช้นี้มีอยู่ในระบบแล้ว กรุณาใช้ชื่ออื่น');
    }
  }

  const teacherId = 'T-' + Utilities.getUuid().slice(0, 8);
  const passHash = hashPassword(password);
  const answerHash = hashPassword(securityAnswer);

  sheet.appendRow([
    teacherId,
    fullName,
    username,
    passHash,
    recoveryPin,
    securityQuestion,
    answerHash,
    new Date()
  ]);

  return 'สมัครบัญชีครูสำเร็จ! สามารถเข้าสู่ระบบได้ทันที';
}

function teacherLogin(username, password) {
  username = normalizeText(username).toLowerCase();
  password = normalizeText(password);

  const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
  const data = sheet.getDataRange().getValues();
  const passHash = hashPassword(password);

  for (let i = 1; i < data.length; i++) {
    const [tId, fullName, uName, pHash] = data[i];
    if (String(uName).toLowerCase() === username && pHash === passHash) {
      const token = 'TOKEN-T-' + Utilities.getUuid();
      saveSession(token, 'TEACHER', tId, '', { fullName, username });
      return {
        token: token,
        fullName: fullName
      };
    }
  }

  throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
}

/**
 * ดึงคำถามกันลืมของครูจากชื่อผู้ใช้
 */
function getSecurityQuestion(username) {
  username = normalizeText(username).toLowerCase();
  if (!username) throw new Error('กรุณากรอกชื่อผู้ใช้');

  const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === username) {
      return {
        question: data[i][5] || 'สัตว์เลี้ยงตัวแรกของคุณชื่ออะไร?'
      };
    }
  }

  throw new Error('ไม่พบชื่อผู้ใช้นี้ในระบบ');
}

/**
 * ตรวจสอบ PIN 4 หลัก และคำตอบคำถามกันลืมเพื่อตั้งรหัสผ่านใหม่
 */
function resetPasswordWithSecurity(username, recoveryPin, securityAnswer, newPassword) {
  username = normalizeText(username).toLowerCase();
  recoveryPin = normalizeText(recoveryPin);
  securityAnswer = normalizeText(securityAnswer).toLowerCase();
  newPassword = normalizeText(newPassword);

  if (!username || !recoveryPin || !securityAnswer || !newPassword) {
    throw new Error('กรุณากรอกข้อมูลให้ครบทุกช่อง');
  }
  if (newPassword.length < 6) {
    throw new Error('รหัสผ่านใหม่ต้องมีความยาวอย่างน้อย 6 ตัวอักษร');
  }

  const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
  const data = sheet.getDataRange().getValues();
  const answerHash = hashPassword(securityAnswer);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === username) {
      const savedPin = String(data[i][4]);
      const savedAnswerHash = String(data[i][6]);

      if (savedPin !== recoveryPin) {
        throw new Error('รหัส PIN กู้คืน 4 หลักไม่ถูกต้อง');
      }
      if (savedAnswerHash !== answerHash) {
        throw new Error('คำตอบสำหรับคำถามกันลืมไม่ถูกต้อง');
      }

      // บันทึกรหัสผ่านใหม่
      const newPassHash = hashPassword(newPassword);
      sheet.getRange(i + 1, 4).setValue(newPassHash);
      return 'ตั้งรหัสผ่านใหม่สำเร็จ! กรุณาเข้าสู่ระบบด้วยรหัสผ่านใหม่';
    }
  }

  throw new Error('ไม่พบชื่อผู้ใช้นี้');
}

// -------------------------------------------------------------
// Exam Rooms Management (Strict Isolation by Teacher ID)
// -------------------------------------------------------------

function generateRoomCode(subjectPrefix) {
  let prefix = 'SCI-';
  if (subjectPrefix === 'biology') prefix = 'BIO-';
  else if (subjectPrefix === 'chemistry') prefix = 'CHM-';
  else if (subjectPrefix === 'physics') prefix = 'PHY-';

  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = prefix;
  for (let i = 0; i < 5; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

function createRoom(token, name, subject, classroom, easy, medium, hard, timeLimit) {
  const session = requireTeacher(token);

  name = normalizeText(name);
  subject = normalizeText(subject).toLowerCase() || 'general';
  classroom = normalizeText(classroom);
  easy = parseInt(easy, 10) || 0;
  medium = parseInt(medium, 10) || 0;
  hard = parseInt(hard, 10) || 0;
  timeLimit = parseInt(timeLimit, 10);
  if (isNaN(timeLimit) || timeLimit < 0) timeLimit = 15;

  if (!name || !classroom) {
    throw new Error('กรุณาระบุชื่อห้องสอบและชั้น/ห้อง');
  }
  if (easy + medium + hard <= 0) {
    throw new Error('กรุณาระบุจำนวนข้อสอบอย่างน้อย 1 ข้อ');
  }

  const roomId = 'R-' + Utilities.getUuid().slice(0, 8);
  const code = generateRoomCode(subject);
  const roomSheet = getOrCreateSheet(SHEET_NAMES.ROOMS);

  roomSheet.appendRow([
    roomId,
    session.userId, // แยกตามครูแต่ละคนโดยเด็ดขาด
    code,
    name,
    subject,
    classroom,
    easy,
    medium,
    hard,
    timeLimit,
    true, // IsActive
    new Date()
  ]);

  return {
    id: roomId,
    code: code,
    accessCode: code,
    name: name,
    subject: subject,
    classroom: classroom,
    easy: easy,
    medium: medium,
    hard: hard,
    timeLimit: timeLimit
  };
}

/**
 * ดึงเฉพาะห้องสอบที่เป็นของครูผู้ที่ล็อกอินอยู่เท่านั้น (Strict Isolation)
 */
function getRooms(token) {
  const session = requireTeacher(token);
  const sheet = getOrCreateSheet(SHEET_NAMES.ROOMS);
  const data = sheet.getDataRange().getValues();
  const rooms = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[0];
    const teacherId = row[1];
    const code = row[2];
    const name = row[3];
    const subject = row.length >= 12 ? row[4] : 'general';
    const classroom = row.length >= 12 ? row[5] : row[4];
    const easy = row.length >= 12 ? row[6] : row[5];
    const medium = row.length >= 12 ? row[7] : row[6];
    const hard = row.length >= 12 ? row[8] : row[7];
    const timeLimit = row.length >= 12 ? Number(row[9]) : Number(row[8]);
    const isActive = row.length >= 12 ? row[10] === true : row[9] === true;

    // ตรวจสอบว่าเป็นห้องของครูคนนี้เท่านั้น
    if (teacherId === session.userId && isActive === true) {
      rooms.push({
        id: id,
        code: code,
        name: name,
        subject: subject,
        classroom: classroom,
        easy: Number(easy),
        medium: Number(medium),
        hard: Number(hard),
        timeLimit: timeLimit
      });
    }
  }

  return rooms.reverse();
}

/**
 * ปิดห้องสอบ (ตรวจสอบสิทธิ์ความเป็นเจ้าของก่อนเสมอ)
 */
function closeRoom(token, roomId) {
  const session = requireTeacher(token);
  const sheet = getOrCreateSheet(SHEET_NAMES.ROOMS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === roomId && data[i][1] === session.userId) {
      const activeColIndex = data[i].length >= 12 ? 11 : 10;
      sheet.getRange(i + 1, activeColIndex).setValue(false);
      return 'ปิดห้องสอบเรียบร้อยแล้ว';
    }
  }
  throw new Error('ไม่พบห้องสอบ หรือคุณไม่มีสิทธิ์ในการจัดการห้องสอบนี้');
}

/**
 * ดึงผลคะแนน (เฉพาะครูที่เป็นเจ้าของห้องสอบนี้เท่านั้น)
 */
function getRoomResults(token, roomId) {
  const session = requireTeacher(token);

  // ตรวจสอบว่าห้องนี้เป็นของครูคนนี้จริงหรือไม่
  const roomSheet = getOrCreateSheet(SHEET_NAMES.ROOMS);
  const roomData = roomSheet.getDataRange().getValues();
  let isOwner = false;
  for (let i = 1; i < roomData.length; i++) {
    if (roomData[i][0] === roomId && roomData[i][1] === session.userId) {
      isOwner = true;
      break;
    }
  }

  if (!isOwner) {
    throw new Error('คุณไม่มีสิทธิ์เข้าถึงผลการสอบของห้องนี้');
  }

  const sheet = getOrCreateSheet(SHEET_NAMES.RESULTS);
  const data = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const resId = row[0];
    const rId = row[1];
    const name = row[2];
    const classroom = row[3];
    const number = row[4];
    const score = row[5];
    const total = row[6];
    const percentage = row[7];
    const switchCount = Number(row[8]) || 0;
    const timestamp = row[9];

    if (rId === roomId) {
      results.push({
        id: resId,
        name: name,
        classroom: classroom,
        number: number,
        correct: score,
        total: total,
        percentage: percentage,
        switchCount: switchCount,
        time: timestamp instanceof Date ? Utilities.formatDate(timestamp, 'GMT+7', 'dd/MM/yyyy HH:mm') : String(timestamp)
      });
    }
  }

  return results.reverse();
}

/**
 * อนุญาตสอบใหม่ (ลบผลสอบเดิม) ตรวจสอบสิทธิ์ครูผู้เป็นเจ้าของห้อง
 */
function resetStudentResult(token, resultId) {
  const session = requireTeacher(token);
  const resSheet = getOrCreateSheet(SHEET_NAMES.RESULTS);
  const resData = resSheet.getDataRange().getValues();

  let targetRowIndex = -1;
  let roomId = '';

  for (let i = 1; i < resData.length; i++) {
    if (String(resData[i][0]) === String(resultId)) {
      targetRowIndex = i + 1;
      roomId = String(resData[i][1]);
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error('ไม่พบข้อมูลผลสอบนี้');
  }

  // ตรวจสอบสิทธิ์ความเป็นเจ้าของห้อง
  const roomSheet = getOrCreateSheet(SHEET_NAMES.ROOMS);
  const roomData = roomSheet.getDataRange().getValues();
  let isOwner = false;
  for (let i = 1; i < roomData.length; i++) {
    if (roomData[i][0] === roomId && roomData[i][1] === session.userId) {
      isOwner = true;
      break;
    }
  }

  if (!isOwner) {
    throw new Error('คุณไม่มีสิทธิ์แก้ไขผลสอบของห้องนี้');
  }

  resSheet.deleteRow(targetRowIndex);
  return 'ลบผลสอบเรียบร้อยแล้ว นักเรียนสามารถเข้าสอบใหม่ได้';
}

// -------------------------------------------------------------
// Student Flow & Exam Generation with Choice Shuffling
// -------------------------------------------------------------

function findRoomByCode(code) {
  code = normalizeText(code).toUpperCase();
  const sheet = getOrCreateSheet(SHEET_NAMES.ROOMS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[0];
    const roomCode = row[2];
    const name = row[3];
    const subject = row.length >= 12 ? row[4] : 'general';
    const classroom = row.length >= 12 ? row[5] : row[4];
    const easy = row.length >= 12 ? row[6] : row[5];
    const medium = row.length >= 12 ? row[7] : row[6];
    const hard = row.length >= 12 ? row[8] : row[7];
    const timeLimit = row.length >= 12 ? Number(row[9]) : Number(row[8]);
    const isActive = row.length >= 12 ? row[10] === true : row[9] === true;

    if (String(roomCode).toUpperCase() === code && isActive === true) {
      return {
        id: id,
        code: roomCode,
        name: name,
        subject: subject,
        classroom: classroom,
        easy: Number(easy),
        medium: Number(medium),
        hard: Number(hard),
        timeLimit: timeLimit
      };
    }
  }
  return null;
}

function validateRoomCode(code) {
  const room = findRoomByCode(code);
  if (!room) {
    throw new Error('ไม่พบรหัสห้องสอบ หรือห้องสอบนี้ถูกปิดแล้ว');
  }
  return room;
}

function checkStudentAlreadySubmitted(roomId, classroom, number) {
  const sheet = getOrCreateSheet(SHEET_NAMES.RESULTS);
  const data = sheet.getDataRange().getValues();

  classroom = normalizeText(classroom).toLowerCase();
  number = normalizeText(number);

  for (let i = 1; i < data.length; i++) {
    const rId = String(data[i][1]);
    const cClass = normalizeText(data[i][3]).toLowerCase();
    const cNum = normalizeText(data[i][4]);

    if (rId === roomId && cClass === classroom && cNum === number) {
      return true;
    }
  }
  return false;
}

function startStudentSession(roomCode, studentName, studentClass, studentNumber) {
  studentName = normalizeText(studentName);
  studentClass = normalizeText(studentClass);
  studentNumber = normalizeText(studentNumber);

  if (!studentName || !studentClass || !studentNumber) {
    throw new Error('กรุณากรอกข้อมูลนักเรียนให้ครบถ้วน');
  }

  const room = validateRoomCode(roomCode);

  if (checkStudentAlreadySubmitted(room.id, studentClass, studentNumber)) {
    throw new Error(`นักเรียนเลขที่ ${studentNumber} ชั้น ${studentClass} ได้ส่งข้อสอบในห้องนี้ไปแล้ว ไม่สามารถเข้าสอบซ้ำได้ (หากต้องการสอบใหม่กรุณาติดต่อครูผู้สอน)`);
  }

  const studentToken = 'TOKEN-S-' + Utilities.getUuid();

  saveSession(studentToken, 'STUDENT', studentName, room.code, {
    roomId: room.id,
    studentName: studentName,
    studentClass: studentClass,
    studentNumber: studentNumber
  });

  return {
    token: studentToken,
    room: room
  };
}

/**
 * สุ่มข้อสอบแยกตามวิชา ม.ปลาย พร้อมสลับตัวเลือก ก-ข-ค-ง (Choice Shuffling)
 */
function generateStudentExam(studentToken) {
  const session = requireStudent(studentToken);
  const room = findRoomByCode(session.roomCode);
  if (!room) {
    throw new Error('ห้องสอบไม่พร้อมใช้งาน');
  }

  const allQuestions = getQuestionBank();
  
  // กรองตามวิชาของห้องสอบ (biology, chemistry, physics, general)
  const subjectQuestions = allQuestions.filter(q => {
    if (room.subject && room.subject !== 'general') {
      return q.subject === room.subject;
    }
    return true; // ถ้าเป็น general ให้ใช้ข้อสอบรวมได้
  });

  const easyPool = subjectQuestions.filter(q => q.difficulty === 'easy');
  const mediumPool = subjectQuestions.filter(q => q.difficulty === 'medium');
  const hardPool = subjectQuestions.filter(q => q.difficulty === 'hard');

  const selectedQuestions = [
    ...shuffle(easyPool).slice(0, room.easy),
    ...shuffle(mediumPool).slice(0, room.medium),
    ...shuffle(hardPool).slice(0, room.hard)
  ];

  // 1. สลับลำดับข้อ (Question Shuffling)
  const shuffledExam = shuffle(selectedQuestions);

  // 2. สลับตัวเลือก ก-ข-ค-ง ของแต่ละข้อ (Choice Shuffling)
  const clientQuestions = shuffledExam.map(q => {
    const rawChoices = [q.choiceA, q.choiceB, q.choiceC, q.choiceD];
    const shuffledChoices = shuffle(rawChoices);

    return {
      id: q.id,
      question: q.question,
      subject: q.subject,
      difficulty: q.difficulty,
      choices: shuffledChoices // ส่งตัวเลือกที่สลับลำดับแล้วไปให้นักเรียน
    };
  });

  return {
    room: {
      name: room.name,
      subject: room.subject,
      classroom: room.classroom,
      code: room.code,
      timeLimit: room.timeLimit || 15
    },
    questions: clientQuestions
  };
}

/**
 * ตรวจข้อสอบด้วยข้อความตัวเลือก (Choice Text Matching) ป้องกันปัญหาตัวเลือกสลับที่
 */
function submitStudentExam(studentToken, answers, switchCount) {
  const session = requireStudent(studentToken);
  const room = findRoomByCode(session.roomCode);
  if (!room) {
    throw new Error('ห้องสอบไม่พร้อมใช้งาน');
  }

  if (checkStudentAlreadySubmitted(room.id, session.data.studentClass, session.data.studentNumber)) {
    throw new Error('คุณได้ส่งข้อสอบชุดนี้ไปแล้ว ไม่สามารถส่งซ้ำได้');
  }

  switchCount = parseInt(switchCount, 10) || 0;
  const allQuestions = getQuestionBank();
  const questionMap = new Map();
  allQuestions.forEach(q => questionMap.set(String(q.id), q));

  let correctCount = 0;
  const results = [];

  answers.forEach(ans => {
    const q = questionMap.get(String(ans.id));
    if (q) {
      // หาข้อความของคำตอบที่ถูกต้อง
      let correctText = '';
      if (q.correctAnswer === 'A') correctText = q.choiceA;
      else if (q.correctAnswer === 'B') correctText = q.choiceB;
      else if (q.correctAnswer === 'C') correctText = q.choiceC;
      else if (q.correctAnswer === 'D') correctText = q.choiceD;

      const isCorrect = normalizeText(ans.selectedText) === normalizeText(correctText);
      if (isCorrect) correctCount++;

      results.push({
        id: ans.id,
        isCorrect: isCorrect,
        correctText: correctText,
        explanation: q.explanation || ''
      });
    }
  });

  const total = answers.length;
  const percentage = total > 0 ? Math.round((correctCount / total) * 100) : 0;

  // บันทึกคะแนน
  const sheet = getOrCreateSheet(SHEET_NAMES.RESULTS);
  const resultId = 'RES-' + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([
    resultId,
    room.id,
    session.data.studentName,
    session.data.studentClass,
    session.data.studentNumber,
    correctCount,
    total,
    percentage,
    switchCount,
    new Date()
  ]);

  return {
    correct: correctCount,
    total: total,
    percentage: percentage,
    switchCount: switchCount,
    results: results
  };
}

// -------------------------------------------------------------
// Sessions & Utilities
// -------------------------------------------------------------

function saveSession(token, type, userId, roomCode, data) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SESSIONS);
  sheet.appendRow([token, type, userId, roomCode, JSON.stringify(data), new Date()]);
}

function getSession(token) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SESSIONS);
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === token) {
      let extra = {};
      try {
        extra = JSON.parse(data[i][4]);
      } catch (e) {}
      return {
        token: data[i][0],
        type: data[i][1],
        userId: data[i][2],
        roomCode: data[i][3],
        data: extra,
        createdAt: data[i][5]
      };
    }
  }
  return null;
}

function requireTeacher(token) {
  const session = getSession(token);
  if (!session || session.type !== 'TEACHER') {
    throw new Error('เซสชันครูหมดอายุหรือไม่มีสิทธิ์ กรุณาเข้าสู่ระบบใหม่');
  }
  return session;
}

function requireStudent(token) {
  const session = getSession(token);
  if (!session || session.type !== 'STUDENT') {
    throw new Error('เซสชันนักเรียนไม่ถูกต้อง กรุณาเข้าสู่ห้องสอบใหม่');
  }
  return session;
}

// -------------------------------------------------------------
// High School Science Question Bank (ม.ปลาย: ชีวะ เคมี ฟิสิกส์)
// -------------------------------------------------------------

function getQuestionBank() {
  const sheet = getOrCreateSheet(SHEET_NAMES.QUESTIONS);
  const data = sheet.getDataRange().getValues();
  const questions = [];

  for (let i = 1; i < data.length; i++) {
    const [id, subject, difficulty, qText, a, b, c, d, correct, exp] = data[i];
    if (qText) {
      questions.push({
        id: String(id),
        subject: String(subject).toLowerCase(),
        difficulty: String(difficulty).toLowerCase(),
        question: qText,
        choiceA: a,
        choiceB: b,
        choiceC: c,
        choiceD: d,
        correctAnswer: String(correct).toUpperCase(),
        explanation: exp || ''
      });
    }
  }

  return questions;
}

function seedDefaultHighSchoolQuestions(sheet) {
  const sample = [
    // --- 🧬 ชีววิทยา (Biology) ม.ปลาย ---
    ['BIO-101', 'biology', 'easy', 'ออร์แกเนลล์ใดทำหน้าที่เป็นแหล่งสร้างพลังงาน ATP หลักของเซลล์สิ่งมีชีวิต?', 'ไมโทคอนเดรีย (Mitochondria)', 'ไรโบโซม (Ribosome)', 'กอลจิบอดี (Golgi body)', 'ไลโซโซม (Lysosome)', 'A', 'ไมโทคอนเดรียเป็นโรงงานผลิตพลังงาน ATP ของเซลล์'],
    ['BIO-102', 'biology', 'easy', 'เบสในกรดนิวคลีอิกชนิดใดที่พบเฉพาะใน RNA แต่ไม่พบใน DNA?', 'ยูราซิล (Uracil)', 'ไทมีน (Thymine)', 'กวานีน (Guanine)', 'ไซโทซีน (Cytosine)', 'A', 'RNA จะมีเบส U แทนที่เบส T ใน DNA'],
    ['BIO-201', 'biology', 'medium', 'ระยะใดของการแบ่งเซลล์แบบไมโทซิส (Mitosis) ที่โครโมโซมเรียงตัวตรงกึ่งกลางเซลล์ชัดเจนที่สุด?', 'เมทาเฟส (Metaphase)', 'โพรเฟส (Prophase)', 'แอนาเฟส (Anaphase)', 'เทโลเฟส (Telophase)', 'A', 'Metaphase โครโมโซมจะเรียงตัวที่ Metaphase plate ตรงกลางเซลล์'],
    ['BIO-202', 'biology', 'medium', 'ปฏิกิริยาแสง (Light reaction) ในการสังเคราะห์ด้วยแสงเกิดขึ้นที่โครงสร้างใดของคลอโรพลาสต์?', 'ไทลาคอยด์ (Thylakoid)', 'สโตรมา (Stroma)', 'เยื่อหุ้มชั้นนอก', 'คริสตี (Cristae)', 'A', 'ปฏิกิริยาแสงเกิดขึ้นที่เยื่อไทลาคอยด์ ส่วนคาร์วินไซเคิลเกิดที่สโตรมา'],
    ['BIO-301', 'biology', 'hard', 'การแบ่งเซลล์แบบไมโอซิส (Meiosis) ปรากฏการณ์ Crossing over เกิดขึ้นในระยะใดอย่างเจาะจง?', 'โพรเฟส I (Prophase I)', 'เมทาเฟส I (Metaphase I)', 'แอนาเฟส II (Anaphase II)', 'โพรเฟส II (Prophase II)', 'A', 'Crossing over เกิดในระยะ Pachynema ของ Prophase I'],
    ['BIO-302', 'biology', 'hard', 'ฮอร์โมนพืชชนิดใดมีบทบาทเด่นในการกระตุ้นการยืดตัวของลำต้นและทำลายการพักตัวของเมล็ด?', 'จิบเบอเรลลิน (Gibberellin)', 'แอบไซซิก (Abscisic acid)', 'เอทิลีน (Ethylene)', 'ไซโทไคนิน (Cytokinin)', 'A', 'จิบเบอเรลลินช่วยยืดลำต้นและกระตุ้นการงอกของเมล็ด'],

    // --- 🧪 เคมี (Chemistry) ม.ปลาย ---
    ['CHM-101', 'chemistry', 'easy', 'พันธะเคมีที่เกิดจากการใช้อิเล็กตรอนร่วมกันระหว่างอะตอมอโลหะกับอโลหะคือพันธะใด?', 'พันธะโคเวเลนต์ (Covalent bond)', 'พันธะไอออนิก (Ionic bond)', 'พันธะโลหะ (Metallic bond)', 'พันธะไฮโดรเจน (Hydrogen bond)', 'A', 'พันธะโคเวเลนต์เกิดจากการใช้อิเล็กตรอนคู่ร่วมพันธะ'],
    ['CHM-102', 'chemistry', 'easy', 'สารละลายที่มีค่า pH = 2 มีสมบัติความเป็นกรด-เบสอย่างไร?', 'กรดแก่', 'เบสแก่', 'กลาง', 'เกลือไฮโดรไลซิส', 'A', 'pH < 7 คือกรด ยิ่งน้อยยิ่งมีความเป็นกรดสูง'],
    ['CHM-201', 'chemistry', 'medium', 'แก๊สชนิดหนึ่งมีปริมาตร 22.4 ลิตร ที่สภาวะมาตรฐาน (STP) จะมีจำนวนอนุภาคกี่โมล?', '1.0 โมล', '0.5 โมล', '2.0 โมล', '4.0 โมล', 'A', 'ที่ STP แก๊ส 1 โมล มีปริมาตร 22.4 ลิตร'],
    ['CHM-202', 'chemistry', 'medium', 'ธาตุที่มีการจัดเรียงอิเล็กตรอนเป็น 2, 8, 8, 2 อยู่หมู่ใดและคาบใดในตารางธาตุ?', 'หมู่ 2 คาบ 4', 'หมู่ 4 คาบ 2', 'หมู่ 2 คาบ 3', 'หมู่ 8 คาบ 4', 'A', 'เวเลนซ์อิเล็กตรอน = 2 (หมู่ 2) และมี 4 ระดับพลังงาน (คาบ 4)'],
    ['CHM-301', 'chemistry', 'hard', 'ตามทฤษฎีกรด-เบสของเบรินสเตด-ลาวรี (Brønsted–Lowry) สารที่เป็นคู่เบสของ H2SO4 คือสารใด?', 'HSO4-', 'SO4^2-', 'H3SO4+', 'H3O+', 'A', 'คู่เบสคือสารที่มีโปรตอน (H+) น้อยกว่ากรดตัวนั้นอยู่ 1 ตัว'],
    ['CHM-302', 'chemistry', 'hard', 'สารประกอบอินทรีย์ชนิดใดทำปฏิกิริยากับสารละลายโซเดียมไฮโดรเจนคาร์บอเนต (NaHCO3) แล้วเกิดฟองแก๊ส CO2?', 'กรดคาร์บอกซิลิก (Carboxylic acid)', 'แอลกอฮอล์ (Alcohol)', 'เอสเทอร์ (Ester)', 'อีเทอร์ (Ether)', 'A', 'กรดคาร์บอกซิลิกมีความเป็นกรดพอที่จะทำปฏิกิริยากับ NaHCO3 ได้แก๊ส CO2'],

    // --- ⚡ ฟิสิกส์ (Physics) ม.ปลาย ---
    ['PHY-101', 'physics', 'easy', 'ตามกฎการเคลื่อนที่ข้อที่ 1 ของนิวตัน (Newton\'s First Law) วัตถุจะรักษาสภาพการเคลื่อนที่คงที่เมื่อใด?', 'เมื่อแรงลัพธ์ที่กระทำต่อวัตถุเป็นศูนย์', 'เมื่อมีความเร่งคงที่', 'เมื่อมวลของวัตถุเป็นศูนย์', 'เมื่อไม่มีแรงโน้มถ่วง', 'A', 'เมื่อ ซิกมา F = 0 วัตถุจะอยู่นิ่งหรือเคลื่อนที่ด้วยความเร็วคงที่'],
    ['PHY-102', 'physics', 'easy', 'หน่วยในระบบเอสไอ (SI Unit) ของปริมาณ "งาน" และ "พลังงาน" คือหน่วยใด?', 'จูล (Joule : J)', 'นิวตัน (Newton : N)', 'วัตต์ (Watt : W)', 'ปาสกาล (Pascal : Pa)', 'A', 'งานและพลังงานมีหน่วยเป็น จูล (J)'],
    ['PHY-201', 'physics', 'medium', 'วัตถุมวล 2 kg เคลื่อนที่ด้วยความเร็ว 10 m/s จะมีพลังงานจลน์ (Ek) เท่าใด?', '100 จูล', '50 จูล', '200 จูล', '20 จูล', 'A', 'Ek = (1/2)mv^2 = 0.5 * 2 * (10^2) = 100 J'],
    ['PHY-202', 'physics', 'medium', 'เมื่อแสงเดินทางจากตัวกลางที่มีดัชนีหักเหน้อย ไปยังตัวกลางที่มีดัชนีหักเหมาก รังสีของแสงจะหักเหอย่างไร?', 'เบนเข้าหาเส้นแนวฉาก (Normal line)', 'เบนออกจากเส้นแนวฉาก', 'สะท้อนกลับหมด 100%', 'เดินทางเป็นเส้นตรงตามเดิม', 'A', 'จาก n น้อย ไป n มาก รังสีจะหักเหเบนเข้าหาเส้นแนวฉาก'],
    ['PHY-301', 'physics', 'hard', 'ตามกฎของคูลอมบ์ หากระยะห่างระหว่างจุดประจุไฟฟ้า 2 จุด เพิ่มขึ้นเป็น 3 เท่าของเดิม แรงทางไฟฟ้าจะเปลี่ยนแปลงอย่างไร?', 'ลดลงเหลือ 1/9 เท่าของเดิม', 'ลดลงเหลือ 1/3 เท่าของเดิม', 'เพิ่มขึ้น 3 เท่า', 'เพิ่มขึ้น 9 เท่า', 'A', 'แรงแปรผกผันกับระยะทางยกกำลังสอง (F แปรผกผันกับ r^2) เมื่อ r เพิ่ม 3 เท่า F จะเป็น 1/9'],
    ['PHY-302', 'physics', 'hard', 'คลื่นเสียงความถี่ 680 Hz เคลื่อนที่ในอากาศที่มีอัตราเร็วเสียง 340 m/s คลื่นนี้จะมีความยาวคลื่นกี่เมตร?', '0.5 เมตร', '2.0 เมตร', '0.25 เมตร', '1.5 เมตร', 'A', 'v = f * lambda -> lambda = v / f = 340 / 680 = 0.5 m']
  ];

  sample.forEach(row => sheet.appendRow(row));
}
