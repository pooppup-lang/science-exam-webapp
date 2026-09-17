/**
 * ระบบทำข้อสอบวิทยาศาสตร์ (Science Examination System)
 * Google Apps Script Backend (Code.gs)
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
    .setTitle('ระบบทำข้อสอบวิทยาศาสตร์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
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
    teacherSheet.appendRow(['TeacherID', 'FullName', 'Username', 'PasswordHash', 'CreatedAt']);
  }

  // 2. Rooms
  const roomSheet = getOrCreateSheet(SHEET_NAMES.ROOMS);
  if (roomSheet.getLastRow() === 0) {
    roomSheet.appendRow(['RoomID', 'TeacherID', 'RoomCode', 'RoomName', 'Classroom', 'EasyCount', 'MediumCount', 'HardCount', 'TimeLimit', 'IsActive', 'CreatedAt']);
  }

  // 3. Questions (Question Bank)
  const questionSheet = getOrCreateSheet(SHEET_NAMES.QUESTIONS);
  if (questionSheet.getLastRow() === 0) {
    questionSheet.appendRow(['QuestionID', 'Difficulty', 'Question', 'ChoiceA', 'ChoiceB', 'ChoiceC', 'ChoiceD', 'CorrectAnswer', 'Explanation']);
    seedDefaultQuestions(questionSheet);
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

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
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
// Teacher Authentication & Management
// -------------------------------------------------------------

function registerTeacher(fullName, username, password) {
  fullName = normalizeText(fullName);
  username = normalizeText(username).toLowerCase();
  password = normalizeText(password);

  if (!fullName || !username || !password) {
    throw new Error('กรุณากรอกข้อมูลให้ครบทุกช่อง');
  }
  if (password.length < 6) {
    throw new Error('รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร');
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
  sheet.appendRow([teacherId, fullName, username, passHash, new Date()]);

  return 'สมัครสมาชิกสำเร็จ! กรุณาเข้าสู่ระบบ';
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

function changeTeacherPassword(token, oldPassword, newPassword) {
  const session = requireTeacher(token);
  const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
  const data = sheet.getDataRange().getValues();
  const oldHash = hashPassword(oldPassword);
  const newHash = hashPassword(newPassword);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === session.userId) {
      if (data[i][3] !== oldHash) {
        throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
      }
      sheet.getRange(i + 1, 4).setValue(newHash);
      return 'เปลี่ยนรหัสผ่านเรียบร้อยแล้ว';
    }
  }
  throw new Error('ไม่พบข้อมูลครู');
}

function resetTeacherPassword(username, newPassword) {
  const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
  const data = sheet.getDataRange().getValues();
  const newHash = hashPassword(newPassword);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === username.toLowerCase()) {
      sheet.getRange(i + 1, 4).setValue(newHash);
      return 'รีเซ็ตรหัสผ่านสำเร็จ';
    }
  }
  throw new Error('ไม่พบชื่อผู้ใช้นี้');
}

// -------------------------------------------------------------
// Exam Rooms
// -------------------------------------------------------------

function generateRoomCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = 'SCI-';
  for (let i = 0; i < 6; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

function createRoom(token, name, classroom, easy, medium, hard, timeLimit) {
  const session = requireTeacher(token);

  name = normalizeText(name);
  classroom = normalizeText(classroom);
  easy = parseInt(easy, 10) || 0;
  medium = parseInt(medium, 10) || 0;
  hard = parseInt(hard, 10) || 0;
  timeLimit = parseInt(timeLimit, 10);
  if (isNaN(timeLimit) || timeLimit < 0) timeLimit = 15; // default 15 mins

  if (!name || !classroom) {
    throw new Error('กรุณาระบุชื่อห้องสอบและชั้น/ห้อง');
  }
  if (easy + medium + hard <= 0) {
    throw new Error('กรุณาระบุจำนวนข้อสอบอย่างน้อย 1 ข้อ');
  }

  const roomId = 'R-' + Utilities.getUuid().slice(0, 8);
  const code = generateRoomCode();
  const roomSheet = getOrCreateSheet(SHEET_NAMES.ROOMS);

  roomSheet.appendRow([
    roomId,
    session.userId,
    code,
    name,
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
    classroom: classroom,
    easy: easy,
    medium: medium,
    hard: hard,
    timeLimit: timeLimit
  };
}

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
    const classroom = row[4];
    const easy = row[5];
    const medium = row[6];
    const hard = row[7];
    // Column 8 is TimeLimit if 11 columns, or IsActive if older 10-column version
    let timeLimit = 15;
    let isActive = true;

    if (row.length >= 11) {
      timeLimit = Number(row[8]) || 0;
      isActive = row[9] === true;
    } else {
      isActive = row[8] === true;
    }

    if (teacherId === session.userId && isActive === true) {
      rooms.push({
        id: id,
        code: code,
        name: name,
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

function closeRoom(token, roomId) {
  const session = requireTeacher(token);
  const sheet = getOrCreateSheet(SHEET_NAMES.ROOMS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === roomId && data[i][1] === session.userId) {
      // isActive is column 10 (index 9) if 11 cols, or column 9 (index 8) if 10 cols
      const activeColIndex = data[i].length >= 11 ? 10 : 9;
      sheet.getRange(i + 1, activeColIndex).setValue(false);
      return 'ปิดห้องสอบเรียบร้อยแล้ว';
    }
  }
  throw new Error('ไม่พบห้องสอบ หรือไม่มีสิทธิ์ปิดห้องนี้');
}

function getRoomResults(token, roomId) {
  requireTeacher(token);
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
    const switchCount = row.length >= 10 ? (Number(row[8]) || 0) : 0;
    const timestamp = row.length >= 10 ? row[9] : row[8];

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

function resetStudentResult(token, resultId) {
  requireTeacher(token);
  const sheet = getOrCreateSheet(SHEET_NAMES.RESULTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(resultId)) {
      sheet.deleteRow(i + 1);
      return 'ลบผลสอบเรียบร้อยแล้ว นักเรียนสามารถเข้าสอบใหม่ได้';
    }
  }
  throw new Error('ไม่พบข้อมูลผลสอบนี้');
}

// -------------------------------------------------------------
// Student Flow
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
    const classroom = row[4];
    const easy = row[5];
    const medium = row[6];
    const hard = row[7];
    let timeLimit = 15;
    let isActive = true;

    if (row.length >= 11) {
      timeLimit = Number(row[8]) || 0;
      isActive = row[9] === true;
    } else {
      isActive = row[8] === true;
    }

    if (String(roomCode).toUpperCase() === code && isActive === true) {
      return {
        id: id,
        code: roomCode,
        name: name,
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

/**
 * ตรวจสอบว่านักเรียนคนนี้เคยส่งข้อสอบในห้องนี้ไปแล้วหรือไม่
 */
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

  // ตรวจสอบการจำกัดสิทธิ์ 1 คนต่อ 1 ครั้ง
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

function generateStudentExam(studentToken) {
  const session = requireStudent(studentToken);
  const room = findRoomByCode(session.roomCode);
  if (!room) {
    throw new Error('ห้องสอบไม่พร้อมใช้งาน');
  }

  const allQuestions = getQuestionBank();
  const easyPool = allQuestions.filter(q => q.difficulty === 'easy');
  const mediumPool = allQuestions.filter(q => q.difficulty === 'medium');
  const hardPool = allQuestions.filter(q => q.difficulty === 'hard');

  const selectedQuestions = [
    ...shuffle(easyPool).slice(0, room.easy),
    ...shuffle(mediumPool).slice(0, room.medium),
    ...shuffle(hardPool).slice(0, room.hard)
  ];

  const shuffledExam = shuffle(selectedQuestions);

  // ตัดเฉลยออกก่อนส่งไปฝั่งหน้าบ้าน
  const clientQuestions = shuffledExam.map(q => ({
    id: q.id,
    question: q.question,
    choices: [q.choiceA, q.choiceB, q.choiceC, q.choiceD]
  }));

  return {
    room: {
      name: room.name,
      classroom: room.classroom,
      code: room.code,
      timeLimit: room.timeLimit || 15
    },
    questions: clientQuestions
  };
}

function submitStudentExam(studentToken, answers, switchCount) {
  const session = requireStudent(studentToken);
  const room = findRoomByCode(session.roomCode);
  if (!room) {
    throw new Error('ห้องสอบไม่พร้อมใช้งาน');
  }

  // ป้องกันการส่งซ้ำระดับ Server
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
      const isCorrect = String(ans.selected).toUpperCase() === String(q.correctAnswer).toUpperCase();
      if (isCorrect) correctCount++;
      results.push({
        id: ans.id,
        isCorrect: isCorrect,
        correctAnswer: q.correctAnswer
      });
    }
  });

  const total = answers.length;
  const percentage = total > 0 ? Math.round((correctCount / total) * 100) : 0;

  // บันทึกคะแนนลงชีต Results
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

function isActive(item) {
  return item && item.isActive === true;
}

function publicRoom(room) {
  if (!room) return null;
  return {
    code: room.code,
    name: room.name,
    classroom: room.classroom,
    easy: room.easy,
    medium: room.medium,
    hard: room.hard,
    timeLimit: room.timeLimit
  };
}

// -------------------------------------------------------------
// Question Bank
// -------------------------------------------------------------

function getQuestionBank() {
  const sheet = getOrCreateSheet(SHEET_NAMES.QUESTIONS);
  const data = sheet.getDataRange().getValues();
  const questions = [];

  for (let i = 1; i < data.length; i++) {
    const [id, difficulty, qText, a, b, c, d, correct, exp] = data[i];
    if (qText) {
      questions.push({
        id: String(id),
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

function seedDefaultQuestions(sheet) {
  const sample = [
    // Easy
    ['Q-101', 'easy', 'หน่วยพื้นฐานที่เล็กที่สุดของสิ่งมีชีวิตคืออะไร?', 'เซลล์ (Cell)', 'เนื้อเยื่อ (Tissue)', 'อวัยวะ (Organ)', 'ระบบร่างกาย (System)', 'A', 'เซลล์คือหน่วยพื้นฐานของสิ่งมีชีวิต'],
    ['Q-102', 'easy', 'สารใดเป็นตัวทำละลายหลักในร่างกายของมนุษย์?', 'น้ำ', 'ไขมัน', 'โปรตีน', 'น้ำตาล', 'A', 'น้ำคิดเป็นสัดส่วนประมาณ 60-70% ของร่างกาย'],
    ['Q-103', 'easy', 'สิ่งมีชีวิตกลุ่มใดทำหน้าที่เป็นผู้ผลิตในระบบนิเวศ?', 'พืชสีเขียว', 'สัตว์กินพืช', 'แบคทีเรียย่อยสลาย', 'สัตว์กินเนื้อ', 'A', 'พืชสามารถสังเคราะห์ด้วยแสงได้'],
    ['Q-104', 'easy', 'อวัยวะใดทำหน้าที่กรองของเสียออกจากเลือดในระบบขับถ่าย?', 'ไต', 'ปอด', 'หัวใจ', 'กระเพาะอาหาร', 'A', 'ไตทำหน้าที่กรองของเสีย'],
    ['Q-105', 'easy', 'สัญลักษณ์ทางเคมีของแก๊สออกซิเจนคืออะไร?', 'O2', 'CO2', 'H2O', 'N2', 'A', 'แก๊สออกซิเจนโมเลกุลคือ O2'],
    
    // Medium
    ['Q-201', 'medium', 'กระบวนการสังเคราะห์ด้วยแสงของพืชเกิดขึ้นที่ออร์แกเนลล์ใดเป็นหลัก?', 'คลอโรพลาสต์ (Chloroplast)', 'ไมโทคอนเดรีย (Mitochondria)', 'ไรโบโซม (Ribosome)', 'กอลจิบอดี (Golgi Body)', 'A', 'คลอโรพลาสต์มีรงควัตถุคลอโรฟิลล์ดักจับแสง'],
    ['Q-202', 'medium', 'แรงเสียดทานมีทิศทางอย่างไรกับการเคลื่อนที่ของวัตถุ?', 'ตรงข้ามกับทิศทางการเคลื่อนที่', 'ทิศทางเดียวกับการเคลื่อนที่', 'ตั้งฉากกับการเคลื่อนที่', 'ไม่มีทิศทางที่แน่นอน', 'A', 'แรงเสียดทานมีทิศต้านการเคลื่อนที่เสมอ'],
    ['Q-203', 'medium', 'สารละลายที่มีค่า pH เท่ากับ 3 มีสมบัติเป็นอย่างไร?', 'กรด', 'เบส', 'กลาง', 'เกลือ', 'A', 'pH น้อยกว่า 7 มีสมบัติเป็นกรด'],
    ['Q-204', 'medium', 'การถ่ายโอนความร้อนที่ไม่ต้องอาศัยตัวกลางคือข้อใด?', 'การแผ่รังสีความร้อน (Radiation)', 'การนำความร้อน (Conduction)', 'การพาความร้อน (Convection)', 'การระเหย (Evaporation)', 'A', 'การแผ่รังสีความร้อนเดินทางผ่านสุญญากาศได้'],

    // Hard
    ['Q-301', 'hard', 'ตามกฎการเคลื่อนที่ข้อที่ 2 ของนิวตัน (F = ma) หากมวลคงที่และแรงลัพธ์เพิ่มขึ้นเป็น 2 เท่า ความเร่งจะเป็นอย่างไร?', 'เพิ่มขึ้นเป็น 2 เท่า', 'ลดลงครึ่งหนึ่ง', 'คงที่', 'เพิ่มขึ้นเป็น 4 เท่า', 'A', 'ความเร่งแปรผันตรงกับแรงลัพธ์เมื่อมวลคงที่'],
    ['Q-302', 'hard', 'การแบ่งเซลล์แบบไมโอซิส (Meiosis) มีความสำคัญอย่างไรต่อสิ่งมีชีวิต?', 'สร้างเซลล์สืบพันธุ์และรักษาจำนวนโครโมโซมให้คงที่', 'เจริญเติบโตร่างกายและซ่อมแซมส่วนที่สึกหรอ', 'เพิ่มจำนวนเซลล์ร่างกายให้มีโครโมโซมเท่าเดิม', 'สร้างเซลล์ผิวหนังใหม่', 'A', 'ไมโอซิสลดจำนวนโครโมโซมลงครึ่งหนึ่งเพื่อเตรียมสร้างเซลล์สืบพันธุ์'],
    ['Q-303', 'hard', 'ธาตุที่มีเลขอะตอม 12 มีการจัดเรียงอิเล็กตรอนในระดับพลังงานหลักอย่างไร?', '2, 8, 2', '2, 8, 8', '2, 6, 4', '2, 8, 1', 'A', 'แมกนีเซียม (Mg) เลขอะตอม 12 จัดเรียง 2, 8, 2 อยู่หมู่ 2 คาบ 3']
  ];

  sample.forEach(row => sheet.appendRow(row));
}
