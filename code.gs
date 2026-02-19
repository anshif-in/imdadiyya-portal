var SHEET_ID = "1zYODH88AdkEKLmUoqHa8RzIVRjBdLvr247jo0voVDH4";
var ADMIN_PASSWORD = "admin123";

/* =========================
   WEB ROUTING
========================= */

function doGet(e) {

  if (!e || !e.parameter || Object.keys(e.parameter).length === 0)
    return render("login");

  if (e.parameter.page == "marks") return render("mark");
  if (e.parameter.page == "report") return render("report");
  if (e.parameter.page == "admin") return render("admin");

  if (e.parameter.roll) return getReport(e.parameter.roll);
  if (e.parameter.class) return getSubjects(e.parameter.class);
  if (e.parameter.adminClass) return getClassStudents(e.parameter.adminClass, e.parameter.division);
  if (e.parameter.searchName) return searchByName(e.parameter.searchName);
  if (e.parameter.exportClass) return exportClass(e.parameter.exportClass, e.parameter.division);

  return render("login");
}

function render(file) {
  return HtmlService.createHtmlOutputFromFile(file)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================
   POST REQUEST HANDLER
========================= */

function doPost(e) {

  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.openById(SHEET_ID);

  if (data.type == "adminLogin")
    return json({ success: data.password === ADMIN_PASSWORD });

  if (data.type == "login")
    return studentLogin(ss, data);

  if (data.type == "marks")
    return saveMarks(ss, data);

  if (data.type == "update")
    return updateMarks(ss, data);

  if (data.type == "delete")
    return deleteMarks(ss, data.roll);

  if (data.type == "bulkUpload")
    return bulkUpload(ss, data.records);

  return json({ error: "Invalid Request" });
}

/* =========================
   STUDENT LOGIN
========================= */

function studentLogin(ss, data) {

  var sh = ss.getSheetByName("Students");
  if (!sh) return json({ error: "Students sheet missing" });

  var rows = sh.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.roll &&
      rows[i][2] == data.class &&
      rows[i][3] == data.division &&
      rows[i][4] == data.password)
      return json({ success: true, roll: data.roll });
  }

  return json({ success: false });
}

/* =========================
   SAVE MARKS
========================= */

function saveMarks(ss, data) {

  var sh = ss.getSheetByName("Marks");
  if (!sh) return json({ error: "Marks sheet missing" });

  var rows = sh.getDataRange().getValues();

  // Prevent duplicate roll
  for (var i = 1; i < rows.length; i++)
    if (rows[i][2] == data.roll)
      return json({ error: "Marks already exist for this roll" });

  var subjects = getSubjectsArray(ss, data.class);
  var total = 0;
  var row = [data.class, data.division, data.roll, data.name];

  subjects.forEach(function (s) {
    var m = Number(data.subjects[s] || 0);
    row.push(m);
    total += m;
  });

  row.push(total);
  sh.appendRow(row);

  return json({ success: true });
}

/* =========================
   UPDATE MARKS
========================= */

function updateMarks(ss, data) {

  var sh = ss.getSheetByName("Marks");
  var rows = sh.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {

    if (rows[i][2] == data.roll) {

      var subjects = getSubjectsArray(ss, rows[i][0]);
      var total = 0;

      subjects.forEach(function (s, index) {
        var m = Number(data.subjects[s] || 0);
        sh.getRange(i + 1, 5 + index).setValue(m);
        total += m;
      });

      sh.getRange(i + 1, 5 + subjects.length).setValue(total);

      return json({ success: true });
    }
  }

  return json({ error: "Roll not found" });
}

/* =========================
   DELETE MARKS
========================= */

function deleteMarks(ss, roll) {

  var sh = ss.getSheetByName("Marks");
  var rows = sh.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++)
    if (rows[i][2] == roll) {
      sh.deleteRow(i + 1);
      return json({ success: true });
    }

  return json({ error: "Roll not found" });
}

/* =========================
   BULK UPLOAD
========================= */

function bulkUpload(ss, records) {

  var sh = ss.getSheetByName("Marks");
  var existing = sh.getDataRange().getValues();

  var existingRolls = existing.slice(1).map(r => r[2]);

  records.forEach(function (r) {

    if (existingRolls.includes(r.roll)) return; // Skip duplicate

    var subjects = getSubjectsArray(ss, r.class);
    var total = 0;
    var row = [r.class, r.division, r.roll, r.name];

    subjects.forEach(function (s) {
      var m = Number(r[s] || 0);
      row.push(m);
      total += m;
    });

    row.push(total);
    sh.appendRow(row);
  });

  return json({ success: true });
}

/* =========================
   SEARCH
========================= */

function searchByName(name) {

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName("Marks");
  var rows = sh.getDataRange().getValues();

  var list = [];

  for (var i = 1; i < rows.length; i++) {

    if (rows[i][3] &&
      rows[i][3].toString().toLowerCase().includes(name.toLowerCase()))

      list.push({
        class: rows[i][0],
        division: rows[i][1],
        roll: rows[i][2],
        name: rows[i][3],
        total: rows[i][rows[i].length - 1]
      });
  }

  return json({ students: list });
}

/* =========================
   EXPORT CLASS
========================= */

function exportClass(cls, div) {

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName("Marks");
  var rows = sh.getDataRange().getValues();

  var result = rows.filter(r => r[0] == cls && r[1] == div);

  return json({ data: result });
}

/* =========================
   SUBJECTS
========================= */

function getSubjects(className) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return json({ subjects: getSubjectsArray(ss, className) });
}

function getSubjectsArray(ss, className) {

  var sh = ss.getSheetByName("Subjects");
  if (!sh) return [];

  var rows = sh.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++)
    if (rows[i][0] == className)
      return rows[i].slice(1).filter(String);

  return [];
}

/* =========================
   REPORT + RANK
========================= */

function getReport(roll) {

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName("Marks");
  var rows = sh.getDataRange().getValues();

  var student = rows.find(r => r[2] == roll);
  if (!student) return json({ error: "Not Found" });

  var cls = student[0];
  var div = student[1];

  var subjectsList = getSubjectsArray(ss, cls);
  var subjects = {};

  subjectsList.forEach(function(s, index){
    subjects[s] = student[4 + index];
  });

  var classStudents = rows.slice(1)
    .filter(r => r[0] == cls && r[1] == div)
    .sort((a, b) => b[b.length - 1] - a[a.length - 1]);

  var rank = classStudents.findIndex(r => r[2] == roll) + 1;
  var fail = false;

for (var key in subjects) {
  if (Number(subjects[key]) < 35) {
    fail = true;
    break;
  }
}

var status = fail ? "FAILED" : "PASSED";


  return json({
  roll: student[2],
  name: student[3],
  class: cls,
  division: div,
  subjects: subjects,
  total: student[student.length - 1],
  rank: rank,
  status: status
});

}


/* =========================
   JSON HELPER
========================= */

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}
