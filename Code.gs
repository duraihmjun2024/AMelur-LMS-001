/* ==========================================================
   Google Apps Script Backend for AMelur-LMS-001
   Author: Durai | PUPS Anjalam Melur
   Supports: M05 (Levels 4.1-4.8, 3 Trials/Level), Student & Member IDs,
             Daily Completion Tracking & Dashboard Feeds
   ========================================================== */

var IST_TIMEZONE = "Asia/Kolkata";

function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return jsonResponse({ "status": "error", "message": "Server busy." });
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var params = {};

    if (e && e.parameter) {
      for (var key in e.parameter) {
        params[key] = e.parameter[key];
      }
    }

    if (e && e.postData && e.postData.contents) {
      try {
        var jsonBody = JSON.parse(e.postData.contents);
        for (var k in jsonBody) {
          params[k] = jsonBody[k];
        }
      } catch (jsonErr) {}
    }

    var action = String(params.action || "").trim();

    if (action === "getDashboard") {
      return handleGetDashboard(ss);
    }

    if (action === "getAllData" || action === "getLatestSession") {
      return handleGetAllData(ss);
    }

    if (action === "saveUserData") {
      return handleSaveUserData(ss, params);
    }

    if (action === "saveSessionSummary") {
      return handleSaveSessionSummary(ss, params);
    }

    if (action === "submitScore") {
      return handleSubmitScore(ss, params);
    }

    // Default fallback: treat as game score submission
    return handleSubmitScore(ss, params);

  } catch (error) {
    return jsonResponse({
      "status": "error",
      "message": error.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getISTDateKey(dateObj) {
  var d = dateObj || new Date();
  return Utilities.formatDate(d, IST_TIMEZONE, "yyyy-MM-dd");
}

function getFormattedTimestamp() {
  var nowObj = new Date();
  return Utilities.formatDate(nowObj, IST_TIMEZONE, "ddMMMMyy-HH:mm:ss").replace(/([A-Za-z]{3})[A-Za-z]+/, "$1");
}

function cleanTimeText(val) {
  if (!val || val === "--" || val === "---") return "00:00:00";
  var str = String(val).trim();
  if (str.indexOf("T") !== -1) {
    var parts = str.split("T");
    if (parts[1]) return parts[1].split(".")[0];
  }
  return str;
}

function parseSecondsFromVal(val) {
  if (typeof val === "number") return Math.max(0, Math.round(val));
  var str = String(val || "").trim();
  if (!str) return 0;
  if (/^\d+$/.test(str)) return parseInt(str, 10);
  if (/^\d+s$/i.test(str)) return parseInt(str, 10);
  if (str.indexOf(":") !== -1) {
    var p = str.split(":").map(function(n) { return parseInt(n, 10) || 0; });
    if (p.length === 3) return p[0] * 3600 + p[1] * 60 + p[2];
    if (p.length === 2) return p[0] * 60 + p[1];
  }
  return 0;
}

function toDisplayReportId(fullOrRawId, whatsappVal) {
  var idStr = String(fullOrRawId || "").trim().toUpperCase();
  var waDigits = String(whatsappVal || "").replace(/\D/g, "");
  if (idStr.indexOf("M") === 0 && idStr.length > 4) {
    if (waDigits.length >= 3) return "M" + waDigits.slice(-3);
    return "M" + idStr.slice(-3);
  }
  if (idStr.indexOf("M") === 0 && idStr.length <= 4) {
    return idStr;
  }
  return idStr.slice(0, 4) || "STU0";
}

function inferRoleAndStd(fullOrRawId, stdParam) {
  var s = String(stdParam || "").trim().toUpperCase();
  var idStr = String(fullOrRawId || "").trim().toUpperCase();
  if (s === "M" || idStr.indexOf("M") === 0) {
    return { role: "member", std: "M" };
  }
  var stdMatch = s.match(/^[1-5]$/) || idStr.match(/^([1-5])[GB]/);
  return { role: "student", std: stdMatch ? stdMatch[1] : (s || "5") };
}






function handleGetDashboard(ss) {
  var records = [];
  var todayKey = getISTDateKey();
  var gameSheet = ss.getSheetByName("GameData");

  if (gameSheet && gameSheet.getLastRow() > 1) {
    var rows = gameSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[1] && !r[2]) continue;

      var rawDate = String(r[0] || "");
      var dateMatch = rawDate.match(/(\d{4}-\d{2}-\d{2})/);
      var recDate = dateMatch ? dateMatch[1] : todayKey;

      var fullId = String(r[2] || "").trim();
      var whatsappField = String(r[11] || "");
      var dispId = toDisplayReportId(fullId, whatsappField);
      var meta = inferRoleAndStd(fullId, r[10]);

      var scoreVal = parseInt(r[8], 10);
      if (isNaN(scoreVal)) {
        var mScore = String(r[8] || "0").match(/^(\d+)/);
        scoreVal = mScore ? parseInt(mScore[1], 10) : 0;
      }

      var maxScoreVal = parseInt(r[9], 10);
      if (isNaN(maxScoreVal)) {
        var mMax = String(r[8] || "").match(/\/(\d+)/);
        maxScoreVal = mMax ? parseInt(mMax[1], 10) : 4;
      }

      var secVal = parseSecondsFromVal(r[7] || r[5]);

      records.push({
        date: recDate,
        gameCode: String(r[1] || "M05").trim(),
        levelTag: String(r[4] || "").trim(),
        fullId: fullId,
        displayId: dispId,
        name: String(r[3] || "Student").trim(),
        std: meta.std,
        role: meta.role,
        score: scoreVal,
        maxScore: maxScoreVal,
        timeSec: secVal
      });
    }
  }

  return jsonResponse({
    "status": "success",
    "date": todayKey,
    "records": records
  });
}

function handleGetAllData(ss) {
  var gameSheet = ss.getSheetByName("GameData");
  var userSheet = ss.getSheetByName("UserData") || ss.getSheetByName("Users");

  var gameData = [];
  if (gameSheet && gameSheet.getLastRow() > 1) {
    var gRows = gameSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < gRows.length; i++) {
      var row = gRows[i];
      if (row.length >= 9) {
        gameData.push({
          submittedAt: row[0] || "",
          gameCode: row[1] || "",
          studentId: toDisplayReportId(row[2] || "", row[11] || ""),
          name: row[3] || "",
          level1Score: row[4] || "",
          level1Time: cleanTimeText(row[5]),
          level2Score: row[6] || "",
          level2Time: cleanTimeText(row[7]),
          score: row[8] || "0",
          trialId: row[9] || "1"
        });
      }
    }
  }

  var userData = [];
  if (userSheet && userSheet.getLastRow() > 1) {
    var uRows = userSheet.getDataRange().getDisplayValues();
    for (var j = 1; j < uRows.length; j++) {
      var uRow = uRows[j];
      if (uRow[1]) {
        userData.push({
          submittedAt: uRow[0] || "",
          studentId: toDisplayReportId(uRow[1] || "", uRow[7] || ""),
          emisNumber: uRow[2] || "",
          dob: uRow[3] || "",
          name: uRow[4] || "",
          fatherName: uRow[5] || "",
          cell: uRow[6] || "",
          whatsapp: uRow[7] || "",
          school: uRow[8] || "",
          place: uRow[9] || "",
          country: uRow[10] || ""
        });
      }
    }
  }

  var lastGameRecord = gameData.length > 0 ? gameData[gameData.length - 1] : null;
  var lastStudentId = lastGameRecord ? lastGameRecord.studentId : (userData.length > 0 ? userData[userData.length - 1].studentId : "1B01");
  var lastGameCode = lastGameRecord ? lastGameRecord.gameCode : "M05";

  return jsonResponse({
    "status": "success",
    "gameData": gameData,
    "userData": userData,
    "latestStudentId": lastStudentId,
    "latestGameCode": lastGameCode
  });
}

function handleSaveUserData(ss, params) {
  var userSheet = ss.getSheetByName("UserData") || ss.insertSheet("UserData");
  if (userSheet.getLastRow() === 0) {
    userSheet.appendRow([
      "Submitted At",
      "Student/Mem_ID",
      "EMIS Number",
      "DOB",
      "Name/First Name",
      "Father Name/Last Name",
      "Cell",
      "Whatsapp",
      "Institution",
      "Place",
      "Country"
    ]);
  }

  userSheet.appendRow([
    getFormattedTimestamp(),
    params.studentId || params.memberId || "",
    params.emisNumber || "",
    params.dob || "",
    params.name || "",
    params.fatherName || "",
    params.cell || "",
    params.whatsapp || "",
    params.school || "PUPS Anjalam Melur",
    params.place || "",
    params.country || "India"
  ]);

  return jsonResponse({ "status": "success", "message": "User saved" });
}

function handleSaveSessionSummary(ss, params) {
  var summarySheet = ss.getSheetByName("SessionSummary") || ss.insertSheet("SessionSummary");
  if (summarySheet.getLastRow() === 0) {
    summarySheet.appendRow([
      "Submitted At",
      "Game Code",
      "Student ID",
      "Name",
      "Trial 1 Score",
      "Trial 1 Time",
      "Trial 2 Score",
      "Trial 2 Time",
      "Trial 3 Score",
      "Trial 3 Time",
      "Cumulative Total Score",
      "Cumulative Total Time"
    ]);
  }

  summarySheet.appendRow([
    getFormattedTimestamp(),
    params.gameCode || "",
    params.studentId || "",
    params.name || "",
    params.t1Score || "0",
    params.t1Time || "00:00:00",
    params.t2Score || "0",
    params.t2Time || "00:00:00",
    params.t3Score || "0",
    params.t3Time || "00:00:00",
    params.totalScore || "0",
    params.totalTime || "00:00:00"
  ]);

  return jsonResponse({ "status": "success", "message": "Session summary saved" });
}

function handleSubmitScore(ss, params) {
  var gameSheet = ss.getSheetByName("GameData") || ss.insertSheet("GameData");
  if (gameSheet.getLastRow() === 0) {
    gameSheet.appendRow([
      "Submitted At",
      "Game Code",
      "Student ID",
      "Name",
      "Level Tag",
      "Trial Number",
      "Score",
      "Time Seconds",
      "Best Score",
      "Max Score",
      "Standard / Role",
      "Whatsapp"
    ]);
  }

  var gameCode = String(params.gameCode || params.gameName || "M05").trim();
  var studentID = String(params.studentID || params.studentId || params.memberId || "").trim();
  var name = String(params.name || "").trim();
  var levelTag = String(params.levelTag || params.level1Score || "").trim();
  var trialNumber = String(params.trialNumber || params.trialId || "1").trim();
  var scoreVal = (params.score !== undefined && params.score !== null && params.score !== "")
    ? params.score
    : "0";
  var secVal = parseSecondsFromVal(params.timeSeconds || params.timeFormatted || params.level1Time || 0);
  var bestScoreVal = (params.totalBestScore !== undefined && params.totalBestScore !== null && params.totalBestScore !== "")
    ? params.totalBestScore
    : scoreVal;
  var maxScoreVal = (params.maxScore !== undefined && params.maxScore !== null && params.maxScore !== "")
    ? params.maxScore
    : 4;
  var meta = inferRoleAndStd(studentID, params.std);
  var whatsapp = String(params.whatsapp || "").trim();

  gameSheet.appendRow([
    getFormattedTimestamp(),
    gameCode,
    studentID,
    name,
    levelTag,
    trialNumber,
    scoreVal,
    secVal,
    bestScoreVal,
    maxScoreVal,
    meta.std,
    whatsapp
  ]);

  return jsonResponse({
    "status": "success",
    "gameCode": gameCode,
    "displayId": toDisplayReportId(studentID, whatsapp),
    "levelTag": levelTag,
    "trialNumber": trialNumber,
    "score": scoreVal
  });
}