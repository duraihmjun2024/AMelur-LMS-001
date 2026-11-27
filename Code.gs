/* ==========================================================
   Google Apps Script Backend for Quiz/Game Tracking (EMIS + DOB + 3-Trial Support)
   Author: Durai | PUPS Anjalam Melur
   ========================================================== */

function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Server busy." }))
                         .setMimeType(ContentService.MimeType.JSON);
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

    var action = params.action || "";

    if (action === "getAllData" || action === "getLatestSession") {
      var gameSheet = ss.getSheetByName("GameData");
      var userSheet = ss.getSheetByName("UserData") || ss.getSheetByName("Users");
      
      var gameData = [];
      if (gameSheet && gameSheet.getLastRow() > 1) {
        var gRows = gameSheet.getDataRange().getDisplayValues();
        for (var i = 1; i < gRows.length; i++) {
          var row = gRows[i];
          if (row.length >= 9) {
            function cleanTimeText(val) {
              if (!val || val === "--" || val === "---") return "00:00:00";
              var str = String(val).trim();
              if (str.includes('T')) {
                var parts = str.split('T');
                if (parts[1]) return parts[1].split('.')[0];
              }
              return str;
            }

            gameData.push({
              submittedAt: row[0] || "",
              gameCode: row[1] || "",
              studentId: row[2] || "",
              name: row[3] || "",
              level1Score: row[4] || "0/10-0%",
              level1Time: cleanTimeText(row[5]),
              level2Score: row[6] || "0/10-0%",
              level2Time: cleanTimeText(row[7]),
              score: row[8] || "0/20-0%",
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
              studentId: uRow[1] || "",
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
      var lastGameCode = lastGameRecord ? lastGameRecord.gameCode : "E02";

      return ContentService.createTextOutput(JSON.stringify({
        "status": "success",
        "gameData": gameData,
        "userData": userData,
        "latestStudentId": lastStudentId,
        "latestGameCode": lastGameCode
      })).setMimeType(ContentService.MimeType.JSON);
    }

    function getFormattedTimestamp() {
      var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      var nowObj = new Date();
      var dd = String(nowObj.getDate()).padStart(2, '0');
      var mmm = months[nowObj.getMonth()];
      var yy = String(nowObj.getFullYear()).slice(-2);
      var hh = String(nowObj.getHours()).padStart(2, '0');
      var mm = String(nowObj.getMinutes()).padStart(2, '0');
      var ssSec = String(nowObj.getSeconds()).padStart(2, '0');
      return dd + mmm + yy + "-" + hh + ":" + mm + ":" + ssSec;
    }

    if (action === "saveUserData") {
      var userSheet = ss.getSheetByName("UserData") || ss.insertSheet("UserData");
      if (userSheet.getLastRow() === 0) {
        userSheet.appendRow([
          "Submitted At", 
          "Student/Mem_ID", 
          "EMIS Number (10 Digits)",
          "DOB",
          "Name/First Name", 
          "Father Name/Last Name", 
          "Cell", 
          "Whatsapp with country code", 
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
      return ContentService.createTextOutput(JSON.stringify({ "status": "success", "message": "User saved" })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "saveSessionSummary") {
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
        params.t1Score || "0/0-0%",
        params.t1Time || "00:00:00",
        params.t2Score || "0/0-0%",
        params.t2Time || "00:00:00",
        params.t3Score || "0/0-0%",
        params.t3Time || "00:00:00",
        params.totalScore || "0/0-0%",
        params.totalTime || "00:00:00"
      ]);

      return ContentService.createTextOutput(JSON.stringify({ "status": "success", "message": "Session summary saved" })).setMimeType(ContentService.MimeType.JSON);
    }

    var gameSheet = ss.getSheetByName("GameData") || ss.insertSheet("GameData");
    if (gameSheet.getLastRow() === 0) {
      gameSheet.appendRow(["Submitted At", "Game Code", "Student ID", "Name", "Level 1 Score", "Level 1 Time", "Level 2 Score", "Level 2 Time", "Total Score", "Trial ID"]);
    }

    var gameCode = params.gameCode || params.gameName || "E02";
    var studentID = params.studentID || params.studentId || params.memberId || "";
    var name = params.name || "";
    var level1Score = params.level1Score || "0/10-0%";
    var level1Time = params.level1Time || "00:00:00";
    var level2Score = params.level2Score || "0/10-0%";
    var level2Time = params.level2Time || "00:00:00";
    var score = params.score || "0/20-0%";
    var trialId = params.trialId || "1";

    gameSheet.appendRow([getFormattedTimestamp(), gameCode, studentID, name, level1Score, level1Time, level2Score, level2Time, score, trialId]);

    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "gameCode": gameCode,
      "studentID": studentID,
      "name": name,
      "score": score,
      "trialId": trialId
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}