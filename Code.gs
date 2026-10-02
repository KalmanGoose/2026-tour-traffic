// ==========================================
// 2026 環法賽 - 現場交管放人打卡與 GPS 座標統計後端
// 部署方式：貼入 Google Apps Script 並發布為 Web 應用程式
// ==========================================

const FOLDER_NAME = "2026環法_交管回報照片";
const SHEET_NAME_REPORTS = "交管就位與座標統計";

/**
 * 處理 POST 請求 (接收志工姓名、點位、GPS 經緯度、現場照片)
 */
function doPost(e) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000); // 避免併發寫入衝突

    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      return createJsonResponse({ status: "error", message: "無有效內容" });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let reportSheet = ss.getSheetByName(SHEET_NAME_REPORTS);
    
    // 若無統計工作表則自動建立，並排版欄位
    if (!reportSheet) {
      reportSheet = ss.insertSheet(SHEET_NAME_REPORTS);
      reportSheet.appendRow([
        "打卡時間",
        "志工姓名",
        "分配點位",
        "緯度 (Lat)",
        "經度 (Lng)",
        "GPS精準度(公尺)",
        "Google Maps 定位查看",
        "現場狀態",
        "現場照片",
        "備註說明"
      ]);
      reportSheet.getRange("A1:J1").setBackground("#FCD34D").setFontWeight("bold");
      reportSheet.setFrozenRows(1);
    }

    // 處理照片儲存至 Google Drive
    let photoUrl = "";
    if (data.imageBase64) {
      photoUrl = saveImageToDrive(data.imageBase64, data.pointId, data.userName);
    }

    // 產生 Google Maps 連結
    let mapLink = "無座標";
    if (data.lat && data.lng) {
      mapLink = `https://www.google.com/maps?q=${data.lat},${data.lng}`;
    }

    // 寫入 Google Sheet
    const now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
    reportSheet.appendRow([
      now,
      data.userName || "",
      data.pointId || "未指定",
      data.lat || "",
      data.lng || "",
      data.accuracy ? `±${data.accuracy}m` : "",
      mapLink,
      data.status || "已就位",
      photoUrl || "無照片",
      data.note || ""
    ]);

    lock.releaseLock();
    return createJsonResponse({
      status: "success",
      message: "打卡與座標回報成功！",
      timestamp: now,
      mapLink: mapLink,
      photoUrl: photoUrl
    });

  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: error.toString()
    });
  }
}

/**
 * 處理 GET 請求 (提供前端讀取最新打卡統計或測試連線)
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 測試連線
    if (action === "ping") {
      return createJsonResponse({ status: "ok", time: new Date().toISOString() });
    }

    // 讀取目前所有人的打卡與座標記錄
    let reportSheet = ss.getSheetByName(SHEET_NAME_REPORTS);
    if (!reportSheet) {
      return createJsonResponse({ status: "success", data: [] });
    }

    const rows = reportSheet.getDataRange().getValues();
    const list = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[0] && !row[1]) continue;
      list.push({
        time: row[0],
        userName: row[1],
        pointId: row[2],
        lat: row[3],
        lng: row[4],
        accuracy: row[5],
        mapLink: row[6],
        status: row[7],
        photoUrl: row[8],
        note: row[9]
      });
    }

    return createJsonResponse({
      status: "success",
      data: list
    });

  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: error.toString()
    });
  }
}

/**
 * 將 Base64 照片儲存到 Google Drive 資料夾
 */
function saveImageToDrive(base64Data, pointId, userName) {
  try {
    let folders = DriveApp.getFoldersByName(FOLDER_NAME);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);

    const parts = base64Data.split(",");
    const meta = parts[0];
    const raw = parts.length > 1 ? parts[1] : parts[0];

    let mimeType = "image/jpeg";
    if (meta.indexOf("image/png") !== -1) mimeType = "image/png";

    const decoded = Utilities.base64Decode(raw);
    const timestamp = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmmss");
    const cleanUser = (userName || "志工").replace(/[\/\\:*?"<>|]/g, "_");
    const cleanPoint = (pointId || "點位").replace(/[\/\\:*?"<>|]/g, "_");
    const fileName = `${cleanPoint}_${cleanUser}_${timestamp}.jpg`;

    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl();
  } catch (err) {
    Logger.log("儲存照片失敗: " + err.toString());
    return "上傳失敗: " + err.toString();
  }
}

/**
 * 輔助函式：產生 JSON 回傳
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
