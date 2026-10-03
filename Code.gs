// ===================================================================
// 運動賽事前進指揮系統 - 現場交管打卡與 GPS 座標統計雲端後端 (Google Apps Script)
// 支援多賽事多分頁獨立儲存：【2026環法挑戰賽】 ＆ 【Dan Cup 丹盃越野登山車賽】
// 部署方式：貼入現有 Apps Script 專案，點擊「部署」➔「管理部署作業」➔「編輯」➔「新版本」
// ===================================================================

const DEFAULT_FOLDER_TOUR = "2026環法_交管回報照片";
const DEFAULT_SHEET_TOUR = "交管就位與座標統計";

const DEFAULT_FOLDER_DANCUP = "DanCup_交管回報照片";
const DEFAULT_SHEET_DANCUP = "DanCup_交管與打卡統計";

/**
 * 依據 event 參數自動判斷目標工作表與 Google Drive 資料夾
 */
function getTargetConfig(eventName) {
  if (eventName && eventName.toString().toLowerCase() === "dancup") {
    return {
      folderName: DEFAULT_FOLDER_DANCUP,
      sheetName: DEFAULT_SHEET_DANCUP,
      headerColor: "#F97316" // Dan Cup 越野活力橘
    };
  }
  return {
    folderName: DEFAULT_FOLDER_TOUR,
    sheetName: DEFAULT_SHEET_TOUR,
    headerColor: "#FCD34D" // 環法經典黃
  };
}

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

    const eventName = data.event || "tour";
    const config = getTargetConfig(eventName);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let reportSheet = ss.getSheetByName(config.sheetName);
    
    // 若無該賽事分頁則自動建立獨立分頁，不覆寫現有資料！
    if (!reportSheet) {
      reportSheet = ss.insertSheet(config.sheetName);
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
      reportSheet.getRange("A1:J1").setBackground(config.headerColor).setFontWeight("bold");
      reportSheet.setFrozenRows(1);
    }

    // 處理照片儲存至 Google Drive 專屬資料夾
    let photoUrl = "";
    if (data.imageBase64) {
      photoUrl = saveImageToDrive(data.imageBase64, data.pointId, data.userName, config.folderName);
    }

    // 產生 Google Maps 連結
    let mapLink = "無座標";
    if (data.lat && data.lng) {
      mapLink = `https://www.google.com/maps?q=${data.lat},${data.lng}`;
    }

    // 寫入 Google Sheet 獨立分頁
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
      message: `【${eventName.toUpperCase()}】打卡與座標回報成功！已寫入分頁「${config.sheetName}」`,
      sheet: config.sheetName,
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
    const action = e.parameter ? e.parameter.action : "";
    const eventName = (e.parameter && e.parameter.event) ? e.parameter.event : "tour";
    const config = getTargetConfig(eventName);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 測試連線
    if (action === "ping") {
      return createJsonResponse({ status: "ok", event: eventName, time: new Date().toISOString() });
    }

    // 讀取該賽事分頁的所有打卡與座標記錄
    let reportSheet = ss.getSheetByName(config.sheetName);
    if (!reportSheet) {
      return createJsonResponse({ status: "success", event: eventName, sheet: config.sheetName, data: [] });
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
      event: eventName,
      sheet: config.sheetName,
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
 * 將 Base64 照片儲存到 Google Drive 專屬資料夾
 */
function saveImageToDrive(base64Data, pointId, userName, folderName) {
  try {
    const targetFolder = folderName || DEFAULT_FOLDER_TOUR;
    let folders = DriveApp.getFoldersByName(targetFolder);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(targetFolder);

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
