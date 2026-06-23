function onFormSubmitTrigger(e) {
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName(); // 抓取工作表名稱（含編號，例如 "1. 中信"）
  var range = e.range;
  var row = range.getRow();
  var values = e.namedValues;

  if (sheetName === "系統錯誤紀錄") return;

  var name = values["姓名"] ? values["姓名"][0] : "";

  // 核心欄位
  var gender = values["生理性別"] ? values["生理性別"][0] : "";
  var idNumber = values["身份證字號"] ? values["身份證字號"][0] : "";
  var birth = values["出生年月日"] ? values["出生年月日"][0] : "";
  var edu = values["教育狀況"] ? values["教育狀況"][0].trim() : "";
  var marriage = values["婚姻狀況"] ? values["婚姻狀況"][0].trim() : "";
  var phone = values["聯絡電話(手機)"] ? values["聯絡電話(手機)"][0] : "";
  var address = values["聯絡地址"] ? values["聯絡地址"][0] : "";
  var emName = values["緊急聯絡人姓名"] ? values["緊急聯絡人姓名"][0] : "";
  var emPhone = values["緊急聯絡人電話"] ? values["緊急聯絡人電話"][0] : "";
  var emRelation = values["緊急聯絡人關係"] ? values["緊急聯絡人關係"][0] : "";
  var qSelect = values["問題陳述(可複選)"] ? values["問題陳述(可複選)"][0] : "";
  var sourceSelect = values["資訊來源(可複選)"]
    ? values["資訊來源(可複選)"][0]
    : "";
  var consultTime = values["第一次諮商時間"] ? values["第一次諮商時間"][0] : "";

  // 彈性欄位
  var empId = values["員工編號"] ? values["員工編號"][0] : "無";
  var subCompany = values["隸屬子公司"] ? values["隸屬子公司"][0] : "無";
  var city = values["工作縣市"] ? values["工作縣市"][0] : "無";
  var factoryArea = values["所屬廠區"] ? values["所屬廠區"][0] : "無";
  var identityStatus = values["身分別"] ? values["身分別"][0] : "無";

  // 自動剔除工作表名稱前方的數字與點，用於存檔名
  var cleanCompanyName = sheetName.replace(/^\d+\.\s*/, "");

  // ==========================================
  // 1. 【智慧日期格式化】處理 出生年月日 與 諮商時間
  // ==========================================

  // (A) 處理出生年月日：轉成乾淨的 yyyy/m/d
  var formattedBirth = "";
  if (birth) {
    var bDate = new Date(birth);
    if (!isNaN(bDate.getTime()) && birth.toString().indexOf("GMT") !== -1) {
      formattedBirth =
        bDate.getFullYear() +
        "/" +
        (bDate.getMonth() + 1) +
        "/" +
        bDate.getDate();
    } else {
      formattedBirth = birth; // 原本就是字串的話就維持原樣
    }
  }

  // (B) 處理諮商時間：轉成 m/d(星期) hh:mm (24小時制)
  var formattedConsultTime = "未提供時間";
  if (consultTime) {
    var cDate =
      consultTime instanceof Date ? consultTime : new Date(consultTime);

    if (!isNaN(cDate.getTime())) {
      var month = cDate.getMonth() + 1;
      var date = cDate.getDate();
      var hours = ("0" + cDate.getHours()).slice(-2);
      var minutes = ("0" + cDate.getMinutes()).slice(-2);
      var weekdays = ["日", "一", "二", "三", "四", "五", "六"];
      var dayOfWeek = weekdays[cDate.getDay()];
      formattedConsultTime =
        month + "/" + date + "(" + dayOfWeek + ")" + hours + ":" + minutes;
    } else {
      formattedConsultTime = consultTime;
    }
  }

  // (C) 手機號碼格式化為 09XX-XXX-XXX
  var formattedPhone = "格式錯誤";
  if (phone) {
    // 🔧 修正後：用 String() 強制轉型，並把所有單引號 ' 與非數字全拔光
    var cleanPhone = String(phone).replace(/['\D]/g, "");
    if (cleanPhone.length === 10 && cleanPhone.indexOf("09") === 0) {
      formattedPhone =
        cleanPhone.slice(0, 4) +
        "-" +
        cleanPhone.slice(4, 7) +
        "-" +
        cleanPhone.slice(7);
    } else {
      formattedPhone = phone;
    }
  }

  // ==========================================
  // 3. 利用「時間戳記」自動計算【跨年自動歸零】的流水號
  // ==========================================
  var formattedSeq;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // 等待最多 15 秒，超時會拋出例外

    var timestampStr = values["時間戳記"] ? values["時間戳記"][0] : "";
    var currentYear = new Date().getFullYear();

    if (timestampStr) {
      var dateParts = timestampStr.split(" ")[0].split("/");
      if (dateParts.length === 3) {
        currentYear = parseInt(dateParts[0], 10);
      }
    }
    var twYear = currentYear - 1911;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var timestampColIndex = headers.indexOf("時間戳記") + 1;

    var yearlyCount = 1;
    if (timestampColIndex > 0 && row > 2) {
      var allTimestamps = sheet
        .getRange(2, timestampColIndex, row - 2, 1)
        .getValues();
      for (var i = allTimestamps.length - 1; i >= 0; i--) {
        if (allTimestamps[i][0]) {
          var checkStr = allTimestamps[i][0].toString();
          var checkYear = currentYear;
          var parts = checkStr.split(" ")[0].split("/");
          if (parts.length === 3) {
            checkYear = parseInt(parts[0], 10);
          }
          if (checkYear === currentYear) {
            yearlyCount++;
          }
        }
      }
    }

    var formattedNum = ("000" + yearlyCount).slice(-3);
    formattedSeq = twYear.toString() + formattedNum;
  } finally {
    lock.releaseLock();
  }

  // ==========================================
  // 5. 設定您的雲端資料夾 ID
  // ==========================================
  var scriptProperties = PropertiesService.getScriptProperties();
  var templateFolderId = scriptProperties.getProperty("TEMPLATE_FOLDER_ID");
  var targetFolderId = scriptProperties.getProperty("TARGET_FOLDER_ID");

  var templateFolder, targetFolder;
  try {
    templateFolder = DriveApp.getFolderById(templateFolderId);
    targetFolder = DriveApp.getFolderById(targetFolderId);
  } catch (err) {
    logStatus(cleanCompanyName, name, "失敗", "資料夾 ID 設定錯誤或權限不足");
    return;
  }

  // ==========================================
  // 6. 尋找範本與錯誤紀錄功能
  // ==========================================
  var files = templateFolder.getFilesByName(sheetName);
  if (!files.hasNext()) {
    logStatus(
      cleanCompanyName,
      name,
      "失敗",
      "在範本資料夾中找不到名為【" + sheetName + "】的範本文件",
    );
    return;
  }
  var templateFile = files.next();

  // ==========================================
  // 7. 從 EAP聯絡 總表讀取NCL與諮商所資訊
  // ==========================================
  var nclText = ""; // 預設無 NCL 字眼
  var firmText = ""; // 預設無諮商所後綴

  try {
    var eapSS = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("EAP_SHEET_ID"),
    );
    var eapSheetName = twYear + "年";
    var eapSheet = eapSS.getSheetByName(eapSheetName);

    if (eapSheet) {
      var eapLastRow = eapSheet.getLastRow();
      if (eapLastRow >= 2) {
        var eapHeaders = eapSheet
          .getRange(1, 1, 1, eapSheet.getLastColumn())
          .getValues()[0];
        var colPhone =
          eapHeaders.indexOf("手機") !== -1
            ? eapHeaders.indexOf("手機")
            : eapHeaders.indexOf("電話");
        var colNcl = eapHeaders.indexOf("個案");
        var colFirm = eapHeaders.indexOf("地點");

        if (colPhone === -1 || colNcl === -1 || colFirm === -1) {
          throw new Error(
            "EAP總表找不到必要欄位（手機／個案／地點），請確認標題列",
          );
        }

        var maxCol = Math.max(colPhone, colNcl, colFirm) + 1;
        var eapData = eapSheet
          .getRange(2, 1, eapLastRow - 1, maxCol)
          .getValues();
        var cleanTargetPhone = String(phone).replace(/['\D]/g, "");

        for (var e = eapData.length - 1; e >= 0; e--) {
          var eapPhoneStr = eapData[e][colPhone]
            .toString()
            .replace(/['\D]/g, "");

          if (eapPhoneStr === cleanTargetPhone) {
            if (eapData[e][colNcl].toString().indexOf("NCL") !== -1) {
              nclText = "NCL";
            }
            if (eapData[e][colFirm]) {
              firmText = eapData[e][colFirm].toString().trim();
            }
            break;
          }
        }
      }
    }
  } catch (err) {
    // 防錯機制：維持不加字，不讓主流程崩潰
  }

  // 依諮商日期建立子資料夾
  var finalTargetFolder = targetFolder; // 預設：直接放在目標資料夾
  var folderCreationWarning = false;

  // 嘗試從 formattedConsultTime 解析日期，格式為 "m/d(星期)HH:mm"
  var consultDateLabel = "";
  if (
    formattedConsultTime &&
    formattedConsultTime !== "未提供時間" &&
    consultTime
  ) {
    try {
      var cDateForFolder =
        consultTime instanceof Date ? consultTime : new Date(consultTime);
      if (!isNaN(cDateForFolder.getTime())) {
        // 取得 "m/d" 作為資料夾名稱（例如 "6/23"）
        consultDateLabel =
          cDateForFolder.getMonth() + 1 + "/" + cDateForFolder.getDate();
      }
    } catch (err) {
      consultDateLabel = "";
    }
  }

  if (consultDateLabel) {
    // 在目標資料夾中尋找是否已存在同名子資料夾
    var subFolders = targetFolder.getFoldersByName(consultDateLabel);
    if (subFolders.hasNext()) {
      finalTargetFolder = subFolders.next(); // 已存在，直接使用
    } else {
      finalTargetFolder = targetFolder.createFolder(consultDateLabel); // 不存在，新建立
    }
  } else {
    // 讀取不到諮商日期，維持放在原目標資料夾，並寄信通知
    folderCreationWarning = true;
    var alertEmail =
      PropertiesService.getScriptProperties().getProperty("ALERT_EMAIL");
    if (alertEmail) {
      MailApp.sendEmail(
        alertEmail,
        "【自動化系統通知】個案諮商表格無法歸入日期資料夾",
        "您好：\n\n系統在製作以下個案的諮商表格時，無法讀取諮商日期，檔案已直接存放於目標根資料夾，請人工確認歸檔位置。\n\n" +
          "━━━━━━━━━━━━━━━━━━━━━━\n" +
          "• 企業名稱：" +
          cleanCompanyName +
          "\n" +
          "• 個案姓名：" +
          name +
          "\n" +
          "• 手機號碼：" +
          formattedPhone +
          "\n" +
          "• 諮商時間欄位原始值：" +
          (consultTime || "（空白）") +
          "\n" +
          "━━━━━━━━━━━━━━━━━━━━━━\n\n" +
          "本信件由系統自動發送。",
      );
    }
  }

  // 若有 NCL 會變 (NCL王大明)，若有諮商所會直接黏在檔名最後面
  var newFileName =
    formattedSeq +
    cleanCompanyName +
    "諮商表格(" +
    nclText +
    name +
    ")" +
    firmText;
  var newFile = templateFile.makeCopy(newFileName, finalTargetFolder);
  var doc = DocumentApp.openById(newFile.getId());
  var body = doc.getBody();

  // ==========================================
  // 8. 進行「一般與彈性欄位」文字代換
  // ==========================================
  body.replaceText("{{姓名}}", name);
  body.replaceText("{{身份證字號}}", idNumber);
  body.replaceText("{{出生年月日}}", formattedBirth);
  body.replaceText("{{聯絡電話\\(手機\\)}}", formattedPhone);
  body.replaceText("{{聯絡地址}}", address);
  body.replaceText("{{諮商時間}}", formattedConsultTime);
  body.replaceText("{{員工編號}}", empId);
  body.replaceText("{{子公司}}", subCompany);
  body.replaceText("{{工作縣市}}", city);
  body.replaceText("{{所屬廠區}}", factoryArea);
  body.replaceText("{{身分別}}", identityStatus);
  body.replaceText("{{緊急聯絡人姓名}}", emName);
  body.replaceText("{{緊急聯絡人電話}}", emPhone);
  body.replaceText("{{緊急聯絡人關係}}", emRelation);

  // ==========================================
  // 9. 【核取方塊代換邏輯】
  // ==========================================

  // (1) 生理性別
  body.replaceText("{{性別_男}}", gender === "男" ? "■" : "□");
  body.replaceText("{{性別_女}}", gender === "女" ? "■" : "□");

  // (2) 教育狀況 (單選、含自填其他)
  var eduList = [
    "國（初）中以下",
    "高中（職）",
    "五專（二專/三專/五專）",
    "大學/二技/四技",
    "碩士",
    "博士",
  ];
  body.replaceText("{{教育_國中}}", edu === "國（初）中以下" ? "■" : "□");
  body.replaceText("{{教育_高中}}", edu === "高中（職）" ? "■" : "□");
  body.replaceText(
    "{{教育_五專}}",
    edu === "五專（二專/三專/五專）" ? "■" : "□",
  );
  body.replaceText("{{教育_大學}}", edu === "大學/二技/四技" ? "■" : "□");
  body.replaceText("{{教育_碩士}}", edu === "碩士" ? "■" : "□");
  body.replaceText("{{教育_博士}}", edu === "博士" ? "■" : "□");
  if (edu && eduList.indexOf(edu) === -1) {
    body.replaceText("{{教育_其他}}", "■ 其他：" + edu);
  } else {
    body.replaceText("{{教育_其他}}", "□ 其他：");
  }

  // (3) 婚姻狀況 (單選、含自填其他)
  var marriageList = ["未婚", "已婚", "離婚", "同居", "分居"];
  body.replaceText("{{婚姻_未婚}}", marriage === "未婚" ? "■" : "□");
  body.replaceText("{{婚姻_已婚}}", marriage === "已婚" ? "■" : "□");
  body.replaceText("{{婚姻_離婚}}", marriage === "離婚" ? "■" : "□");
  body.replaceText("{{婚姻_同居}}", marriage === "同居" ? "■" : "□");
  body.replaceText("{{婚姻_分居}}", marriage === "分居" ? "■" : "□");
  if (marriage && marriageList.indexOf(marriage) === -1) {
    body.replaceText("{{婚姻_其他}}", "■ 其他：" + marriage);
  } else {
    body.replaceText("{{婚姻_其他}}", "□ 其他：");
  }

  // (4) 問題陳述 (複選、含自填其他)
  var qList = [
    "壓力與情緒",
    "情感與關係",
    "伴侶與家族",
    "嬰幼兒問題",
    "親子議題",
    "青少年議題",
    "自我成長",
    "職涯或轉職",
    "身心疾病",
    "失落與創傷",
    "性或性別議題",
    "失眠困擾",
  ];
  var qOther = [];
  body.replaceText(
    "{{問_壓力}}",
    qSelect.indexOf("壓力與情緒") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_情感}}",
    qSelect.indexOf("情感與關係") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_伴侶}}",
    qSelect.indexOf("伴侶與家族") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_嬰幼}}",
    qSelect.indexOf("嬰幼兒問題") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_親子}}",
    qSelect.indexOf("親子議題") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_青少}}",
    qSelect.indexOf("青少年議題") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_自我}}",
    qSelect.indexOf("自我成長") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_職涯}}",
    qSelect.indexOf("職涯或轉職") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_身心}}",
    qSelect.indexOf("身心疾病") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_失落}}",
    qSelect.indexOf("失落與創傷") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_性別}}",
    qSelect.indexOf("性或性別議題") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{問_失眠}}",
    qSelect.indexOf("失眠困擾") !== -1 ? "■" : "□",
  );
  if (qSelect) {
    var userQItems = qSelect.split(", ");
    for (var i = 0; i < userQItems.length; i++) {
      if (qList.indexOf(userQItems[i]) === -1) qOther.push(userQItems[i]);
    }
  }
  if (qOther.length > 0) {
    body.replaceText("{{問_其他}}", "■ 其他問題：" + qOther.join(", "));
  } else {
    body.replaceText("{{問_其他}}", "□ 其他問題：");
  }

  // (5) 資訊來源 (複選、含自填其他)
  var sList = ["講座", "隨身小卡", "同事推薦", "HR或主管轉介"];
  var sOther = [];
  body.replaceText(
    "{{來源_講座}}",
    sourceSelect.indexOf("講座") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{來源_小卡}}",
    sourceSelect.indexOf("隨身小卡") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{來源_同事}}",
    sourceSelect.indexOf("同事推薦") !== -1 ? "■" : "□",
  );
  body.replaceText(
    "{{來源_HR}}",
    sourceSelect.indexOf("HR或主管轉介") !== -1 ? "■" : "□",
  );
  if (sourceSelect) {
    var userSItems = sourceSelect.split(", ");
    for (var i = 0; i < userSItems.length; i++) {
      if (sList.indexOf(userSItems[i]) === -1) sOther.push(userSItems[i]);
    }
  }
  if (sOther.length > 0) {
    body.replaceText("{{來源_其它}}", "■ 其它：" + sOther.join(", "));
  } else {
    body.replaceText("{{來源_其它}}", "□ 其它：");
  }

  doc.saveAndClose();

  // ==========================================
  // 10. 成功完成後，自動將 Google Sheet 的「已提供」欄位勾選起來或整列反灰
  // ==========================================
  var providedColIndex = headers.indexOf("已提供") + 1;
  if (providedColIndex > 0) {
    // 條件一：如果這家企業有「已提供」欄位，就自動打勾
    sheet.getRange(row, providedColIndex).setValue(true);
  } else {
    // 條件二：如果沒有該欄位，就將該列（從第 1 欄到最後一欄）整列刷成淺灰色 1
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#f3f3f3");
  }

  updateYearlyChecklist(formattedPhone, twYear);
  updateCaseNameInEapSheet(formattedPhone, name, twYear);

  logStatus(
    cleanCompanyName,
    name,
    "成功",
    "已成功自動生成 Word 表格檔案",
    formattedPhone,
    formattedConsultTime,
  );
}

// 【修正調整】擴充輔助功能，讓它能安全接收 6 個參數、加入絕對防呆，沒傳入手機和時間時自動用空字串代替，不會卡死
function logStatus(company, name, status, msg, phone, time) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errorSheet = ss.getSheetByName("系統錯誤紀錄");
  if (errorSheet) {
    var p = typeof phone !== "undefined" ? phone : "";
    var t = typeof time !== "undefined" ? time : "";
    errorSheet.appendRow([new Date(), company, name, status, msg, p, t]);
    SpreadsheetApp.flush();
  }
}

// =========================================================================
// 功能 B：定時發送【今日自動化工作日報表】（設定每天 17:00 執行一次）
// =========================================================================
function sendDailyErrorSummaryEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errorSheet = ss.getSheetByName("系統錯誤紀錄");
  if (!errorSheet) return;

  var lastRow = errorSheet.getLastRow();
  if (lastRow < 1) return;

  var recordData = errorSheet.getRange(1, 1, lastRow, 7).getValues();

  var successList = [];
  var failureList = [];

  for (var i = 0; i < recordData.length; i++) {
    var execTime = recordData[i][0];
    var company = recordData[i][1];
    var name = recordData[i][2];
    var status = recordData[i][3];
    var msg = recordData[i][4];
    var phone = recordData[i][5];
    var time = recordData[i][6];

    // 1. 處理【諮商時間】的整形手術
    var formattedTime =
      time && time.toString().trim() ? time.toString().trim() : "未提供";

    // 2. 處理【系統製作時間】的整形手術
    var formattedExecTime = "";
    if (execTime instanceof Date) {
      formattedExecTime = Utilities.formatDate(
        execTime,
        Session.getScriptTimeZone(),
        "MM/dd HH:mm",
      );
    } else {
      formattedExecTime = execTime ? execTime.toString() : "未知時間";
    }

    if (status === "成功") {
      successList.push(
        "• [" +
          formattedExecTime +
          " 成功] " +
          company +
          " " +
          name +
          " | 手機：" +
          phone +
          " | 來談時間：" +
          formattedTime,
      );
    } else if (status === "失敗") {
      failureList.push(
        "• [" +
          formattedExecTime +
          " 錯誤] " +
          company +
          " " +
          name +
          " -> 原因：" +
          msg,
      );
    }
  }

  var emailBody =
    "您好：\n\n以下是今日自動化系統的【個案資料製作執行日報表】\n";
  emailBody += "========================================\n\n";

  emailBody +=
    "✅【今日成功自動製作清單】(共 " + successList.length + " 筆)：\n";
  if (successList.length > 0) {
    emailBody += successList.join("\n") + "\n";
  } else {
    emailBody += "（今日無成功製作之檔案）\n";
  }

  emailBody += "\n----------------------------------------\n\n";

  emailBody +=
    "❌【今日系統遭遇錯誤清單】(共 " + failureList.length + " 筆)：\n";
  if (failureList.length > 0) {
    emailBody +=
      failureList.join("\n") + "\n\n請儘速檢查您的雲端範本名稱或權限設定。\n";
  } else {
    emailBody += "（今日系統無任何錯誤紀錄）\n";
  }

  emailBody += "\n========================================\n";
  emailBody +=
    "本信件由系統每日 17:00 自動彙總發送，本日紀錄已在發信後自動重設。";

  var recipients =
    PropertiesService.getScriptProperties().getProperty("SUMMARY_EMAILS");
  MailApp.sendEmail(
    recipients,
    "【自動化系統】今日個案資料製作執行日報表",
    emailBody,
  );

  if (lastRow > 1) {
    errorSheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }
}

// =========================================================================
// 第二模組：個案資料批發查詢與再製模組
// =========================================================================
function batchSearchAndGenerateWord() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var searchSheet = ss.getSheetByName("批量製作資料");
  if (!searchSheet) {
    SpreadsheetApp.getUi().alert(
      "系統找不到名為【批量製作資料】的工作表，請先建立分頁。",
    );
    return;
  }

  var lastRow = searchSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("請先在 A 欄輸入想要查詢的手機號碼！");
    return;
  }

  var searchPhones = searchSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var sheets = ss.getSheets();

  var sheetDataCache = {};
  for (var k = 0; k < sheets.length; k++) {
    var sName = sheets[k].getName();
    if (sName === "系統錯誤紀錄" || sName === "批量製作資料") continue;
    sheetDataCache[sName] = sheets[k].getDataRange().getValues();
  }

  for (var s = 0; s < searchPhones.length; s++) {
    var rawSearchPhone = searchPhones[s][0].toString().trim();
    if (!rawSearchPhone) continue;

    var cleanSearchPhone = rawSearchPhone.replace(/['\D]/g, "");
    var found = false;
    var rowResult = "";
    var showTimeResult = "—";

    // 新增：跨企業時間的紀錄
    var latestTimestamp = new Date(0);
    var targetSheetToMake = null;
    var targetRowToMake = 0;

    for (var k = 0; k < sheets.length; k++) {
      var currentSheet = sheets[k];
      var currentSheetName = currentSheet.getName();

      if (
        currentSheetName === "系統錯誤紀錄" ||
        currentSheetName === "批量製作資料"
      )
        continue;

      var sheetLastRow = currentSheet.getLastRow();
      if (sheetLastRow < 2) continue;

      var cachedData = sheetDataCache[currentSheetName];
      var headers = cachedData[0];
      var phoneColIndex = headers.indexOf("聯絡電話(手機)") + 1;

      if (phoneColIndex === 0) continue;

      var allSheetPhones = cachedData.slice(1).map(function (row) {
        return [row[phoneColIndex - 1]];
      });

      for (var r = allSheetPhones.length - 1; r >= 0; r--) {
        var sheetPhoneStr = allSheetPhones[r][0]
          .toString()
          .replace(/['\D]/g, "");

        if (sheetPhoneStr === cleanSearchPhone) {
          found = true;
          var targetRow = r + 2;

          // 1. 抓出這筆資料的時間戳記
          var tsColIndex = headers.indexOf("時間戳記") + 1;
          var rowTimestamp =
            tsColIndex > 0
              ? new Date(
                  currentSheet.getRange(targetRow, tsColIndex).getValue(),
                )
              : new Date(0);

          // 2. 如果這家公司的填表時間更晚，就更新最高紀錄本
          if (rowTimestamp > latestTimestamp) {
            latestTimestamp = rowTimestamp;
            targetSheetToMake = currentSheet;
            targetRowToMake = targetRow;
          }
          break;
        }
      } // 👈 內層迴圈結束 (這家公司看完了)
    } // 👈 外層迴圈結束 (大腦已經把 40 幾家公司全看完了！！)

    // ==========================================
    // 3. 當所有公司都看完了，才在最外面進行「唯一一次」的最新資料撈取與 Word 製作
    // ==========================================
    if (found && targetSheetToMake !== null) {
      var finalHeaders = targetSheetToMake
        .getRange(1, 1, 1, targetSheetToMake.getLastColumn())
        .getValues()[0];

      // (A) 撈取最新這家公司的個案姓名
      var finalNameCol = finalHeaders.indexOf("姓名") + 1;
      var caseName =
        finalNameCol > 0
          ? targetSheetToMake.getRange(targetRowToMake, finalNameCol).getValue()
          : "未知";

      // (B) 撈取最新這家公司的諮商時間（用來呈現在查詢分頁的 D 欄）
      var timeColIndex = finalHeaders.indexOf("第一次諮商時間") + 1;
      var rawConsultTime =
        timeColIndex > 0
          ? targetSheetToMake.getRange(targetRowToMake, timeColIndex).getValue()
          : "";
      showTimeResult = "未提供時間";

      if (rawConsultTime) {
        var cDate = new Date(rawConsultTime);
        if (!isNaN(cDate.getTime())) {
          var month = cDate.getMonth() + 1;
          var date = cDate.getDate();
          var hours = ("0" + cDate.getHours()).slice(-2);
          var minutes = ("0" + cDate.getMinutes()).slice(-2);
          var weekdays = ["日", "一", "二", "三", "四", "五", "六"];
          var dayOfWeek = weekdays[cDate.getDay()];
          showTimeResult =
            month + "/" + date + "(" + dayOfWeek + ")" + hours + ":" + minutes;
        } else {
          showTimeResult = rawConsultTime.toString();
        }
      }

      // (C) 真正動工製作唯一一份最準確的 Word 檔案
      try {
        var mockEvent = createMockEvent(
          targetSheetToMake,
          targetRowToMake,
          finalHeaders,
        );
        onFormSubmitTrigger(mockEvent);
        rowResult =
          "✅ 已製作：" +
          targetSheetToMake.getName().replace(/^\d+\.\s*/, "") +
          " " +
          caseName;
      } catch (err) {
        rowResult =
          "⚠️ 找到：" +
          targetSheetToMake.getName() +
          " " +
          caseName +
          " 但製作Word失敗: " +
          err.message;
      }
    }

    if (!found) {
      rowResult = "❌ 查無資料";
      showTimeResult = "—";
    }

    var currentRow = s + 2;
    searchSheet.getRange(currentRow, 2).setValue(rowResult);
    searchSheet.getRange(currentRow, 3).setValue(new Date());
    searchSheet.getRange(currentRow, 4).setValue(showTimeResult);
  }

  SpreadsheetApp.getUi().alert("🎉 個案資料製作已執行完畢！");
}

function createMockEvent(sheet, row, headers) {
  var rowValues = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  var namedValues = {};
  for (var i = 0; i < headers.length; i++) {
    var headerName = headers[i].toString().trim();
    if (headerName) {
      if (rowValues[i] instanceof Date) {
        if (headerName === "時間戳記") {
          var d = rowValues[i];
          namedValues[headerName] = [
            d.getFullYear() +
              "/" +
              (d.getMonth() + 1) +
              "/" +
              d.getDate() +
              " 上午 00:00:00",
          ];
        } else {
          namedValues[headerName] = [rowValues[i]];
        }
      } else {
        if (rowValues[i] instanceof Date) {
          namedValues[headerName] = isNaN(rowValues[i].getTime())
            ? [""]
            : [rowValues[i]];
        } else {
          namedValues[headerName] = [
            rowValues[i] !== null && rowValues[i] !== undefined
              ? rowValues[i].toString()
              : "",
          ];
        }
      }
    }
  }

  return {
    range: sheet.getRange(row, 1),
    namedValues: namedValues,
  };
}

// =========================================================================
// 擴充功能：在 Google Sheet 上方建立一鍵點擊的實用按鈕選單
// =========================================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🛠️ 批量製作資料")
    .addItem("點選即開始製作", "batchSearchAndGenerateWord")
    .addToUi();
}

// =========================================================================
// 第三模組：新試算表年度總表自動註記工具（由主功能自動調用）
// =========================================================================
function updateYearlyChecklist(phone, twYear) {
  var myEmail =
    PropertiesService.getScriptProperties().getProperty("ALERT_EMAIL");
  var targetSS = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("EAP_SHEET_ID"),
  );

  var sheetName = twYear + "年";
  var targetSheet = targetSS.getSheetByName(sheetName);

  if (!targetSheet) {
    MailApp.sendEmail(
      myEmail,
      "【自動化系統警告】找不到年度工作表",
      "系統試圖註記但找不到名為【" + sheetName + "】的工作表。",
    );
    return;
  }

  var lastRow = targetSheet.getLastRow();
  if (lastRow < 2) {
    MailApp.sendEmail(
      myEmail,
      "【自動化系統通知】年度工作表尚無資料",
      "工作表【" + sheetName + "】目前是空的。",
    );
    return;
  }

  // 先讀標題列，動態找欄位索引
  var headers = targetSheet
    .getRange(1, 1, 1, targetSheet.getLastColumn())
    .getValues()[0];
  var colDone = headers.indexOf("已完成");
  var colPhone =
    headers.indexOf("手機") !== -1
      ? headers.indexOf("手機")
      : headers.indexOf("電話");
  var colFilled = headers.indexOf("表格已填");

  if (colDone === -1 || colPhone === -1 || colFilled === -1) {
    MailApp.sendEmail(
      myEmail,
      "【自動化系統警告】EAP總表找不到必要欄位",
      "在【" +
        sheetName +
        "】找不到「已完成」、「手機」或「表格已填」欄位，請確認標題列。",
    );
    return;
  }

  var maxCol = Math.max(colDone, colPhone, colFilled) + 1;
  var dataRange = targetSheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
  var cleanTargetPhone = phone.toString().replace(/['\D]/g, "");
  var foundAndMarked = false;

  for (var i = dataRange.length - 1; i >= 0; i--) {
    var currentPhoneStr = dataRange[i][colPhone]
      .toString()
      .replace(/['\D]/g, "");
    var isCheckboxTrue = dataRange[i][colDone];

    if (currentPhoneStr === cleanTargetPhone) {
      if (isCheckboxTrue === false) {
        var targetRow = i + 2;
        targetSheet.getRange(targetRow, colFilled + 1).setValue("V(系統)");
        foundAndMarked = true;
        break;
      }
    }
  }

  // 若繞完整張表，完全沒有搜尋到手機相符且欄A為false的登記資料，寄信通知
  if (!foundAndMarked) {
    MailApp.sendEmail(
      myEmail,
      "【自動化系統通知】EAP聯絡總表查無可註記之個案",
      "個案手機號碼：【" +
        phone +
        "】已成功製作 Word 檔，但在新表的【" +
        sheetName +
        "】工作表中找不到手機相符且欄A為「未勾選(false)」的登記紀錄，無法自動註記，請人工核對。",
    );
  }
}

// =========================================================================
// 完成製作後，將 EAP 總表的「個案」欄補全為全名+稱呼
// =========================================================================
function updateCaseNameInEapSheet(phone, fullName, twYear) {
  try {
    var targetSS = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("EAP_SHEET_ID"),
    );

    var sheetName = twYear + "年";
    var targetSheet = targetSS.getSheetByName(sheetName);
    if (!targetSheet) {
      Logger.log("[updateCase] 找不到工作表：" + sheetName);
      return;
    }

    var lastRow = targetSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("[updateCase] 工作表無資料");
      return;
    }

    var headers = targetSheet
      .getRange(1, 1, 1, targetSheet.getLastColumn())
      .getValues()[0];

    var colPhone =
      headers.indexOf("手機") !== -1
        ? headers.indexOf("手機")
        : headers.indexOf("電話");
    var colCase = headers.indexOf("個案");

    Logger.log("[updateCase] colPhone=" + colPhone + " colCase=" + colCase);

    if (colPhone === -1 || colCase === -1) {
      Logger.log(
        "[updateCase] 找不到必要欄位，headers=" + JSON.stringify(headers),
      );
      return;
    }

    var maxCol = Math.max(colPhone, colCase) + 1;
    var data = targetSheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
    var cleanTargetPhone = phone.toString().replace(/['\D]/g, "");

    Logger.log(
      "[updateCase] 搜尋手機=" +
        cleanTargetPhone +
        " 全名=" +
        fullName +
        " 年度=" +
        twYear,
    );

    for (var i = data.length - 1; i >= 0; i--) {
      var rowPhone = data[i][colPhone].toString().replace(/['\D]/g, "");

      // 每筆都印出來比對
      Logger.log("[updateCase] 第" + (i + 2) + "列 手機=" + rowPhone);

      if (rowPhone !== cleanTargetPhone) continue;

      var original = data[i][colCase].toString().trim();
      Logger.log("[updateCase] 找到！原始個案欄=" + original);

      var hasNcl = original.indexOf("NCL") !== -1;
      var prefix = hasNcl ? "NCL" : "";
      var hasSir = original.indexOf("先生") !== -1;
      var hasMiss = original.indexOf("小姐") !== -1;

      if (!hasSir && !hasMiss) {
        Logger.log("[updateCase] 無稱呼，略過不處理");
        break;
      }

      var suffix = hasSir ? "先生" : "小姐";
      var newValue = prefix + fullName + suffix;

      Logger.log("[updateCase] 原值=" + original + " 新值=" + newValue);

      if (newValue === original) {
        Logger.log("[updateCase] 新舊相同，不寫入");
        break;
      }

      targetSheet.getRange(i + 2, colCase + 1).setValue(newValue);
      Logger.log("[updateCase] 已寫入第" + (i + 2) + "列");
      break;
    }
  } catch (err) {
    Logger.log("[updateCase] 發生錯誤：" + err.message);
  }
}
