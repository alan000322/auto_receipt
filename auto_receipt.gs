// 全局變量，用於存儲待處理的行
var pendingRows = [];

function onEditTrigger(e) {
  Logger.log('onEditTrigger executed');
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // 檢查是否有多行被編輯
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  
  for (var i = 0; i < numRows; i++) {
    var currentRow = startRow + i;
    var jCellValue = sheet.getRange(currentRow, 10).getValue().toString().trim();
    var kCellValue = sheet.getRange(currentRow, 11).getValue();
    
    if (jCellValue === "未電子開立" && kCellValue === "") {
      Logger.log('Row ' + currentRow + ' needs processing');
      sheet.getRange(currentRow, 11).setValue("等待處理...");
      pendingRows.push(currentRow);
    }
  }
  
  // 如果有待處理的行，開始處理
  if (pendingRows.length > 0) {
    Logger.log('Starting to process ' + pendingRows.length + ' rows');
    processPendingRows(sheet);
  }
}

function processPendingRows(sheet) {
  while (pendingRows.length > 0) {
    var row = pendingRows.shift(); // 從隊列中取出第一個行
    Logger.log('Processing row: ' + row);
    sheet.getRange(row, 11).setValue("製作中...");
    
    try {
      generateReceipt(sheet, row);
      sheet.getRange(row, 11).setValue("自動程式已完成收據製作");
      Logger.log('Receipt generated successfully for row ' + row);
    } catch (error) {
      Logger.log('Error processing row ' + row + ': ' + error.toString());
      sheet.getRange(row, 11).setValue("錯誤: " + error.toString());
    }
  }
}

// 輔助函數：格式化金額
function formatAmount(amount) {
  // 檢查 amount 是否為數字或可以轉換為數字
  var num = parseFloat(amount);
  if (isNaN(num)) {
    return amount; // 如果不是有效數字，返回原始值
  }
  
  // 將數字轉換為字符串，並用正則表達式添加逗號
  return num.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// Main Functions
function generateReceipt(sheet, row) {
  Logger.log('Starting generateReceipt for row: ' + row);
  
  try {
    var data = sheet.getRange(row, 1, 1, 11).getValues()[0]; // 注意：現在獲取11列數據
    Logger.log('Retrieved data: ' + JSON.stringify(data));
    
    var [序號, 姓名, 開立日期, 證號, 項目, 用途, 總額, 地址, 電話, 收據狀態] = data;
    
    // 格式化日期
    var formattedDate = Utilities.formatDate(new Date(開立日期), Session.getScriptTimeZone(), "yyyy/MM/dd");
    Logger.log('Formatted date: ' + formattedDate);

    // 格式化金額
    var formattedAmount = formatAmount(總額);
    Logger.log('Formatted amount: ' + formattedAmount);

    // 生成中文大寫金額
    // var chineseAmount = numberToChinese(總額);
    // Logger.log('Chinese amount: ' + chineseAmount);

    // 開啟模板文件
    var templateFile = DriveApp.getFilesByName("收據模板").next();
    Logger.log('Template file found: ' + templateFile.getName());
    
    var newDoc = templateFile.makeCopy("收據_" + 姓名 + "_" + 序號, DriveApp.getFoldersByName("收據製作").next());
    Logger.log('New document created: ' + newDoc.getName());
    
    var doc = DocumentApp.openById(newDoc.getId());
    var body = doc.getBody();
    
    // 替換文檔中的佔位符
    Logger.log('Replacing placeholders in the document');
    body.replaceText("{{序號}}", 序號);
    body.replaceText("{{姓名}}", 姓名);
    body.replaceText("{{開立日期}}", formattedDate);
    body.replaceText("{{證號}}", 證號);
    body.replaceText("{{項目}}", 項目);
    body.replaceText("{{用途}}", 用途);
    body.replaceText("{{總額}}", formattedAmount);
    // body.replaceText("{{中文總額}}", chineseAmount); // 新增的中文大寫金額
    body.replaceText("{{地址}}", 地址);
    body.replaceText("{{電話}}", 電話);
    
    doc.saveAndClose();
    Logger.log('Document saved and closed');
    
    // Convert to PDF
    var pdf = DriveApp.getFileById(newDoc.getId()).getAs("application/pdf");
    // Find the "收據製作" folder
    var receiptFolder = DriveApp.getFoldersByName("收據製作").next();
    
    // Find or create the "收據" subfolder within "收據製作"
    var pdfFolder;
    var subFolders = receiptFolder.getFoldersByName("收據");
    if (subFolders.hasNext()) {
      pdfFolder = subFolders.next();
    } else {
      pdfFolder = receiptFolder.createFolder("收據");
    }
    
    Logger.log('PDF will be saved in folder: ' + pdfFolder.getName());
    
    var pdfFile = pdfFolder.createFile(pdf);
    pdfFile.setName("收據_" + 姓名 + "_" + 序號 + ".pdf");
    Logger.log('PDF created: ' + pdfFile.getName());
    
    // 在成功生成 PDF 後，更新 K 欄的狀態
    sheet.getRange(row, 11).setValue("自動程式已完成收據製作");
    Logger.log('Receipt generated successfully, spreadsheet updated');
    
  } catch (error) {
    Logger.log('Error in generateReceipt: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    // 如果出錯，更新 K 欄的狀態
    sheet.getRange(row, 11).setValue("錯誤: " + error.toString());
  }
}

// 手動測試函數
function manualTest() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var testRow = 2; // 假設要測試第二行
  Logger.log('Starting manual test for row: ' + testRow);
  generateReceipt(sheet, testRow);
  Logger.log('Manual test completed');
}
