// ==========================================
// 1. 網頁渲染 (網頁標題：包子記帳)
// ==========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('包子記帳')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

// 💡 效能優化：精準讀取有效資料，避開幾千行的空白列
function getSheetDataOptimized(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

// ==========================================
// 2. 獲取資料 (極速同步核心)
// ==========================================
function getAppData() {
  var acctData = getSheetDataOptimized('帳戶與信用卡');
  var banks = [];
  var cards = [];
  var dashboard = { banks: [], cards: [], totalAssets: 0, totalLiabilities: 0, netWorth: 0, investAssets: 0, liquidAssets: 0 };
  
  for(var i=1; i<acctData.length; i++) {
    if(acctData[i][0]) {
       var bankName = acctData[i][0].toString();
       var balance = Number(acctData[i][2]) || 0;
       var note = acctData[i][3] ? acctData[i][3].toString() : ''; 
       banks.push(bankName);
       dashboard.banks.push({ name: bankName, balance: balance, note: note });
       dashboard.totalAssets += balance;
       
       // 💡 智慧辨識投資帳戶 (僅看富邦證券)
       if (bankName.indexOf('富邦證券') !== -1) {
           dashboard.investAssets += balance;
       }
    }
    if(acctData[i][5]) {
       var debt = Number(acctData[i][7]) || 0; 
       cards.push(acctData[i][5]);
       dashboard.cards.push({ 
         name: acctData[i][5], 
         totalDebt: debt, 
         remain: acctData[i][8], 
         billDate: acctData[i][9], 
         payDate: acctData[i][10], 
         currentBill: Number(acctData[i][12]) || 0 
       });
       dashboard.totalLiabilities += debt;
    }
  }
  dashboard.liquidAssets = dashboard.totalAssets - dashboard.investAssets;
  dashboard.netWorth = dashboard.totalAssets - dashboard.totalLiabilities;
  
  var catData = getSheetDataOptimized('AI 專用分析與預算');
  var categories = [];
  for(var i=1; i<catData.length; i++) {
    if(catData[i][0]) categories.push(catData[i][0].toString());
  }
  
  return { banks: banks, cards: cards, categories: categories, dashboard: dashboard };
}

function getSyncData(monthStr) {
  return { appData: getAppData(), history: getHistory(monthStr) };
}

// ==========================================
// 3. 系統管理功能 (包含極速回傳)
// ==========================================
function addCategory(name, monthStr) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AI 專用分析與預算').appendRow([name, 0, 0, 0]); 
    SpreadsheetApp.flush();
    return { msg: "✅ 分類「" + name + "」新增成功！", syncData: getSyncData(monthStr) };
  } catch(e) { return { msg: "❌ 失敗：" + e.toString(), syncData: null }; }
}

function addBankAccount(bankData, monthStr) {
  try {
    if (typeof bankData === 'string') bankData = { name: bankData, initial: 0, note: '' };
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('帳戶與信用卡');
    var lastRow = sheet.getLastRow() + 1;
    var r = 2; while(sheet.getRange(r, 1).getValue() !== "") r++;
    sheet.getRange(r, 1).setValue(bankData.name); 
    sheet.getRange(r, 2).setValue(bankData.initial);    
    sheet.getRange(r, 4).setValue(bankData.note || '');    
    SpreadsheetApp.flush();
    return { msg: "✅ 銀行「" + bankData.name + "」新增成功！", syncData: getSyncData(monthStr) };
  } catch(e) { return { msg: "❌ 失敗：" + e.toString(), syncData: null }; }
}

function updateBankNote(bankName, newNote, monthStr) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('帳戶與信用卡');
    var data = getSheetDataOptimized('帳戶與信用卡');
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === bankName) {
        sheet.getRange(i + 1, 4).setValue(newNote); 
        SpreadsheetApp.flush();
        return { msg: "✅ 備註更新成功！", syncData: getSyncData(monthStr) };
      }
    }
    return { msg: "❌ 找不到該銀行帳戶", syncData: null };
  } catch (e) { return { msg: "❌ 發生錯誤：" + e.toString(), syncData: null }; }
}

function addCreditCard(cardData, monthStr) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('帳戶與信用卡');
    var r = 2; while(sheet.getRange(r, 6).getValue() !== "") r++;
    sheet.getRange(r, 6).setValue(cardData.name);         
    sheet.getRange(r, 7).setValue(cardData.limit);        
    sheet.getRange(r, 10).setValue(cardData.billDate);    
    sheet.getRange(r, 11).setValue(cardData.payDate);     
    sheet.getRange(r, 12).setValue(cardData.autoPayAcc);  
    sheet.getRange(r, 8).setFormula(`=SUMIFS('紀錄頁'!D:D, '紀錄頁'!F:F, F${r}) - SUMIFS('紀錄頁'!D:D, '紀錄頁'!G:G, F${r})`); 
    sheet.getRange(r, 9).setFormula(`=G${r} - H${r} - SUMIFS('固定支出與轉帳'!J:J, '固定支出與轉帳'!G:G, F${r})`); 
    sheet.getRange(r, 13).setFormula(`=MAX(0, H${r} - SUMIFS('紀錄頁'!D:D, '紀錄頁'!F:F, F${r}, '紀錄頁'!A:A, ">=" & (IF(DAY(TODAY())>=J${r}, DATE(YEAR(TODAY()), MONTH(TODAY()), J${r}), DATE(YEAR(TODAY()), MONTH(TODAY())-1, J${r})) + 1)))`); 
    sheet.getRange(r, 14).setFormula(`=SUMIFS('紀錄頁'!D:D, '紀錄頁'!F:F, F${r}, '紀錄頁'!A:A, ">" & IF(DAY(TODAY())>=J${r}, DATE(YEAR(TODAY()), MONTH(TODAY()), J${r}), DATE(YEAR(TODAY()), MONTH(TODAY())-1, J${r})))`); 
    sheet.getRange(r, 15).setFormula(`=G${r}-I${r}`); 
    SpreadsheetApp.flush();
    return { msg: "✅ 信用卡「" + cardData.name + "」新增成功！\n\n🎉 系統已為您全自動寫入所有計算公式！", syncData: getSyncData(monthStr) };
  } catch(e) { return { msg: "❌ 失敗：" + e.toString(), syncData: null }; }
}

// ==========================================
// 4. 寫入紀錄與 LINE 自動推播通知
// ==========================================
function recordData(data, monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (data.isFixed) {
      var fixedSheet = ss.getSheetByName('固定支出與轉帳');
      var d = new Date(data.date);
      if (data.type === '固定支出' || data.type === '訂閱') data.type = '支出';
      fixedSheet.appendRow([data.item, data.type, data.amount, d.getDate(), data.totalPeriods || "", data.paidPeriods || "", data.accountOut || '', data.accountIn || '', data.category || '']);
      var targetRow = fixedSheet.getLastRow();
      fixedSheet.getRange(targetRow, 10).setFormula('=IF(E' + targetRow + '="", 0, C' + targetRow + '*(E' + targetRow + '-F' + targetRow + '))');
      SpreadsheetApp.flush();
      sendLineMessage("🗓️ 【排程新增】\n已將「" + data.item + "」加入自動扣款排程！");
      return { msg: "✅ 已加入自動扣款排程！", syncData: getSyncData(monthStr) };
      
    } else {
      var sheet = ss.getSheetByName('紀錄頁');
      sheet.appendRow([new Date(data.date), data.type, data.item, data.amount, data.category || '', data.accountOut || '', data.accountIn || '']);
      sortRecordSheet();
      SpreadsheetApp.flush(); 

      var lineMsg = "";
      var amt = Number(data.amount);
      
      // 💡 安靜模式：只推播點數折抵、大額消費(>=3000)、與收入
      if (data.type === '點數折抵') {
        lineMsg = "🎁 【點數折抵成功】\n📍 項目：" + data.item + "\n💰 省下了：$" + Math.abs(amt).toLocaleString() + "\n🎊 小秘書：太棒了！又省下一筆錢囉 🥟";
      } else if (data.type === '支出') {
        if (amt >= 3000 && data.item.indexOf('生日') === -1 && data.item.indexOf('過年') === -1) {
           lineMsg = "💸 【大額消費警報】\n支出：$" + amt.toLocaleString() + "\n📍 項目：" + data.item + "\n💳 帳戶：" + (data.accountOut || '未指定') + "\n⚠️ 請留意本月預算控制喔！";
        }
      } else if (data.type === '收入') {
        lineMsg = "💰 【入帳通知】\n太棒了！「" + data.item + "」進帳了 +$" + amt.toLocaleString() + "！";
      }

      if (lineMsg !== "") {
        sendLineMessage(lineMsg);
      }

      return { msg: "✅ 記帳成功！", syncData: getSyncData(monthStr) };
    }
  } catch(e) { return { msg: "❌ 發生錯誤：" + e.toString(), syncData: null }; }
}

function updateRecordData(data, monthStr) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('紀錄頁');
    sheet.getRange(data.row, 1, 1, 7).setValues([[new Date(data.date), data.type, data.item, data.amount, data.category || '', data.accountOut || '', data.accountIn || '']]);
    sortRecordSheet();
    SpreadsheetApp.flush();
    return { msg: "💾 紀錄更新成功！", syncData: getSyncData(monthStr) };
  } catch(e) { return { msg: "❌ 發生錯誤：" + e.toString(), syncData: null }; }
}

function sortRecordSheet() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('紀錄頁');
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({column: 1, ascending: true});
  } catch(e) { }
}

function deleteRecord(rowNum, monthStr) {
  try { 
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('紀錄頁').deleteRow(rowNum); 
    SpreadsheetApp.flush();
    return { msg: "✅ 刪除成功！", syncData: getSyncData(monthStr) }; 
  } catch(e) { return { msg: "❌ 失敗：" + e.toString(), syncData: null }; }
}

// ==========================================
// 5. 其他資料讀取功能
// ==========================================
function getHistory(monthStr) {
  var data = getSheetDataOptimized('紀錄頁');
  var results = [];
  for(var i = data.length - 1; i >= 1; i--) {
    var dateVal = data[i][0]; if (!dateVal) continue;
    var d = new Date(dateVal), m = (d.getMonth() + 1).toString().padStart(2, '0'), y = d.getFullYear();
    if (y + '-' + m === monthStr) {
      results.push({
        row: i + 1, date: m + '/' + d.getDate().toString().padStart(2, '0'), fullDate: y + '-' + m + '-' + d.getDate().toString().padStart(2, '0'),
        type: data[i][1], item: data[i][2], amount: Number(data[i][3]) || 0, category: data[i][4] || '未分類', accOut: data[i][5] || '', accIn: data[i][6] || '' 
      });
    }
  }
  return results;
}

function getAccountDetails(accountName) {
  var data = getSheetDataOptimized('紀錄頁'), details = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][5] === accountName || data[i][6] === accountName) {
      var d = new Date(data[i][0]), dateStr = (d.getMonth() + 1) + '/' + d.getDate();
      var effect = (data[i][6] === accountName) ? "+ $" + data[i][3].toLocaleString() : "- $" + data[i][3].toLocaleString();
      details.push({ date: dateStr, item: data[i][2], type: data[i][1], effect: effect, effectClass: (data[i][6] === accountName ? "val-pos" : "val-neg") });
    }
    if (details.length >= 30) break;
  }
  return details;
}

function getFixedExpenses() {
  var data = getSheetDataOptimized('固定支出與轉帳');
  var list = [];
  for(var i = 1; i < data.length; i++) {
    if(data[i][0] && data[i][0] !== "") {
      list.push({ item: data[i][0], type: data[i][1], amount: Number(data[i][2]) || 0, day: data[i][3], total: data[i][4] || '-', paid: data[i][5] || '-', account: data[i][6] || '未指定帳戶', category: data[i][8] || '' });
    }
  }
  return list;
}

// ==========================================
// 6. 🌙 每日晚安排程與報表
// ==========================================
function processDailyFixed() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fixedSheet = ss.getSheetByName('固定支出與轉帳');
  var recordSheet = ss.getSheetByName('紀錄頁');
  
  var todayDate = new Date();
  var today = todayDate.getDate(); 
  var currentMonth = todayDate.getMonth();
  var currentYear = todayDate.getFullYear();
  var lastDayOfMonth = new Date(currentYear, currentMonth + 1, 0).getDate(); 
  var isLastDay = (today === lastDayOfMonth);
  
  var deductedItems = [], completedInstallments = [], hasNewRecords = false; 
  
  var fixedData = fixedSheet.getDataRange().getValues();
  for(var i=1; i<fixedData.length; i++) {
    if (!fixedData[i][0]) continue; 
    var item = fixedData[i][0], type = fixedData[i][1], amount = fixedData[i][2], payDate = parseInt(fixedData[i][3]);
    var totalPeriods = fixedData[i][4], paidPeriods = fixedData[i][5], accOut = fixedData[i][6], accIn = fixedData[i][7], category = fixedData[i][8];
    
    if(payDate === today || (isLastDay && payDate > today)) {
      if(totalPeriods !== "" && paidPeriods !== "") { if(Number(paidPeriods) >= Number(totalPeriods)) continue; }
      recordSheet.appendRow([new Date(), type, item, amount, category, accOut, accIn]);
      deductedItems.push(item + " ($" + amount.toLocaleString() + ")");
      hasNewRecords = true;
      if(totalPeriods !== "" && paidPeriods !== "") {
        var newPaid = Number(paidPeriods) + 1;
        fixedSheet.getRange(i+1, 6).setValue(newPaid);
        if (newPaid >= Number(totalPeriods)) completedInstallments.push(item);
      }
    }
  }
  if (hasNewRecords) sortRecordSheet(); 

  var todaySpent = 0, monthTotalSpent = 0, todayDetails = {}, todayAccountDetails = {};
  var excludeCategories = ['ETF定期定額', '學貸'];
  var recordData = recordSheet.getDataRange().getValues();
  
  for (var i = 1; i < recordData.length; i++) {
    var rDateVal = recordData[i][0]; if (!rDateVal) continue;
    var rDate = new Date(rDateVal), rType = (recordData[i][1] || '').toString().trim(), rItem = (recordData[i][2] || '').toString().trim(), rAmount = parseFloat(recordData[i][3]) || 0, rCategory = (recordData[i][4] || '').toString().trim() || '未分類', rAccOut = (recordData[i][5] || '').toString().trim();
    if (['支出', '固定支出', '訂閱', '固定'].includes(rType) && excludeCategories.indexOf(rCategory) === -1) {
      if (rDate.getFullYear() === currentYear && rDate.getMonth() === currentMonth) {
        monthTotalSpent += rAmount;
        if (rDate.getDate() === today) {
          todaySpent += rAmount;
          var accName = rAccOut || '未指定帳戶';
          todayAccountDetails[accName] = (todayAccountDetails[accName] || 0) + rAmount;
          if (!todayDetails[rCategory]) todayDetails[rCategory] = { total: 0, items: [] };
          todayDetails[rCategory].total += rAmount;
          if (todayDetails[rCategory].items.indexOf(rItem) === -1) todayDetails[rCategory].items.push(rItem);
        }
      }
    }
  }

  var budgetData = ss.getSheetByName('AI 專用分析與預算').getDataRange().getValues(), totalBudget = 0;
  for (var i = 1; i < budgetData.length; i++) { if (budgetData[i][0]) { totalBudget += (Number(budgetData[i][1]) || 0); } }

  var acctData = ss.getSheetByName('帳戶與信用卡').getDataRange().getValues(), bankBalances = {}, cardWarnings = [];
  for (var i = 1; i < acctData.length; i++) { if (acctData[i][0]) bankBalances[acctData[i][0]] = Number(acctData[i][2]) || 0; }
  var todayTime = todayDate.getTime();
  for (var i = 1; i < acctData.length; i++) {
    var cardName = acctData[i][5], payDateRaw = acctData[i][10], deductAcct = acctData[i][11], currentBill = Number(acctData[i][12]) || 0; 
    if (cardName && payDateRaw && currentBill > 0) {
      var cPayDate = parseInt(payDateRaw), daysToPay = -1;
      for (var d = 1; d <= 3; d++) { var futureDate = new Date(todayTime + d * 24 * 60 * 60 * 1000); if (futureDate.getDate() === cPayDate) { daysToPay = d; break; } }
      if (daysToPay !== -1) {
        var acctBalance = bankBalances[deductAcct] || 0;
        if (acctBalance < currentBill) cardWarnings.push("🚨 扣款提醒：" + cardName + " 將於 " + daysToPay + " 天後扣繳，「" + deductAcct + "」餘額不足，請記得補足差額 $" + (currentBill - acctBalance).toLocaleString() + "！");
      }
    }
  }

  var daysLeft = lastDayOfMonth - today, remainingBudget = totalBudget - monthTotalSpent;
  var tomorrowSafeLimit = daysLeft > 0 ? Math.floor(remainingBudget / daysLeft) : remainingBudget; if(tomorrowSafeLimit < 0) tomorrowSafeLimit = 0;

  var weekdays = ['日', '一', '二', '三', '四', '五', '六'];
  var msg = "🌙 【包子秘書晚安總結】 (" + (currentMonth+1) + "/" + today + " 星期" + weekdays[todayDate.getDay()] + ")\n━━━━━━━━━━━━\n";

  if (todaySpent === 0) {
    msg += "📊 今日花費結算：$ 0\n🎉 恭喜達成「無消費日」！\n太棒了，今天完美守住了錢包 🥟\n\n";
  } else {
    msg += "📊 今日花費總計：$ " + todaySpent.toLocaleString() + "\n\n";
    if (Object.keys(todayAccountDetails).length > 0) { for (var acc in todayAccountDetails) { msg += "💳 於「" + acc + "」支出了 $" + todayAccountDetails[acc].toLocaleString() + "\n"; } msg += "\n"; }
    msg += "🏷️ 項目明細：\n";
    var sortedCats = Object.keys(todayDetails).sort(function(a, b) { return todayDetails[b].total - todayDetails[a].total; });
    var numEmojis = ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣', '6️⃣', '7️⃣', '8️⃣', '9️⃣', '🔟'];
    for (var j = 0; j < Math.min(10, sortedCats.length); j++) {
      var cat = sortedCats[j], icon = numEmojis[j] || '🔸', catTotal = todayDetails[cat].total;
      var catItemsStr = todayDetails[cat].items.join('、'); if(catItemsStr.length > 15) catItemsStr = catItemsStr.substring(0, 15) + '...';
      msg += icon + " " + cat + "：$ " + catTotal.toLocaleString() + " (" + catItemsStr + ")\n";
    }
    msg += "\n";
  }

  msg += "━━━━━━━━━━━━\n🏃‍♂️ 預算配速指示\n💡 本月總預算：$ " + totalBudget.toLocaleString() + "\n💡 目前剩餘預算：$ " + remainingBudget.toLocaleString() + "\n";
  if (remainingBudget < 0) { msg += "⚠️ 哎呀！目前預算已透支 $" + Math.abs(remainingBudget).toLocaleString() + "，接下來幾天請盡量暫緩非必要消費喔！🥲\n"; } 
  else { msg += "💡 距離月底還有：" + daysLeft + " 天\n👉 明日安全額度：$ " + tomorrowSafeLimit.toLocaleString() + "\n"; if (todaySpent === 0) msg += "(因為今天沒花錢，明天的可用額度變多囉！)\n"; }

  var alerts = [];
  if (deductedItems.length > 0) alerts.push("🔸 已自動扣款：\n" + deductedItems.join("\n"));
  if (completedInstallments.length > 0) alerts.push("🎉 恭喜！「" + completedInstallments.join('、') + "」分期已全數繳清！");
  if (cardWarnings.length > 0) alerts = alerts.concat(cardWarnings);
  if (alerts.length > 0) { msg += "━━━━━━━━━━━━\n🤖 系統自動執行與警報\n" + alerts.join("\n\n") + "\n"; }

  msg += "━━━━━━━━━━━━\n";
  var quotes = ["今天整理了一天的資料辛苦啦！早點休息喔 ✨", "把錢變成自己喜歡的樣子是一種理財，把錢好好存下來更是一種成就感。💤", "記帳是為了更自由的生活，我們每天都在往目標邁進一步！💪"];
  msg += "💬 秘書悄悄話：\n" + quotes[Math.floor(Math.random() * quotes.length)];
  sendLineMessage(msg);

  if (isLastDay) sendEndOfMonthReport();
  if (todayDate.getDay() === 0) sendWeeklyReport(); 
}

function sendLineMessage(message) {
  var channelToken = 'mCVsMtHT30VvCl4NHRhj9Igv5DUBRWWIbEbhlyWPkJhPcBX472p6dbiZlQUmhGnsR7dg5+H4pYKPFudD18crmqftMZyhMYp4H9qNPPjKrcQjkGfeARIFWt4pFxvT2373hEqs2vkCvyb7a5e7fE6PMgdB04t89/1O/w1cDnyilFU='; 
  var myUserId = 'U1cc6ad6014c9fc13924699fd3433186c'; 
  var url = 'https://api.line.me/v2/bot/message/push';
  var payload = { "to": myUserId, "messages": [{ "type": "text", "text": message }] };
  try { UrlFetchApp.fetch(url, { "method": "post", "headers": { "Content-Type": "application/json", "Authorization": "Bearer " + channelToken }, "payload": JSON.stringify(payload) }); } catch (e) {}
}

function manualTriggerReport() {
  try { sendEndOfMonthReport(); return "✅ 最新財務報表已成功發送至您的 LINE！"; } catch(e) { return "❌ 發送失敗：" + e.toString(); }
}

function sendEndOfMonthReport() {
  var today = new Date();
  var appData = getAppData();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var budgetSheet = ss.getSheetByName('AI 專用分析與預算');
  var budgetData = budgetSheet.getDataRange().getValues();
  var totalBudget = 0, overBudgetItems = [];

  for (var i = 1; i < budgetData.length; i++) {
    if (budgetData[i][0]) {
      totalBudget += (Number(budgetData[i][1]) || 0);
      var catName = budgetData[i][0];
      var catRemain = Number(budgetData[i][3]) || 0;
      if (catRemain < 0) {
        overBudgetItems.push("▪️ " + catName + " (超支 $" + Math.abs(catRemain).toLocaleString() + ")");
      }
    }
  }

  var recordData = ss.getSheetByName('紀錄頁').getDataRange().getValues();
  var categoryTotals = {}, cardTotals = {};
  var excludeCategories = ['ETF定期定額', '學貸'];
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth();

  var totalIncome = 0, totalSpent = 0, fixedSpent = 0, floatingSpent = 0;
  var fixedList = getFixedExpenses();
  var fixedItemNames = new Set(fixedList.map(function(f) { return f.item; }));

  for (var i = recordData.length - 1; i >= 1; i--) {
    var dateVal = recordData[i][0]; if (!dateVal) continue;
    var recordDate = new Date(dateVal);
    
    if (recordDate.getFullYear() === currentYear && recordDate.getMonth() === currentMonth) {
      var type = (recordData[i][1] || '').toString().trim();
      var amount = parseFloat(recordData[i][3]) || 0;
      var category = (recordData[i][4] || '').toString().trim() || '未分類';
      var accOut = (recordData[i][5] || '').toString().trim();
      var itemName = (recordData[i][2] || '').toString().trim();
      
      if (type === '收入') {
        totalIncome += amount;
      } else if (['支出', '固定支出', '訂閱', '固定'].includes(type) && excludeCategories.indexOf(category) === -1) {
        totalSpent += amount;
        categoryTotals[category] = (categoryTotals[category] || 0) + amount;
        if (accOut) cardTotals[accOut] = (cardTotals[accOut] || 0) + amount;

        if (fixedItemNames.has(itemName)) { fixedSpent += amount; } else { floatingSpent += amount; }
      }
    }
  }

  var remainingBudget = totalBudget - totalSpent;
  var statusIcon = remainingBudget >= 0 ? "✅" : "🚨";

  var actualSavings = totalIncome - totalSpent;
  var savingsRate = totalIncome > 0 ? Math.round((actualSavings / totalIncome) * 100) : 0;
  var fixedRatio = totalSpent > 0 ? Math.round((fixedSpent / totalSpent) * 100) : 0;
  var floatingRatio = totalSpent > 0 ? (100 - fixedRatio) : 0;

  var sortedCategories = Object.keys(categoryTotals).sort(function(a, b) { return categoryTotals[b] - categoryTotals[a]; });
  var sortedCards = Object.keys(cardTotals).sort(function(a, b) { return cardTotals[b] - cardTotals[a]; });
  var numEmojis = ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣'];

  var isLastDay = new Date(today.getTime() + 86400000).getDate() === 1;
  var titleMsg = isLastDay ? "📅【 " + (currentMonth + 1) + "月 期末財務總結 】📅" : "📊【 " + (currentMonth + 1) + "月 財務即時戰況 】📊";

  var msg = titleMsg + "\n━━━━━━━━━━━━\n";
  msg += "💰 淨資產：$" + appData.dashboard.netWorth.toLocaleString() + "\n";
  msg += "  (流動資產 $" + appData.dashboard.liquidAssets.toLocaleString() + " ｜ 長期投資 $" + appData.dashboard.investAssets.toLocaleString() + ")\n";
  msg += "  💡 提醒：投資帳戶為不可隨意動用之資金\n\n";
  
  msg += "📥 本月收入：$" + totalIncome.toLocaleString() + "\n";
  msg += "🏦 實際結餘：$" + actualSavings.toLocaleString() + "\n";
  msg += "   (儲蓄率 " + savingsRate + "%)\n━━━━━━━━━━━━\n";
  msg += "🔹 本月總預算：$" + totalBudget.toLocaleString() + "\n";
  msg += "🔹 本月總花費：$" + totalSpent.toLocaleString() + "\n";
  msg += statusIcon + " 預算結餘：$" + remainingBudget.toLocaleString() + "\n\n";
  
  msg += "💸 花費結構：\n固定 " + fixedRatio + "% ｜ 浮動 " + floatingRatio + "%\n━━━━━━━━━━━━\n";

  if (overBudgetItems.length > 0) msg += "🚨【 超支項目點名 】\n" + overBudgetItems.join("\n") + "\n\n";
  else msg += "🌟【 超支項目點名 】\n太棒了！皆在預算內！\n\n";

  msg += "🏆【 本月支出 Top 5 】\n";
  for (var j = 0; j < Math.min(5, sortedCategories.length); j++) {
    msg += numEmojis[j] + " " + sortedCategories[j] + "：$" + categoryTotals[sortedCategories[j]].toLocaleString() + "\n";
  }
  if (sortedCategories.length === 0) msg += "本月無支出紀錄\n";
  
  if (sortedCards.length > 0) {
    msg += "\n💳 主力帳戶：\n" + sortedCards[0] + " ($" + cardTotals[sortedCards[0]].toLocaleString() + ")\n";
  }
  msg += "━━━━━━━━━━━━\n";

  if (remainingBudget >= 0) msg += "💡 總結：本月成功守住預算 🥟 繼續保持喔！";
  else msg += "💡 總結：稍微透支了 🥲 記帳是為了找問題，下個月再抓緊！💪";

  sendLineMessage(msg);
}

function sendWeeklyReport() {
  var today = new Date();
  var appData = getAppData();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var recordData = ss.getSheetByName('紀錄頁').getDataRange().getValues();
  
  var budgetSheet = ss.getSheetByName('AI 專用分析與預算');
  var budgetData = budgetSheet.getDataRange().getValues();
  var totalBudget = 0;
  for (var i = 1; i < budgetData.length; i++) {
    if (budgetData[i][0]) totalBudget += (Number(budgetData[i][1]) || 0);
  }
  
  var sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(today.getDate() - 7);
  sevenDaysAgo.setHours(0, 0, 0, 0);
  var endOfDay = new Date();
  endOfDay.setHours(23, 59, 59, 999);
  
  var weeklySpent = 0, categoryTotals = {}, cardTotals = {};
  var excludeCategories = ['ETF定期定額', '學貸'];
  var monthTotalSpent = 0;
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth();
  
  for (var i = recordData.length - 1; i >= 1; i--) {
    var dateVal = recordData[i][0]; if (!dateVal) continue;
    var recordDate = new Date(dateVal);
    var type = (recordData[i][1] || '').toString().trim();
    var amount = parseFloat(recordData[i][3]) || 0;
    var category = (recordData[i][4] || '').toString().trim() || '未分類';
    var accOut = (recordData[i][5] || '').toString().trim();
    
    if (['支出', '固定支出', '訂閱', '固定'].includes(type) && excludeCategories.indexOf(category) === -1) {
      if (recordDate.getFullYear() === currentYear && recordDate.getMonth() === currentMonth) {
        monthTotalSpent += amount;
      }
      if (recordDate >= sevenDaysAgo && recordDate <= endOfDay) {
        weeklySpent += amount;
        categoryTotals[category] = (categoryTotals[category] || 0) + amount;
        if (accOut) cardTotals[accOut] = (cardTotals[accOut] || 0) + amount;
      }
    }
  }

  var remainingBudget = totalBudget - monthTotalSpent;
  var lastDayOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  var daysLeft = lastDayOfMonth - today.getDate();

  var fixedSheet = ss.getSheetByName('固定支出與轉帳');
  var fixedData = fixedSheet.getDataRange().getValues();
  var upcomingFixed = [];
  var upcomingFixedTotal = 0;
  
  for (var i = 1; i <= 7; i++) {
    var checkDate = new Date(today.getTime() + i * 86400000);
    var checkDay = checkDate.getDate();
    var checkLastDay = new Date(checkDate.getFullYear(), checkDate.getMonth() + 1, 0).getDate();
    
    for(var r=1; r<fixedData.length; r++) {
        if(!fixedData[r][0]) continue;
        var pDay = parseInt(fixedData[r][3]);
        if (pDay === checkDay || (checkDay === checkLastDay && pDay > checkDay)) {
            var fType = fixedData[r][1];
            if (fType === '支出' || fType === '固定支出' || fType === '訂閱') {
                var tPeriods = fixedData[r][4];
                var pPeriods = fixedData[r][5];
                if (tPeriods !== "" && pPeriods !== "" && Number(pPeriods) >= Number(tPeriods)) continue;
                
                upcomingFixed.push(fixedData[r][0]);
                upcomingFixedTotal += (Number(fixedData[r][2]) || 0);
            }
        }
    }
  }

  var sortedCategories = Object.keys(categoryTotals).sort(function(a, b) { return categoryTotals[b] - categoryTotals[a]; });
  var sortedCards = Object.keys(cardTotals).sort(function(a, b) { return cardTotals[b] - cardTotals[a]; });
  var dailyAvg = Math.round(weeklySpent / 7);
  var numEmojis = ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣'];

  var msg = "🩺【 本週花費健康檢查 】🩺\n━━━━━━━━━━━━\n";
  msg += "💰 目前淨資產：$" + appData.dashboard.netWorth.toLocaleString() + "\n";
  msg += "  (流動資產 $" + appData.dashboard.liquidAssets.toLocaleString() + " ｜ 長期投資 $" + appData.dashboard.investAssets.toLocaleString() + ")\n";
  msg += "  💡 提醒：投資帳戶為不可隨意動用之資金\n\n";

  msg += "💡 本月總預算：$" + totalBudget.toLocaleString() + "\n";
  msg += "💡 本月剩餘預算：$" + remainingBudget.toLocaleString() + "\n";
  msg += "   (距月底 " + daysLeft + " 天)\n\n";
  
  msg += "🔹 近7天花費：$" + weeklySpent.toLocaleString() + "\n";
  msg += "🔹 平均日花費：$" + dailyAvg.toLocaleString() + "\n━━━━━━━━━━━━\n";
  
  if (upcomingFixedTotal > 0) {
    var uniqueUpcoming = [...new Set(upcomingFixed)];
    var displayNames = uniqueUpcoming.slice(0, 3).join('、') + (uniqueUpcoming.length > 3 ? '...' : '');
    msg += "📅 下週自動扣款預告：\n總計 $" + upcomingFixedTotal.toLocaleString() + "\n(" + displayNames + ")\n━━━━━━━━━━━━\n";
  }
  
  msg += "🏆【 本週支出 Top 5 】\n";
  for (var j = 0; j < Math.min(5, sortedCategories.length); j++) {
    msg += numEmojis[j] + " " + sortedCategories[j] + "：$" + categoryTotals[sortedCategories[j]].toLocaleString() + "\n";
  }
  if (sortedCategories.length === 0) msg += "本週無支出紀錄\n";
  
  if (sortedCards.length > 0) {
    var topCard = sortedCards[0], topCardAmount = cardTotals[topCard];
    var cardPercent = weeklySpent > 0 ? Math.round((topCardAmount / weeklySpent) * 100) : 0;
    if (cardPercent >= 50 && topCardAmount > 2000) {
      msg += "\n⚠️【 用卡警報 】\n「" + topCard + "」佔花費 " + cardPercent + "%\n($" + topCardAmount.toLocaleString() + ")！留意集中刷卡。\n";
    }
  }
  msg += "━━━━━━━━━━━━\n";

  if (weeklySpent > 5000) msg += "💡 總結：燒錢有點快囉！迎接新的一週，稍微放慢腳步吧 💪";
  else msg += "💡 總結：花費配速很不錯！繼續保持優良理財習慣 🥟";

  sendLineMessage(msg);
}