// ── 全域狀態 ─────────────────────────────────────────────
let globalData = null, myChart = null, currentHistoryData = [], currentActionRow = null;
let currentMainType = '支出';
window.currentCatMap = {};

// ── 啟動 ─────────────────────────────────────────────────
window.onload = async function () {
  // 從 LIFF 回調讀取 LINE User ID
  const urlParams = new URLSearchParams(window.location.search);
  const lineUserIdFromUrl = urlParams.get('lineUserId');
  const lineNameFromUrl = urlParams.get('lineName');
  if (lineUserIdFromUrl) {
    localStorage.setItem('pendingLineUserId', lineUserIdFromUrl);
    if (lineNameFromUrl) localStorage.setItem('pendingLineName', lineNameFromUrl);
    window.history.replaceState({}, '', window.location.pathname);
  }

  // 註冊 Service Worker
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }

  initTheme();

  if (!Auth.isSignedIn()) {
    showLoginScreen();
    return;
  }

  await startApp();
};

async function startApp() {
  showAppScreen();
  initMonthPicker();
  initIconPreview();

  // 載入快取（立即顯示）
  const cd = localStorage.getItem('appDataCache');
  if (cd) {
    globalData = JSON.parse(cd);
    renderDashboard(globalData);
    buildCategoryGrid(globalData.categories);
    applySettingsToUI(globalData.settings);
  }
  const m = document.getElementById('monthPicker').value;
  const ch = localStorage.getItem('historyCache_' + m);
  if (ch) { currentHistoryData = JSON.parse(ch); renderHistoryList(currentHistoryData); drawChart(currentHistoryData); }

  // 確認試算表存在後同步最新資料
  try {
    await Sheets.initSpreadsheet();
    manualUpdate();
  } catch(e) {
    showAlert('初始化試算表失敗：' + e.message);
  }
}

// ── 登入 / 登出 UI ────────────────────────────────────────
function showLoginScreen() {
  document.getElementById('loginScreen').style.display = 'flex';
  document.getElementById('appWrapper').style.display = 'none';
}

function showAppScreen() {
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('appWrapper').style.display = 'block';
  const info = Auth.getUserInfo();
  if (info) {
    document.getElementById('userAvatar').src = info.picture || '';
    document.getElementById('userName').innerText = info.name || '';
    document.getElementById('userEmail').innerText = info.email || '';
  }
}

async function handleSignIn() {
  try {
    document.getElementById('signInBtn').innerText = '登入中...';
    document.getElementById('signInBtn').disabled = true;
    await Auth.getToken();
    await Sheets.initSpreadsheet();
    await startApp();
  } catch(e) {
    document.getElementById('signInBtn').innerText = '使用 Google 帳號登入';
    document.getElementById('signInBtn').disabled = false;
    showAlert('登入失敗：' + e.message);
  }
}

function handleSignOut() {
  showConfirm('確定要登出嗎？', () => Auth.signOut());
}

// ── 同步 ─────────────────────────────────────────────────
function applySyncData(res, m) {
  if (!res) return;
  globalData = res.appData;
  localStorage.setItem('appDataCache', JSON.stringify(res.appData));
  renderDashboard(res.appData);
  buildCategoryGrid(res.appData.categories);
  applySettingsToUI(res.appData.settings);
  currentHistoryData = res.history;
  localStorage.setItem('historyCache_' + m, JSON.stringify(res.history));
  renderHistoryList(res.history);
  drawChart(res.history);
  document.getElementById('syncBtn').classList.remove('syncing');

  // LIFF 回調：自動儲存 LINE User ID
  const pendingLineId = localStorage.getItem('pendingLineUserId');
  if (pendingLineId) {
    localStorage.removeItem('pendingLineUserId');
    document.getElementById('setting_lineUserId').value = pendingLineId;
    saveSettings().then(() => showAlert('✅ LINE 帳號已成功連結！'));
  }
}

function manualUpdate() {
  const btn = document.getElementById('syncBtn');
  if (btn.classList.contains('syncing')) return;
  btn.classList.add('syncing');
  const m = document.getElementById('monthPicker').value;
  Sheets.getSyncData(m)
    .then(res => applySyncData(res, m))
    .catch(e => { btn.classList.remove('syncing'); showAlert('同步失敗：' + e.message); });
}

// ── 設定套用 ──────────────────────────────────────────────
function applySettingsToUI(settings) {
  if (!settings) return;
  document.getElementById('ui-invest-label').innerText = settings.investLabel || '長期投資';
  const fields = ['appName','investKeyword','investLabel','excludeCategories',
                  'alertThreshold','alertIgnoreWords','lineUserId'];
  fields.forEach(f => {
    const el = document.getElementById('setting_' + f);
    if (el) el.value = settings[f] || '';
  });
  updateLineStatus(settings.lineUserId);
}

function updateLineStatus(userId) {
  const connectedEl = document.getElementById('lineConnectedStatus');
  const disconnectedEl = document.getElementById('lineDisconnectedStatus');
  const nameEl = document.getElementById('lineConnectedName');
  if (userId) {
    const name = localStorage.getItem('pendingLineName') || '';
    if (connectedEl) connectedEl.style.display = 'block';
    if (disconnectedEl) disconnectedEl.style.display = 'none';
    if (nameEl) nameEl.innerText = name ? `帳號：${name}` : `ID：${userId.substring(0, 12)}…`;
  } else {
    if (connectedEl) connectedEl.style.display = 'none';
    if (disconnectedEl) disconnectedEl.style.display = 'block';
  }
}

function connectLine() {
  location.href = 'https://liff.line.me/' + CONFIG.LIFF_ID;
}

function disconnectLine() {
  showConfirm('確定要取消 LINE 連結嗎？', async () => {
    document.getElementById('setting_lineUserId').value = '';
    localStorage.removeItem('pendingLineName');
    await saveSettings();
    updateLineStatus('');
    showAlert('✅ 已取消 LINE 連結');
  });
}

async function saveSettings() {
  const btn = document.getElementById('saveSettingsBtn');
  btn.innerText = '儲存中...'; btn.disabled = true;
  const fields = ['appName','investKeyword','investLabel','excludeCategories',
                  'alertThreshold','alertIgnoreWords','lineUserId'];
  const settings = { ...CONFIG.DEFAULT_SETTINGS };
  fields.forEach(f => {
    const el = document.getElementById('setting_' + f);
    if (el) settings[f] = el.value;
  });
  try {
    await Sheets.saveSettings(settings);
    showAlert('✅ 設定已儲存！');
    manualUpdate();
  } catch(e) { showAlert('❌ 儲存失敗：' + e.message); }
  btn.innerText = '儲存設定'; btn.disabled = false;
}

// ── 儀表板 ────────────────────────────────────────────────
function renderDashboard(d) {
  document.getElementById('ui-networth').innerText = '$ ' + d.dashboard.netWorth.toLocaleString();
  document.getElementById('ui-liquid').innerText   = '$ ' + d.dashboard.liquidAssets.toLocaleString();
  document.getElementById('ui-invest').innerText   = '$ ' + d.dashboard.investAssets.toLocaleString();
  document.getElementById('ui-liabilities').innerText = '$ ' + d.dashboard.totalLiabilities.toLocaleString();

  let accs = '<option value="">請選擇帳戶...</option>'
    + '<optgroup label="銀行">' + d.banks.map(b => `<option value="${b}">${b}</option>`).join('') + '</optgroup>'
    + '<optgroup label="卡片">' + d.cards.map(c => `<option value="${c}">${c}</option>`).join('') + '</optgroup>'
    + '<optgroup label="管理選項"><option value="ADD_BANK" style="color:var(--primary);font-weight:bold;">[ ＋ 新增銀行... ]</option><option value="ADD_CARD" style="color:var(--primary);font-weight:bold;">[ ＋ 新增信用卡... ]</option></optgroup>';
  document.getElementById('accountOut').innerHTML = accs;
  document.getElementById('accountIn').innerHTML  = accs;

  document.getElementById('dashContent').innerHTML =
    `<div class="card"><div class="dash-title">🏦 銀行</div>
      ${d.dashboard.banks.map(b =>
        `<div class="dash-item" onclick="openDetails('${b.name}')">
          <div style="display:flex;flex-direction:column;align-items:flex-start;">
            <span>${b.name}</span>
            <span style="font-size:11px;color:#888;margin-top:4px;cursor:pointer;" onclick="event.stopPropagation();editBankNote('${b.name}','${(b.note||'').replace(/'/g,"\\'")}')">
              ${b.note || '點擊新增備註'}
            </span>
          </div>
          <span class="val-pos">$ ${b.balance.toLocaleString()}</span>
        </div>`).join('')}
    </div>
    <div class="card"><div class="dash-title">💳 信用卡</div>
      ${d.dashboard.cards.map(c =>
        `<div class="dash-item" style="flex-direction:column;align-items:stretch;" onclick="openDetails('${c.name}')">
          <div style="display:flex;justify-content:space-between;width:100%;">
            <span>${c.name}</span>
            <span class="${c.currentBill>0?'val-neg':'val-pos'}">${c.currentBill>0?'$ '+c.currentBill.toLocaleString():'已繳清'}</span>
          </div>
          <div style="font-size:11px;color:#999;margin-top:5px;">本月應付: ${c.totalDebt.toLocaleString()} / 剩餘: ${c.remain.toLocaleString()}</div>
        </div>`).join('')}
    </div>`;
}

// ── 搜尋 ─────────────────────────────────────────────────
function filterHistory() {
  const k = document.getElementById('searchInput').value.toLowerCase();
  let exp = 0, inc = 0, count = 0;
  document.querySelectorAll('#historyList .detail-row').forEach(r => {
    const match = r.innerText.toLowerCase().includes(k);
    r.style.display = match ? 'flex' : 'none';
  });
  if (!k) { document.getElementById('searchSummary').style.display = 'none'; return; }
  currentHistoryData.forEach(d => {
    if (`${d.item} ${d.category} ${d.accOut} ${d.accIn} ${d.amount} ${d.date}`.toLowerCase().includes(k)) {
      count++;
      if (['支出','繳卡費','固定支出','訂閱','借出'].includes(d.type) || d.type==='點數折抵') exp += d.amount;
      else if (['收入','收回借款','收回'].includes(d.type)) inc += d.amount;
    }
  });
  let html = `<div style="font-weight:bold;color:var(--text);margin-bottom:4px;">🔍 找到 ${count} 筆相符紀錄</div>`;
  if (exp || inc) {
    html += `<div style="display:flex;justify-content:center;gap:15px;margin-top:6px;">`;
    if (exp) html += `<span>支出：<span class="val-neg">$ ${exp.toLocaleString()}</span></span>`;
    if (inc) html += `<span>收入：<span class="val-pos">$ ${inc.toLocaleString()}</span></span>`;
    html += `</div>`;
  }
  document.getElementById('searchSummary').innerHTML = html;
  document.getElementById('searchSummary').style.display = 'block';
}

// ── 分類圖示 ──────────────────────────────────────────────
const iconMap = { '餐飲':'🍔','交通':'🚗','居住':'🏠','娛樂':'🎬','購物':'🛍️','生活':'🛒','醫療':'🏥','教育':'📚','進修':'📖','保險':'🛡️','薪水':'💰','投資':'📈','手續費':'🏧','水電':'💧','瓦斯':'🔥','電信':'📱','網路':'🌐' };

function parseCategory(c) {
  const emojiRegex = /[\p{Emoji_Presentation}\p{Extended_Pictographic}]/u;
  const match = c.match(emojiRegex);
  if (match) return { icon: match[0], name: c.replace(match[0], '').trim() };
  for (let key in iconMap) { if (c.includes(key)) return { icon: iconMap[key], name: c.trim() }; }
  return { icon: '🏷️', name: c.trim() };
}

function buildCategoryGrid(cats) {
  document.getElementById('categoryGrid').innerHTML =
    cats.map(c => {
      const p = parseCategory(c);
      return `<div class="cat-item" data-cat="${c}" onclick="selectGridCategory('${c.replace(/'/g,"\\'")}',this)">
        <div class="cat-icon">${p.icon}</div><div class="cat-name">${p.name}</div>
      </div>`;
    }).join('') +
    `<div class="cat-item" style="border:1px dashed var(--primary);color:var(--primary);" onclick="addGridCategory()">
      <div class="cat-icon">➕</div><div class="cat-name">新增分類</div>
    </div>`;
}

function selectGridCategory(v, el) {
  document.getElementById('category').value = v;
  document.querySelectorAll('.cat-item').forEach(i => i.classList.remove('active'));
  if (el) el.classList.add('active');
}

function addGridCategory() {
  openInputModal('新增分類名稱', name => {
    if (!name) return;
    document.getElementById('syncBtn').classList.add('syncing');
    const m = document.getElementById('monthPicker').value;
    Sheets.addCategory(name, m).then(res => { showAlert(res.msg); applySyncData(res.syncData, m); });
  });
}

// ── 帳戶選單 ──────────────────────────────────────────────
let activeSelectEl = null;
function checkNewAccount(sel) {
  if (sel.value.startsWith('ADD_')) {
    activeSelectEl = sel;
    sel.value === 'ADD_BANK' ? openBankModal() : openCcModal();
  }
}

function editBankNote(name, currentNote) {
  openInputModal(`編輯 ${name} 的備註`, newNote => {
    document.getElementById('syncBtn').classList.add('syncing');
    const m = document.getElementById('monthPicker').value;
    Sheets.updateBankNote(name, newNote, m).then(res => { showAlert(res.msg); applySyncData(res.syncData, m); });
  }, currentNote);
}

function openBankModal() {
  ['bankName','bankInitial','bankNote'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('bankModal').style.display = 'flex';
  setTimeout(() => document.getElementById('bankName').focus(), 100);
}
function closeBankModal() { document.getElementById('bankModal').style.display='none'; if(activeSelectEl?.value==='ADD_BANK') activeSelectEl.value=''; }
function submitBankModal() {
  const name = document.getElementById('bankName').value.trim();
  const initial = document.getElementById('bankInitial').value;
  const note = document.getElementById('bankNote').value.trim();
  if (!name || !initial) return showAlert('請填寫銀行名稱與初始金額！');
  document.getElementById('bankModal').style.display = 'none';
  if (activeSelectEl) { activeSelectEl.disabled = true; }
  document.getElementById('syncBtn').classList.add('syncing');
  const m = document.getElementById('monthPicker').value;
  Sheets.addBankAccount({ name, initial, note }, m).then(res => {
    if (activeSelectEl) activeSelectEl.disabled = false;
    showAlert(res.msg); applySyncData(res.syncData, m);
  });
}

function openCcModal() {
  ['ccName','ccLimit','ccBillDate','ccPayDate'].forEach(id => document.getElementById(id).value = '');
  let opts = '<option value="">(不指定自動扣款帳戶)</option>';
  if (globalData?.banks) globalData.banks.forEach(b => { opts += `<option value="${b}">${b}</option>`; });
  document.getElementById('ccAutoPayAccount').innerHTML = opts;
  document.getElementById('ccModal').style.display = 'flex';
}
function closeCcModal() { document.getElementById('ccModal').style.display='none'; if(activeSelectEl?.value==='ADD_CARD') activeSelectEl.value=''; }
function submitCcModal() {
  const name=document.getElementById('ccName').value.trim(), limit=document.getElementById('ccLimit').value;
  const billDate=document.getElementById('ccBillDate').value, payDate=document.getElementById('ccPayDate').value;
  const autoPayAcc=document.getElementById('ccAutoPayAccount').value;
  if (!name||!limit||!billDate||!payDate) return showAlert('請填寫完整資訊！');
  document.getElementById('ccModal').style.display='none';
  if (activeSelectEl) activeSelectEl.disabled = true;
  document.getElementById('syncBtn').classList.add('syncing');
  const m = document.getElementById('monthPicker').value;
  Sheets.addCreditCard({ name, limit, billDate, payDate, autoPayAcc }, m).then(res => {
    if (activeSelectEl) activeSelectEl.disabled = false;
    showAlert(res.msg); applySyncData(res.syncData, m);
  });
}

// ── 新增記錄 Modal ────────────────────────────────────────
function openRecordModal() {
  document.getElementById('recordModalTitle').innerText = '📝 新增紀錄';
  document.getElementById('editRow').value = '';
  ['amount','item','totalPeriods','paidPeriods'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('recordDate').value = new Date().toISOString().split('T')[0];
  document.getElementById('btnFixedExp').style.display = 'block';
  document.getElementById('btnSub').style.display = 'block';
  document.getElementById('templateWrapper').style.display = 'block';
  selectGridCategory('', null);
  selectMainType('支出', document.querySelector('#mainTypeScroll .type-btn'));
  renderTemplates();
  document.getElementById('recordModal').style.display = 'flex';
}

function openEditModal(row) {
  const d = currentHistoryData.find(x => x.row === row);
  if (!d) return;
  document.getElementById('recordModalTitle').innerText = '✏️ 編輯紀錄';
  document.getElementById('editRow').value = row;
  document.getElementById('recordDate').value = d.fullDate;
  document.getElementById('amount').value = d.type==='點數折抵' ? Math.abs(d.amount) : d.amount;
  document.getElementById('item').value = d.item;
  document.getElementById('btnFixedExp').style.display = 'none';
  document.getElementById('btnSub').style.display = 'none';
  document.getElementById('templateWrapper').style.display = 'none';
  document.querySelectorAll('.cat-item').forEach(el => { if(el.getAttribute('data-cat')===d.category) selectGridCategory(d.category,el); });
  let baseType = d.type;
  if (baseType==='收回借款') baseType='收回';
  if (['固定支出','訂閱','固定'].includes(baseType)) baseType='支出';
  document.querySelectorAll('#recordForm .type-btn').forEach(b => { if(b.innerText.includes(baseType)) selectMainType(baseType,b); });
  setTimeout(() => { document.getElementById('accountOut').value=d.accOut; document.getElementById('accountIn').value=d.accIn; }, 100);
  document.getElementById('recordModal').style.display = 'flex';
}

function closeRecordModal() { document.getElementById('recordModal').style.display = 'none'; }

function selectMainType(v, b) {
  document.querySelectorAll('#recordForm .type-btn').forEach(x => x.classList.remove('active'));
  if (b) b.classList.add('active');
  currentMainType = v;
  if (v==='手續費') {
    document.querySelectorAll('.cat-item').forEach(el => {
      if (el.getAttribute('data-cat')?.includes('手續費')) selectGridCategory(el.getAttribute('data-cat'), el);
    });
  }
  updateFormUI();
}

function updateFormUI() {
  const outEl=document.getElementById('accountOut'), inEl=document.getElementById('accountIn');
  const hintEl=document.getElementById('dateHint'), btn=document.getElementById('submitBtn');
  const fixedTypeWrapper=document.getElementById('fixedTypeWrapper'), periodsWrapper=document.getElementById('periodsWrapper');
  const typeInput=document.getElementById('type'), isEdit=document.getElementById('editRow').value!=='';
  outEl.style.display='block'; inEl.style.display='none'; fixedTypeWrapper.style.display='none';
  periodsWrapper.style.display='none'; hintEl.style.display='none';
  btn.innerText = isEdit ? '💾 儲存修改' : '送出紀錄';
  let actualType = currentMainType;
  if (currentMainType==='支出') { actualType='支出'; }
  else if (currentMainType==='收入') { actualType='收入'; outEl.style.display='none'; inEl.style.display='block'; }
  else if (currentMainType==='轉帳'||currentMainType==='繳卡費') { actualType='轉帳'; inEl.style.display='block'; }
  else if (currentMainType==='借出') { actualType='借出'; }
  else if (currentMainType==='收回') { actualType='收回借款'; outEl.style.display='none'; inEl.style.display='block'; }
  else if (currentMainType==='訂閱') { actualType='訂閱'; if(!isEdit){ hintEl.style.display='block'; periodsWrapper.style.display='flex'; btn.innerText='🗓️ 加入自動扣款排程'; } }
  else if (currentMainType==='固定支出') {
    if (!isEdit) {
      fixedTypeWrapper.style.display='block'; hintEl.style.display='block';
      periodsWrapper.style.display='flex'; btn.innerText='🗓️ 加入自動扣款排程';
      const subType = document.querySelector('input[name="fixedType"]:checked').value;
      actualType = subType;
      if (subType==='轉帳') inEl.style.display='block';
    } else { actualType='支出'; }
  }
  else if (currentMainType==='點數折抵') { actualType='點數折抵'; }
  else if (currentMainType==='手續費') { actualType='支出'; }
  typeInput.value = actualType;
  if (outEl.style.display==='none') outEl.value='';
  if (inEl.style.display==='none') inEl.value='';
}

// ── 表單送出 ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('recordForm').onsubmit = function(e) {
    e.preventDefault();
    const b = document.getElementById('submitBtn');
    b.disabled = true; b.innerText = '處理中...';
    const rawAmount = Number(document.getElementById('amount').value);
    const finalType = document.getElementById('type').value;
    const finalAmount = currentMainType==='點數折抵' ? -Math.abs(rawAmount) : rawAmount;
    const fd = {
      row: document.getElementById('editRow').value,
      date: document.getElementById('recordDate').value,
      type: finalType,
      isFixed: ['固定支出','訂閱'].includes(currentMainType) && !document.getElementById('editRow').value,
      totalPeriods: document.getElementById('totalPeriods').value,
      paidPeriods: document.getElementById('paidPeriods').value,
      amount: finalAmount,
      item: document.getElementById('item').value,
      category: document.getElementById('category').value,
      accountOut: document.getElementById('accountOut').value,
      accountIn: document.getElementById('accountIn').value
    };
    const m = document.getElementById('monthPicker').value;
    document.getElementById('syncBtn').classList.add('syncing');
    const action = fd.row ? Sheets.updateRecordData(fd, m) : Sheets.recordData(fd, m);
    action.then(res => {
      b.disabled=false; b.innerText='送出紀錄';
      closeRecordModal(); showAlert(res.msg); applySyncData(res.syncData, m);
    }).catch(err => {
      b.disabled=false; b.innerText='送出紀錄';
      showAlert('❌ 發生錯誤：' + err.message);
    });
  };
});

// ── 操作選單 ──────────────────────────────────────────────
function openActionMenu(r)  { currentActionRow=r; document.getElementById('actionModal').style.display='flex'; }
function closeActionMenu()  { document.getElementById('actionModal').style.display='none'; currentActionRow=null; }
function handleEditAction() { if(currentActionRow) openEditModal(currentActionRow); closeActionMenu(); }
function handleDeleteAction() {
  if (currentActionRow) {
    const targetRow = currentActionRow;
    showConfirm('確定要刪除這筆紀錄嗎？', () => {
      document.getElementById('syncBtn').classList.add('syncing');
      const m = document.getElementById('monthPicker').value;
      Sheets.deleteRecord(targetRow, m).then(res => { showAlert(res.msg); applySyncData(res.syncData, m); });
    });
  }
  closeActionMenu();
}

// ── 詳情 Modal ────────────────────────────────────────────
function openDetails(n) {
  document.getElementById('detailModal').style.display='flex';
  document.getElementById('detailTitle').innerText = n;
  document.getElementById('detailList').innerHTML = '<div style="text-align:center;padding:20px;color:#888;">讀取中...</div>';
  Sheets.getAccountDetails(n).then(ds => {
    document.getElementById('detailList').innerHTML = ds.length===0
      ? '<div style="text-align:center;padding:20px;color:#888;">無紀錄</div>'
      : ds.map(d => `<div class="detail-row"><div><b>${d.item}</b><br><small style="color:#999">${d.date}</small></div><div class="${d.effectClass}" style="font-weight:bold;">${d.effect}</div></div>`).join('');
  });
}

function viewFixedExpenses() {
  document.getElementById('detailModal').style.display='flex';
  document.getElementById('detailTitle').innerText='自動扣款排程';
  document.getElementById('detailList').innerHTML='<div style="text-align:center;padding:20px;color:#888;">讀取中...</div>';
  Sheets.getFixedExpenses().then(list => {
    if (list.length===0) { document.getElementById('detailList').innerHTML='<div style="text-align:center;padding:20px;color:#888;">目前無排程</div>'; return; }
    let total=0, groups={};
    list.forEach(f => { const a=f.account||'未指定'; if(!groups[a]) groups[a]={total:0,items:[]}; groups[a].items.push(f); groups[a].total+=f.amount; if(f.category) total+=f.amount; });
    let html=`<div style="font-weight:bold;font-size:16px;margin-bottom:15px;text-align:center;padding-bottom:15px;border-bottom:1px solid var(--border);">每月總支出: <span class="val-neg">$ ${total.toLocaleString()}</span></div>`;
    const colors=[{bg:'#e3f2fd',b:'#2196f3',t:'#0d47a1'},{bg:'#fce4ec',b:'#f44336',t:'#b71c1c'},{bg:'#e8f5e9',b:'#4caf50',t:'#1b5e20'},{bg:'#fff3e0',b:'#ff9800',t:'#e65100'}];
    let ci=0;
    for (let a in groups) {
      const c=colors[ci%colors.length], dark=document.body.classList.contains('dark-mode');
      html+=`<div style="background:${dark?'#2c2c2c':c.bg};border-left:5px solid ${c.b};padding:10px 15px;margin-top:15px;font-size:14px;font-weight:bold;color:${dark?'#e0e0e0':c.t};border-radius:4px 8px 8px 4px;display:flex;justify-content:space-between;"><span>🏦 ${a}</span><span style="color:var(--danger);">$ ${groups[a].total.toLocaleString()}</span></div>`;
      html+=groups[a].items.map(f=>`<div class="detail-row" style="padding:12px 5px;margin-left:5px;"><div><b>${f.item}</b><br><small style="color:#999">${f.day}號扣款 (${f.type})</small></div><div style="text-align:right;"><span class="val-neg">$ ${f.amount.toLocaleString()}</span><br><small style="color:#bbb">${f.paid}/${f.total} 期</small></div></div>`).join('');
      ci++;
    }
    document.getElementById('detailList').innerHTML=html;
  });
}

function closeDetails() { document.getElementById('detailModal').style.display='none'; }

// ── 分類明細 ──────────────────────────────────────────────
function showCategoryRecords(label) {
  const records = window.currentCatMap[label] || [];
  document.getElementById('detailModal').style.display='flex';
  document.getElementById('detailTitle').innerText = label + ' 明細';
  if (!records.length) { document.getElementById('detailList').innerHTML='<div style="text-align:center;padding:20px;color:#888;">無紀錄</div>'; return; }
  const sorted = [...records].sort((a,b) => new Date(b.fullDate)-new Date(a.fullDate));
  document.getElementById('detailList').innerHTML = sorted.map(d => {
    const vc = d.type==='點數折抵' ? 'val-pos' : 'val-neg';
    return `<div class="detail-row" style="padding:12px 0;">
      <div style="flex:1;min-width:0;padding-right:12px;">
        <div style="margin-bottom:4px;"><b style="font-size:15px;color:var(--text);">${d.item}</b></div>
        <div style="font-size:12px;color:#999;">${d.date} · ${d.accOut||''}${d.accIn?' ➔ '+d.accIn:''}</div>
      </div>
      <div style="display:flex;align-items:center;gap:12px;">
        <span class="${vc}" style="font-weight:bold;font-size:15px;">$ ${d.amount.toLocaleString()}</span>
        <div onclick="openActionMenu(${d.row});document.getElementById('detailModal').style.display='none';" style="padding:10px 5px;color:#aaa;cursor:pointer;font-size:20px;line-height:1;margin-right:-5px;font-weight:bold;">⋮</div>
      </div>
    </div>`;
  }).join('');
}

// ── 圖表 ─────────────────────────────────────────────────
function drawChart(h) {
  const settings = globalData?.settings || {};
  const excluded = (settings.excludeCategories || '').split(',').map(s => s.trim()).filter(Boolean);
  let ct = {}, tt = 0, catMap = {};
  h.forEach(d => {
    if (['支出','固定支出','訂閱','點數折抵'].includes(d.type) && !excluded.includes(d.category)) {
      const p = parseCategory(d.category);
      const key = `${p.icon} ${p.name}`;
      ct[key] = (ct[key]||0) + d.amount; tt += d.amount;
      if (!catMap[key]) catMap[key] = [];
      catMap[key].push(d);
    }
  });
  if (tt <= 0) { document.getElementById('chartCard').style.display='none'; return; }
  document.getElementById('chartCard').style.display='block';
  const sorted = Object.keys(ct).filter(k=>ct[k]>0).sort((a,b)=>ct[b]-ct[a]);
  let topCats={}, otherTotal=0, otherRecords=[];
  sorted.forEach((cat,i) => {
    if (i<5) topCats[cat]=ct[cat];
    else { otherTotal+=ct[cat]; otherRecords.push(...catMap[cat]); }
  });
  if (otherTotal>0) { topCats['📦 其他']=otherTotal; catMap['📦 其他']=otherRecords; }
  window.currentCatMap = catMap;
  const labels=Object.keys(topCats), data=Object.values(topCats);
  const colors=['#FF3B30','#FF9500','#FFCC00','#34C759','#5AC8FA','#AF52DE'];
  const isDark=document.body.classList.contains('dark-mode');
  document.getElementById('chartLegend').innerHTML = labels.map((label,i) => {
    const pct=Math.round((data[i]/tt)*100);
    return `<div class="clickable-legend" onclick="showCategoryRecords('${label}')" style="display:flex;justify-content:space-between;align-items:center;font-size:14px;">
      <div style="display:flex;align-items:center;gap:12px;"><span style="display:inline-block;width:12px;height:12px;border-radius:50%;background-color:${colors[i%colors.length]};"></span><span style="color:var(--text);font-weight:600;">${label}</span></div>
      <div style="display:flex;align-items:center;gap:15px;"><span style="color:#888;width:40px;text-align:right;font-size:13px;">${pct}%</span><span style="font-weight:bold;color:var(--text);width:70px;text-align:right;">$${data[i].toLocaleString()}</span></div>
    </div>`;
  }).join('');
  document.getElementById('chartCenterTotal').innerHTML = `<div style="font-size:12px;color:#888;margin-bottom:4px;">總花費</div><div style="font-size:24px;font-weight:bold;color:var(--text);line-height:1;">$${tt.toLocaleString()}</div>`;
  if (myChart) myChart.destroy();
  myChart = new Chart(document.getElementById('expenseChart').getContext('2d'), {
    type:'doughnut',
    data:{ labels, datasets:[{ data, backgroundColor:colors, borderWidth:4, borderColor:isDark?'#1e1e1e':'#ffffff', hoverOffset:4 }] },
    options:{
      responsive:true, maintainAspectRatio:false, cutout:'75%',
      plugins:{ legend:{display:false}, tooltip:{enabled:true,callbacks:{label:c=>' $'+c.parsed.toLocaleString()}} },
      onHover:(e,els)=>{ e.native.target.style.cursor=els[0]?'pointer':'default'; },
      onClick:(e,els)=>{ if(els.length>0) showCategoryRecords(labels[els[0].index]); }
    }
  });
}

// ── 交易記錄列表 ──────────────────────────────────────────
function renderHistoryList(h) {
  document.getElementById('historyList').innerHTML = h.length===0
    ? '<div style="text-align:center;padding:20px;color:#888;">無紀錄</div>'
    : h.map(d => {
        const p = parseCategory(d.category);
        const vc = (d.type==='點數折抵'||['收入','收回借款','收回'].includes(d.type)) ? 'val-pos' : 'val-neg';
        return `<div class="detail-row" style="align-items:center;">
          <div style="flex:1;min-width:0;padding-right:12px;">
            <div style="display:flex;align-items:center;flex-wrap:wrap;gap:6px;margin-bottom:6px;">
              <b style="font-size:15px;color:var(--text);">${d.item}</b>
              <small style="color:#888;background:var(--input-bg);border:1px solid var(--border);padding:2px 6px;border-radius:6px;font-size:11px;white-space:nowrap;">${p.icon} ${p.name}</small>
            </div>
            <div style="font-size:12px;color:#999;line-height:1.4;">${d.date} · ${d.accOut||''}${d.accIn?' ➔ '+d.accIn:''}</div>
          </div>
          <div style="display:flex;align-items:center;flex-shrink:0;gap:12px;">
            <span class="${vc}" style="font-weight:bold;font-size:16px;">$ ${d.amount.toLocaleString()}</span>
            <div onclick="openActionMenu(${d.row})" style="padding:10px 5px;color:#aaa;cursor:pointer;font-size:20px;line-height:1;margin-right:-5px;font-weight:bold;">⋮</div>
          </div>
        </div>`;
      }).join('');
}

// ── 模板 ─────────────────────────────────────────────────
function renderTemplates() {
  const t = JSON.parse(localStorage.getItem('appTemplates')||'[]');
  document.getElementById('templateList').innerHTML = t.length
    ? t.map((x,i) => { const p=parseCategory(x.category); return `<div style="display:inline-flex;align-items:center;background:var(--input-bg);border:1px solid var(--border);border-radius:10px;padding:8px 12px;cursor:pointer;" onclick="applyTemplate(${i})"><span style="margin-right:6px;">${p.icon}</span><b>${x.item}</b><small style="margin-left:5px;color:#888;">$${x.amount}</small></div>`; }).join('')
    : '<span style="color:#aaa;font-size:12px;padding:5px 0;">尚無模板，填妥資料後點擊右上角新增</span>';
}

function saveCurrentAsTemplate() {
  const i=document.getElementById('item').value.trim(), a=document.getElementById('amount').value, c=document.getElementById('category').value;
  if (!i||!a||!c) return showAlert('請填寫完整內容再儲存模板！');
  const t=JSON.parse(localStorage.getItem('appTemplates')||'[]');
  t.push({ item:i, amount:a, category:c, type:document.getElementById('type').value, accOut:document.getElementById('accountOut').value, accIn:document.getElementById('accountIn').value });
  localStorage.setItem('appTemplates', JSON.stringify(t));
  renderTemplates(); showAlert('✅ 模板已儲存！');
}

function applyTemplate(idx) {
  const t=JSON.parse(localStorage.getItem('appTemplates')||'[]')[idx]; if(!t) return;
  document.querySelectorAll('#recordForm .type-btn').forEach(b => { if(b.innerText.includes(t.type)) selectMainType(t.type,b); });
  document.getElementById('amount').value=t.amount; document.getElementById('item').value=t.item;
  document.querySelectorAll('.cat-item').forEach(el => { if(el.getAttribute('data-cat')===t.category) selectGridCategory(t.category,el); });
  setTimeout(() => { document.getElementById('accountOut').value=t.accOut||''; document.getElementById('accountIn').value=t.accIn||''; }, 50);
}

// ── 頁面切換 ──────────────────────────────────────────────
function switchPage(p, n) {
  document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(x=>x.classList.remove('active'));
  document.getElementById('page-'+p).classList.add('active'); n.classList.add('active');
  const titles={history:'交易紀錄',dashboard:'資產總覽',settings:'系統設定'};
  document.getElementById('headerTitle').innerText=titles[p];
}

// ── 主題 ─────────────────────────────────────────────────
function initTheme() {
  const s=localStorage.getItem('appTheme')||'auto'; setTheme(s);
  document.querySelectorAll('.theme-opt').forEach(el=>{
    if((s==='light'&&el.innerText.includes('亮色'))||(s==='dark'&&el.innerText.includes('深色'))||(s==='auto'&&el.innerText.includes('跟隨'))) el.classList.add('active');
  });
}

function setTheme(m, b) {
  if(b){ document.querySelectorAll('.theme-opt').forEach(x=>x.classList.remove('active')); b.classList.add('active'); }
  localStorage.setItem('appTheme',m);
  const dark=m==='dark'||(m==='auto'&&window.matchMedia('(prefers-color-scheme: dark)').matches);
  document.body.classList.toggle('dark-mode',dark);
  if(myChart?.data.datasets[0]){ myChart.data.datasets[0].borderColor=dark?'#1e1e1e':'#ffffff'; myChart.update(); }
}

// ── 工具函式 ──────────────────────────────────────────────
function initMonthPicker() {
  const d=new Date(); document.getElementById('monthPicker').value=`${d.getFullYear()}-${(d.getMonth()+1).toString().padStart(2,'0')}`;
}

function sendReportNow(b) {
  b.innerText='⏳ 發送中...'; b.disabled=true;
  Sheets.sendLineMessage('📊 包子記帳財務報表（手動觸發）')
    .then(() => showAlert('✅ 已發送至 LINE！\n請確認有填入 LINE User ID 並儲存設定。'))
    .catch(e => showAlert('❌ 發送失敗：' + e.message))
    .finally(() => { b.innerText='立即發送'; b.disabled=false; });
}

function openSheet() {
  const id = Sheets.sid();
  if (id) window.open(`https://docs.google.com/spreadsheets/d/${id}`, '_blank');
}

function clearCache() {
  const theme=localStorage.getItem('appTheme'), tpls=localStorage.getItem('appTemplates');
  localStorage.clear();
  if(theme) localStorage.setItem('appTheme',theme);
  if(tpls) localStorage.setItem('appTemplates',tpls);
  showAlert('快取已清除，請重新整理頁面。');
}

// ── 桌面圖示 ──────────────────────────────────────────────
function handleIconUpload(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const img = new Image();
    img.onload = function() {
      const icon192 = _resizeImage(img, 192);
      const icon512 = _resizeImage(img, 512);
      localStorage.setItem('customIcon192', icon192);
      localStorage.setItem('customIcon512', icon512);
      updateDynamicManifest();
      _showIconPreview(icon192);
      showAlert('✅ 圖示已更新！\n請重新將 App 加入桌面以套用新圖示。');
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function _resizeImage(img, size) {
  const canvas = document.createElement('canvas');
  canvas.width = size; canvas.height = size;
  const ctx = canvas.getContext('2d');
  // 置中裁切為正方形
  const min = Math.min(img.width, img.height);
  const sx = (img.width - min) / 2;
  const sy = (img.height - min) / 2;
  ctx.drawImage(img, sx, sy, min, min, 0, 0, size, size);
  return canvas.toDataURL('image/png');
}

function _showIconPreview(icon192) {
  const box = document.getElementById('iconPreviewBox');
  const status = document.getElementById('iconStatusText');
  if (icon192) {
    box.innerHTML = `<img src="${icon192}" style="width:100%;height:100%;object-fit:cover;">`;
    if (status) status.innerText = '使用自訂圖示';
  } else {
    box.innerHTML = '🥟';
    if (status) status.innerText = '使用預設圖示';
  }
}

function resetIcon() {
  localStorage.removeItem('customIcon192');
  localStorage.removeItem('customIcon512');
  updateDynamicManifest();
  _showIconPreview(null);
  showAlert('✅ 已恢復預設圖示。\n請重新將 App 加入桌面以套用。');
}

function updateDynamicManifest() {
  const icon192 = localStorage.getItem('customIcon192');
  const icon512 = localStorage.getItem('customIcon512');
  const appName = globalData?.settings?.appName || '包子記帳';
  const manifest = {
    name: appName, short_name: appName,
    start_url: './', display: 'standalone',
    background_color: '#ffffff', theme_color: '#000000',
    icons: icon192
      ? [{ src: icon192, sizes: '192x192', type: 'image/png' },
         { src: icon512 || icon192, sizes: '512x512', type: 'image/png' }]
      : [{ src: './icons/icon-192.png', sizes: '192x192', type: 'image/png', purpose: 'any maskable' },
         { src: './icons/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'any maskable' }]
  };
  const blob = new Blob([JSON.stringify(manifest)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.querySelector('link[rel="manifest"]');
  if (link) link.href = url;
}

function initIconPreview() {
  const icon192 = localStorage.getItem('customIcon192');
  _showIconPreview(icon192 || null);
  updateDynamicManifest();
}

// ── 彈出提示 ──────────────────────────────────────────────
function showAlert(m) { document.getElementById('msgText').innerText=m; document.getElementById('msgModal').style.display='flex'; }

let confirmCallback=null, confirmCancelCallback=null;
function showConfirm(msg,onConfirm,onCancel){ document.getElementById('confirmText').innerText=msg; document.getElementById('confirmModal').style.display='flex'; confirmCallback=onConfirm; confirmCancelCallback=onCancel; }
function closeConfirmModal(){ document.getElementById('confirmModal').style.display='none'; if(confirmCancelCallback) confirmCancelCallback(); }
function executeConfirm(){ document.getElementById('confirmModal').style.display='none'; if(confirmCallback) confirmCallback(); }

let inputCallback=null;
function openInputModal(title,onConfirm,initial=''){ document.getElementById('inputModalTitle').innerText=title; document.getElementById('inputModalText').value=initial; document.getElementById('inputModal').style.display='flex'; setTimeout(()=>document.getElementById('inputModalText').focus(),100); inputCallback=onConfirm; }
function closeInputModal(){ document.getElementById('inputModal').style.display='none'; }
function confirmInputModal(){ const v=document.getElementById('inputModalText').value.trim(); document.getElementById('inputModal').style.display='none'; if(inputCallback) inputCallback(v); }
