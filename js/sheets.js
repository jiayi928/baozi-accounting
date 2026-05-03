const Sheets = {
  BASE: 'https://sheets.googleapis.com/v4/spreadsheets',
  DRIVE: 'https://www.googleapis.com/drive/v3/files',

  // ── 低階 API ──────────────────────────────────────────────

  async _req(url, opts = {}) {
    const token = await Auth.getToken();
    const res = await fetch(url, {
      ...opts,
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        ...(opts.headers || {})
      }
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || '請求失敗 ' + res.status);
    }
    return res.json();
  },

  sid() { return localStorage.getItem('spreadsheetId'); },

  async getValues(range) {
    const r = await this._req(`${this.BASE}/${this.sid()}/values/${encodeURIComponent(range)}`);
    return r.values || [];
  },

  async setValues(range, values) {
    return this._req(
      `${this.BASE}/${this.sid()}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
      { method: 'PUT', body: JSON.stringify({ values }) }
    );
  },

  async appendValues(sheetName, values) {
    return this._req(
      `${this.BASE}/${this.sid()}/values/${encodeURIComponent(sheetName)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
      { method: 'POST', body: JSON.stringify({ values: [values] }) }
    );
  },

  async batchUpdate(requests) {
    return this._req(`${this.BASE}/${this.sid()}:batchUpdate`, {
      method: 'POST', body: JSON.stringify({ requests })
    });
  },

  async getSheetId(sheetName) {
    const info = await this._req(`${this.BASE}/${this.sid()}?fields=sheets.properties`);
    const sheet = info.sheets.find(s => s.properties.title === sheetName);
    return sheet ? sheet.properties.sheetId : null;
  },

  // ── 初始化 ────────────────────────────────────────────────

  async initSpreadsheet() {
    const existing = localStorage.getItem('spreadsheetId');
    if (existing) {
      // 驗證是否還存在
      try {
        await this._req(`${this.BASE}/${existing}?fields=spreadsheetId`);
        return existing;
      } catch(e) {
        localStorage.removeItem('spreadsheetId');
      }
    }

    // 搜尋是否已有建立過的試算表
    const token = await Auth.getToken();
    const search = await fetch(
      `${this.DRIVE}?q=name='${CONFIG.SPREADSHEET_TITLE}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false&fields=files(id,name)`,
      { headers: { Authorization: 'Bearer ' + token } }
    ).then(r => r.json());

    if (search.files && search.files.length > 0) {
      const id = search.files[0].id;
      localStorage.setItem('spreadsheetId', id);
      return id;
    }

    // 建立全新試算表
    return this._createSpreadsheet();
  },

  async _createSpreadsheet() {
    const ss = await this._req(`${this.BASE}`, {
      method: 'POST',
      body: JSON.stringify({
        properties: { title: CONFIG.SPREADSHEET_TITLE },
        sheets: [
          { properties: { title: CONFIG.SHEETS.RECORDS } },
          { properties: { title: CONFIG.SHEETS.ACCOUNTS } },
          { properties: { title: CONFIG.SHEETS.FIXED } },
          { properties: { title: CONFIG.SHEETS.BUDGET } },
          { properties: { title: CONFIG.SHEETS.SETTINGS } }
        ]
      })
    });

    localStorage.setItem('spreadsheetId', ss.spreadsheetId);

    // 寫入標題列與預設資料
    await Promise.all([
      this.setValues(`${CONFIG.SHEETS.RECORDS}!A1:G1`,
        [['日期','類型','項目','金額','分類','帳戶(出)','帳戶(入)']]),
      this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!A1:D1`,
        [['銀行名稱','初始金額','備註','備用']]),
      this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!F1:L1`,
        [['卡片名稱','總額度','結帳日','繳款日','自動扣款帳戶','備用','備用']]),
      this.setValues(`${CONFIG.SHEETS.FIXED}!A1:I1`,
        [['項目','類型','金額','扣款日','總期數','已繳期數','帳戶(出)','帳戶(入)','分類']]),
      this.setValues(`${CONFIG.SHEETS.BUDGET}!A1:B1`,
        [['分類','月預算']]),
      this._writeDefaultSettings(),
      this._writeDefaultCategories()
    ]);

    return ss.spreadsheetId;
  },

  async _writeDefaultSettings() {
    const rows = Object.entries(CONFIG.DEFAULT_SETTINGS).map(([k,v]) => [k, v.toString()]);
    await this.setValues(`${CONFIG.SHEETS.SETTINGS}!A1:B${rows.length}`, rows);
  },

  async _writeDefaultCategories() {
    const rows = CONFIG.DEFAULT_CATEGORIES.map(c => [c, 0]);
    await this.setValues(`${CONFIG.SHEETS.BUDGET}!A2:B${rows.length + 1}`, rows);
  },

  // ── 設定 ──────────────────────────────────────────────────

  async getSettings() {
    const rows = await this.getValues(`${CONFIG.SHEETS.SETTINGS}!A:B`);
    const s = { ...CONFIG.DEFAULT_SETTINGS };
    rows.forEach(([k, v]) => { if (k && v !== undefined) s[k] = v; });
    s.alertThreshold = Number(s.alertThreshold) || 3000;
    return s;
  },

  async saveSettings(settings) {
    const rows = Object.entries(settings).map(([k,v]) => [k, v.toString()]);
    await this.setValues(`${CONFIG.SHEETS.SETTINGS}!A1:B${rows.length}`, rows);
  },

  // ── 取得全部資料（取代 getSyncData）─────────────────────

  async getSyncData(monthStr) {
    const [appData, history] = await Promise.all([
      this.getAppData(),
      this.getHistory(monthStr)
    ]);
    return { appData, history };
  },

  async getAppData() {
    const [acctRows, budgetRows, settings] = await Promise.all([
      this.getValues(`${CONFIG.SHEETS.ACCOUNTS}!A:L`),
      this.getValues(`${CONFIG.SHEETS.BUDGET}!A:B`),
      this.getSettings()
    ]);

    // 所有交易紀錄（計算餘額用）
    const allRecords = await this._getAllRecords();

    const banks = [], cards = [];
    const dashboard = { banks:[], cards:[], totalAssets:0, totalLiabilities:0, netWorth:0, investAssets:0, liquidAssets:0 };

    for (let i = 1; i < acctRows.length; i++) {
      const row = acctRows[i];
      if (row[0]) {
        const name = row[0].toString();
        const initial = Number(row[1]) || 0;
        const note = row[2] ? row[2].toString() : '';
        const balance = this._calcBankBalance(name, initial, allRecords);
        banks.push(name);
        dashboard.banks.push({ name, balance, note });
        dashboard.totalAssets += balance;
        if (name.indexOf(settings.investKeyword) !== -1) dashboard.investAssets += balance;
      }
      if (row[5]) {
        const name = row[5].toString();
        const limit = Number(row[6]) || 0;
        const billDate = Number(row[7]) || 1;
        const payDate = Number(row[8]) || 10;
        const autoPayAcc = row[9] ? row[9].toString() : '';
        const totalDebt = this._calcCardDebt(name, allRecords);
        const currentBill = this._calcCurrentBill(name, totalDebt, billDate, allRecords);
        const remain = limit - totalDebt;
        cards.push(name);
        dashboard.cards.push({ name, totalDebt, remain, billDate, payDate, currentBill });
        dashboard.totalLiabilities += totalDebt;
      }
    }

    dashboard.liquidAssets = dashboard.totalAssets - dashboard.investAssets;
    dashboard.netWorth = dashboard.totalAssets - dashboard.totalLiabilities;

    const categories = [];
    for (let i = 1; i < budgetRows.length; i++) {
      if (budgetRows[i][0]) categories.push(budgetRows[i][0].toString());
    }

    return { banks, cards, categories, dashboard, settings };
  },

  _getAllRecords() {
    return this.getValues(`${CONFIG.SHEETS.RECORDS}!A2:G`).then(rows =>
      rows.filter(r => r[0]).map((r, i) => ({
        row: i + 2,
        date: new Date(r[0]),
        type: r[1] || '',
        item: r[2] || '',
        amount: Number(r[3]) || 0,
        category: r[4] || '未分類',
        accOut: r[5] || '',
        accIn: r[6] || ''
      }))
    );
  },

  _calcBankBalance(name, initial, records) {
    let bal = initial;
    records.forEach(r => {
      if (r.accIn === name) bal += r.amount;
      if (r.accOut === name) bal -= r.amount;
    });
    return bal;
  },

  _calcCardDebt(name, records) {
    let debt = 0;
    const expTypes = ['支出','固定支出','訂閱','固定','點數折抵'];
    records.forEach(r => {
      if (r.accOut === name && expTypes.includes(r.type)) debt += r.amount;
      if (r.accIn === name) debt -= r.amount;
    });
    return Math.max(0, debt);
  },

  _calcCurrentBill(name, totalDebt, billDate, records) {
    const today = new Date();
    const d = today.getDate();
    let start = new Date(today.getFullYear(), today.getMonth(), billDate);
    if (d < billDate) start = new Date(today.getFullYear(), today.getMonth() - 1, billDate);
    start.setDate(start.getDate() + 1);

    let paidSinceBill = 0;
    records.forEach(r => {
      if (r.accIn === name && r.date >= start) paidSinceBill += r.amount;
    });
    return Math.max(0, totalDebt - paidSinceBill);
  },

  // ── 交易紀錄 ──────────────────────────────────────────────

  async getHistory(monthStr) {
    const rows = await this.getValues(`${CONFIG.SHEETS.RECORDS}!A2:G`);
    const results = [];
    for (let i = rows.length - 1; i >= 0; i--) {
      if (!rows[i][0]) continue;
      const d = new Date(rows[i][0]);
      const m = (d.getMonth()+1).toString().padStart(2,'0');
      const y = d.getFullYear();
      if (`${y}-${m}` !== monthStr) continue;
      const dd = d.getDate().toString().padStart(2,'0');
      results.push({
        row: i + 2,
        date: `${m}/${dd}`,
        fullDate: `${y}-${m}-${dd}`,
        type: rows[i][1] || '',
        item: rows[i][2] || '',
        amount: Number(rows[i][3]) || 0,
        category: rows[i][4] || '未分類',
        accOut: rows[i][5] || '',
        accIn: rows[i][6] || ''
      });
    }
    return results;
  },

  async recordData(data, monthStr) {
    if (data.isFixed) {
      const d = new Date(data.date);
      if (data.type === '固定支出' || data.type === '訂閱') data.type = '支出';
      await this.appendValues(CONFIG.SHEETS.FIXED, [
        data.item, data.type, data.amount, d.getDate(),
        data.totalPeriods || '', data.paidPeriods || '',
        data.accountOut || '', data.accountIn || '', data.category || ''
      ]);
      return { msg: '✅ 已加入自動扣款排程！', syncData: await this.getSyncData(monthStr) };
    }
    await this.appendValues(CONFIG.SHEETS.RECORDS, [
      data.date, data.type, data.item, data.amount,
      data.category || '', data.accountOut || '', data.accountIn || ''
    ]);
    await this._sortRecords();
    return { msg: '✅ 記帳成功！', syncData: await this.getSyncData(monthStr) };
  },

  async updateRecordData(data, monthStr) {
    await this.setValues(
      `${CONFIG.SHEETS.RECORDS}!A${data.row}:G${data.row}`,
      [[data.date, data.type, data.item, data.amount,
        data.category || '', data.accountOut || '', data.accountIn || '']]
    );
    await this._sortRecords();
    return { msg: '💾 紀錄更新成功！', syncData: await this.getSyncData(monthStr) };
  },

  async deleteRecord(rowNum, monthStr) {
    const sheetId = await this.getSheetId(CONFIG.SHEETS.RECORDS);
    await this.batchUpdate([{
      deleteDimension: {
        range: { sheetId, dimension: 'ROWS', startIndex: rowNum - 1, endIndex: rowNum }
      }
    }]);
    return { msg: '✅ 刪除成功！', syncData: await this.getSyncData(monthStr) };
  },

  async _sortRecords() {
    const sheetId = await this.getSheetId(CONFIG.SHEETS.RECORDS);
    const rows = await this.getValues(`${CONFIG.SHEETS.RECORDS}!A:G`);
    const lastRow = rows.length;
    if (lastRow <= 1) return;
    await this.batchUpdate([{
      sortRange: {
        range: { sheetId, startRowIndex: 1, endRowIndex: lastRow, startColumnIndex: 0, endColumnIndex: 7 },
        sortSpecs: [{ dimensionIndex: 0, sortOrder: 'ASCENDING' }]
      }
    }]);
  },

  // ── 帳戶管理 ──────────────────────────────────────────────

  async addBankAccount(bankData, monthStr) {
    const rows = await this.getValues(`${CONFIG.SHEETS.ACCOUNTS}!A:A`);
    const nextRow = rows.length + 1;
    await this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!A${nextRow}:D${nextRow}`, [[
      bankData.name, bankData.initial || 0, bankData.note || '', ''
    ]]);
    return { msg: `✅ 銀行「${bankData.name}」新增成功！`, syncData: await this.getSyncData(monthStr) };
  },

  async updateBankNote(bankName, newNote, monthStr) {
    const rows = await this.getValues(`${CONFIG.SHEETS.ACCOUNTS}!A:C`);
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === bankName) {
        await this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!C${i+1}`, [[newNote]]);
        return { msg: '✅ 備註更新成功！', syncData: await this.getSyncData(monthStr) };
      }
    }
    return { msg: '❌ 找不到該銀行帳戶', syncData: null };
  },

  async addCreditCard(cardData, monthStr) {
    const rows = await this.getValues(`${CONFIG.SHEETS.ACCOUNTS}!F:F`);
    const nextRow = Math.max(rows.length + 1, 2);
    await this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!F${nextRow}:L${nextRow}`, [[
      cardData.name, cardData.limit, cardData.billDate,
      cardData.payDate, cardData.autoPayAcc || '', '', ''
    ]]);
    return { msg: `✅ 信用卡「${cardData.name}」新增成功！`, syncData: await this.getSyncData(monthStr) };
  },

  // ── 分類 ──────────────────────────────────────────────────

  async addCategory(name, monthStr) {
    const rows = await this.getValues(`${CONFIG.SHEETS.BUDGET}!A:B`);
    const nextRow = rows.length + 1;
    await this.setValues(`${CONFIG.SHEETS.BUDGET}!A${nextRow}:B${nextRow}`, [[name, 0]]);
    return { msg: `✅ 分類「${name}」新增成功！`, syncData: await this.getSyncData(monthStr) };
  },

  // ── 帳戶明細 ──────────────────────────────────────────────

  async getAccountDetails(accountName) {
    const rows = await this.getValues(`${CONFIG.SHEETS.RECORDS}!A2:G`);
    const details = [];
    for (let i = rows.length - 1; i >= 0; i--) {
      if (!rows[i][0]) continue;
      if (rows[i][5] !== accountName && rows[i][6] !== accountName) continue;
      const d = new Date(rows[i][0]);
      const isIn = rows[i][6] === accountName;
      details.push({
        date: `${d.getMonth()+1}/${d.getDate()}`,
        item: rows[i][2] || '',
        effect: (isIn ? '+ $' : '- $') + (Number(rows[i][3])||0).toLocaleString(),
        effectClass: isIn ? 'val-pos' : 'val-neg'
      });
      if (details.length >= 30) break;
    }
    return details;
  },

  // ── 固定支出 ──────────────────────────────────────────────

  async getFixedExpenses() {
    const rows = await this.getValues(`${CONFIG.SHEETS.FIXED}!A2:I`);
    return rows
      .filter(r => r[0])
      .map(r => ({
        item: r[0], type: r[1], amount: Number(r[2])||0,
        day: r[3], total: r[4]||'-', paid: r[5]||'-',
        account: r[6]||'未指定帳戶', category: r[8]||''
      }));
  },

  // ── LINE 推播（需要後端 Proxy，此為客戶端版本）─────────

  async sendLineMessage(message) {
    const settings = await this.getSettings();
    if (!settings.lineToken || !settings.lineUserId) return;
    // 注意：LINE API 不支援瀏覽器直接呼叫（CORS）
    // 此功能需搭配 Cloud Function proxy 才能運作
    // 暫時以 console.log 替代
    console.log('[LINE]', message);
  }
};
