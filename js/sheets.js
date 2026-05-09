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
      this.setValues(`${CONFIG.SHEETS.RECORDS}!A1:H1`,
        [['日期時間','收支類型','項目名稱','金額','會計科目','轉出帳戶(扣款)','轉入帳戶(存入)','備註']]),
      this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!A1:D1`,
        [['帳戶名稱','初始餘額','目前餘額','備註']]),
      this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!F1:L1`,
        [['信用卡名稱','總信用額度','本月應付卡費','剩餘可用額度','結帳日','繳款日','自動扣款帳戶']]),
      this.setValues(`${CONFIG.SHEETS.FIXED}!A1:J1`,
        [['項目名稱','收支類型','每月金額','每月扣款日','總期數','已繳期數','轉出帳戶(扣款)','轉入帳戶(存入)','會計科目','未繳總餘額']]),
      this.setValues(`${CONFIG.SHEETS.BUDGET}!A1:D1`,
        [['會計科目','每月預算','本月實際花費','剩餘預算差異']]),
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
    const rec = CONFIG.SHEETS.RECORDS;
    const rows = CONFIG.DEFAULT_CATEGORIES.map((c, i) => {
      const r = i + 2;
      return [
        c, 0,
        `=SUMIFS('${rec}'!D:D,'${rec}'!E:E,A${r},'${rec}'!B:B,"支出",'${rec}'!A:A,">="&EOMONTH(TODAY(),-1)+1,'${rec}'!A:A,"<="&EOMONTH(TODAY(),0))`,
        `=B${r}-C${r}`
      ];
    });
    await this.setValues(`${CONFIG.SHEETS.BUDGET}!A2:D${rows.length + 1}`, rows);
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
        const note = row[3] ? row[3].toString() : ''; // D欄：備註（C欄為公式）
        const balance = this._calcBankBalance(name, initial, allRecords);
        banks.push(name);
        dashboard.banks.push({ name, balance, note });
        dashboard.totalAssets += balance;
        if (name.indexOf(settings.investKeyword) !== -1) dashboard.investAssets += balance;
      }
      if (row[5]) {
        const name = row[5].toString();
        const limit = Number(row[6]) || 0;
        const billDate = Number(row[9]) || 1;   // J欄：結帳日
        const payDate = Number(row[10]) || 10;  // K欄：繳款日
        const autoPayAcc = row[11] ? row[11].toString() : ''; // L欄：自動扣款帳戶
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
      const fixedRows = await this.getValues(`${CONFIG.SHEETS.FIXED}!A:A`);
      const nextRow = fixedRows.length + 1;
      await this.setValues(`${CONFIG.SHEETS.FIXED}!A${nextRow}:J${nextRow}`, [[
        data.item, data.type, data.amount, d.getDate(),
        data.totalPeriods || '', data.paidPeriods || '',
        data.accountOut || '', data.accountIn || '', data.category || '',
        `=IF(E${nextRow}="",0,C${nextRow}*(E${nextRow}-F${nextRow}))`
      ]]);
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
    const rec = CONFIG.SHEETS.RECORDS;
    await this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!A${nextRow}:D${nextRow}`, [[
      bankData.name,
      bankData.initial || 0,
      `=B${nextRow}+SUMIFS('${rec}'!D:D,'${rec}'!G:G,A${nextRow})-SUMIFS('${rec}'!D:D,'${rec}'!F:F,A${nextRow})`,
      bankData.note || ''
    ]]);
    return { msg: `✅ 銀行「${bankData.name}」新增成功！`, syncData: await this.getSyncData(monthStr) };
  },

  async updateBankNote(bankName, newNote, monthStr) {
    const rows = await this.getValues(`${CONFIG.SHEETS.ACCOUNTS}!A:A`);
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === bankName) {
        await this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!D${i+1}`, [[newNote]]); // D欄：備註
        return { msg: '✅ 備註更新成功！', syncData: await this.getSyncData(monthStr) };
      }
    }
    return { msg: '❌ 找不到該銀行帳戶', syncData: null };
  },

  async addCreditCard(cardData, monthStr) {
    const rows = await this.getValues(`${CONFIG.SHEETS.ACCOUNTS}!F:F`);
    const nextRow = Math.max(rows.length + 1, 2);
    const rec = CONFIG.SHEETS.RECORDS;
    // F=名稱, G=額度, H=本月應付(公式), I=剩餘額度(公式), J=結帳日, K=繳款日, L=自動扣款帳戶
    await this.setValues(`${CONFIG.SHEETS.ACCOUNTS}!F${nextRow}:L${nextRow}`, [[
      cardData.name,
      cardData.limit,
      `=SUMIFS('${rec}'!D:D,'${rec}'!F:F,F${nextRow})-SUMIFS('${rec}'!D:D,'${rec}'!G:G,F${nextRow})`,
      `=G${nextRow}-H${nextRow}`,
      cardData.billDate,
      cardData.payDate,
      cardData.autoPayAcc || ''
    ]]);
    return { msg: `✅ 信用卡「${cardData.name}」新增成功！`, syncData: await this.getSyncData(monthStr) };
  },

  // ── 分類 ──────────────────────────────────────────────────

  async addCategory(name, monthStr) {
    const rows = await this.getValues(`${CONFIG.SHEETS.BUDGET}!A:A`);
    const nextRow = rows.length + 1;
    const rec = CONFIG.SHEETS.RECORDS;
    await this.setValues(`${CONFIG.SHEETS.BUDGET}!A${nextRow}:D${nextRow}`, [[
      name, 0,
      `=SUMIFS('${rec}'!D:D,'${rec}'!E:E,A${nextRow},'${rec}'!B:B,"支出",'${rec}'!A:A,">="&EOMONTH(TODAY(),-1)+1,'${rec}'!A:A,"<="&EOMONTH(TODAY(),0))`,
      `=B${nextRow}-C${nextRow}`
    ]]);
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

  // ── LINE 推播（透過 GAS Proxy）────────────────────────────

  async sendLineMessage(message) {
    const settings = await this.getSettings();
    if (!settings.lineUserId || !CONFIG.GAS_PROXY_URL) return;
    await fetch(CONFIG.GAS_PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        userId: settings.lineUserId,
        message,
        authKey: CONFIG.LINE_AUTH_KEY
      })
    });
  }
};
