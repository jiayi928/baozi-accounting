const CONFIG = {
  // 請到 https://console.cloud.google.com 建立專案後填入
  CLIENT_ID: '50636553434-f7pqcgi9ovbd8f6a4hi0tcmensg0nbjb.apps.googleusercontent.com',

  SCOPES: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
    'profile',
    'email'
  ].join(' '),

  SPREADSHEET_TITLE: '包子記帳資料',

  SHEETS: {
    RECORDS:    '紀錄頁',
    ACCOUNTS:   '帳戶與信用卡',
    FIXED:      '固定支出與轉帳',
    BUDGET:     'AI 專用分析與預算',
    SETTINGS:   '個人設定'
  },

  DEFAULT_SETTINGS: {
    appName:          '包子記帳',
    appShortName:     '記帳',
    investKeyword:    '富邦證券',
    investLabel:      '長期投資',
    excludeCategories:'ETF定期定額,學貸',
    alertThreshold:   3000,
    alertIgnoreWords: '生日,過年',
    themeColor:       '#000000',
    bgColor:          '#ffffff',
    lineToken:        '',
    lineUserId:       ''
  },

  DEFAULT_CATEGORIES: [
    '🍔 餐飲', '🚗 交通', '🏠 居住', '🎬 娛樂',
    '🛍️ 購物', '🛒 生活', '🏥 醫療', '📚 教育',
    '🛡️ 保險', '💰 薪水', '📈 投資', '🏧 手續費',
    '💧 水電', '📱 電信', '🌐 網路'
  ]
};
