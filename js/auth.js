const Auth = {
  _tokenClient: null,
  _resolveToken: null,

  init() {
    return new Promise((resolve) => {
      google.accounts.id.initialize({
        client_id: CONFIG.CLIENT_ID,
        callback: () => {}
      });
      resolve();
    });
  },

  // 取得有效的 access token（自動更新）
  getToken() {
    const token  = localStorage.getItem('gat');
    const expiry = parseInt(localStorage.getItem('gate') || '0');
    if (token && Date.now() < expiry - 60000) return Promise.resolve(token);
    return this._requestNewToken();
  },

  _requestNewToken() {
    return new Promise((resolve, reject) => {
      if (!this._tokenClient) {
        this._tokenClient = google.accounts.oauth2.initTokenClient({
          client_id: CONFIG.CLIENT_ID,
          scope: CONFIG.SCOPES,
          callback: (resp) => {
            if (resp.error) { reject(resp.error); return; }
            const expiry = Date.now() + resp.expires_in * 1000;
            localStorage.setItem('gat',  resp.access_token);
            localStorage.setItem('gate', expiry.toString());
            this._saveUserInfo(resp.access_token);
            if (this._resolveToken) { this._resolveToken(resp.access_token); this._resolveToken = null; }
            resolve(resp.access_token);
          }
        });
      }
      this._resolveToken = resolve;
      this._tokenClient.requestAccessToken({ prompt: this.isSignedIn() ? '' : 'consent' });
    });
  },

  async _saveUserInfo(token) {
    try {
      const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: 'Bearer ' + token }
      });
      const info = await r.json();
      localStorage.setItem('userInfo', JSON.stringify({
        name:    info.name,
        email:   info.email,
        picture: info.picture
      }));
    } catch(e) {}
  },

  isSignedIn() {
    const expiry = parseInt(localStorage.getItem('gate') || '0');
    return Date.now() < expiry - 60000;
  },

  getUserInfo() {
    try { return JSON.parse(localStorage.getItem('userInfo') || 'null'); } catch(e) { return null; }
  },

  signOut() {
    const token = localStorage.getItem('gat');
    if (token) google.accounts.oauth2.revoke(token, () => {});
    ['gat','gate','userInfo','spreadsheetId','appDataCache'].forEach(k => localStorage.removeItem(k));
    // 保留 historyCache 和 appTemplates
    Object.keys(localStorage)
      .filter(k => k.startsWith('historyCache_'))
      .forEach(k => localStorage.removeItem(k));
    location.reload();
  }
};
