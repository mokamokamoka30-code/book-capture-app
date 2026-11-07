/* Microsoft Entra public client configuration.
   Client ID is an application identifier, not a password or client secret. */
const CLIENT_ID = 'd4fd61ba-0804-4158-aa15-8d1dc3a7560a';
const AUTHORITY = 'https://login.microsoftonline.com/consumers';
const SCOPES = ['User.Read', 'Files.ReadWrite'];

let msalApp = null;
let account = null;
const $ = (id) => document.getElementById(id);
const status = (message) => { $('status').textContent = message; };

function initAuth() {
  if (!window.msal || !window.msal.PublicClientApplication) {
    status('Microsoft認証ライブラリを読み込めません。ページを再読み込みしてください。');
    return;
  }

  try {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: AUTHORITY,
        redirectUri: window.location.origin + window.location.pathname
      },
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false
      }
    });

    msalApp.handleRedirectPromise()
      .then((result) => {
        account = result?.account || msalApp.getAllAccounts()[0] || null;
        if (account) status(`ログイン済み：${account.username}`);
        else status('未ログインです。Microsoftにログインしてください。');
      })
      .catch((e) => status('認証初期化エラー：' + e.message));
  } catch (e) {
    status('認証設定エラー：' + e.message);
  }
}

async function login() {
  if (!msalApp) return status('認証の準備ができていません。ページを再読み込みしてください。');
  try {
    const result = await msalApp.loginPopup({ scopes: SCOPES });
    account = result.account;
    status(`ログイン済み：${account.username}`);
  } catch (e) {
    status('ログインできませんでした：' + e.message);
  }
}

async function getToken() {
  if (!account) throw new Error('先にMicrosoftにログインしてください。');
  try {
    const result = await msalApp.acquireTokenSilent({ scopes: SCOPES, account });
    return result.accessToken;
  } catch {
    const result = await msalApp.acquireTokenPopup({ scopes: SCOPES, account });
    return result.accessToken;
  }
}

async function graph(path, token, options = {}) {
  const response = await fetch('https://graph.microsoft.com/v1.0' + path, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) }
  });
  if (!response.ok) throw new Error(`${response.status} ${await response.text()}`);
  return response.status === 204 ? null : response.json();
}

async function ensureFolder(name, token) {
  let root;
  try {
    root = await graph('/me/drive/root:/冊子電子化', token);
  } catch {
    root = await graph('/me/drive/root/children', token, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: '冊子電子化',
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail'
      })
    });
  }

  const safeName = name.replace(/[\\/:*?"<>|]/g, '_').trim();
  try {
    return await graph(`/me/drive/items/${root.id}:/${encodeURIComponent(safeName)}`, token);
  } catch {
    return await graph(`/me/drive/items/${root.id}/children`, token, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: safeName,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail'
      })
    });
  }
}

async function uploadFiles() {
  const bookId = $('bookId').value.trim();
  const files = [...$('photos').files];
  if (!bookId) return status('冊子番号を入力してください。');
  if (!files.length) return status('写真を選択してください。');

  try {
    status('OneDriveに接続しています…');
    const token = await getToken();
    const folder = await ensureFolder(bookId, token);

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const ext = (file.name.split('.').pop() || 'jpg').toLowerCase().replace(/[^a-z0-9]/g, '') || 'jpg';
      const name = `${String(i + 1).padStart(3, '0')}_${new Date().toISOString().replace(/[:.]/g, '-')}.${ext}`;
      const url = `https://graph.microsoft.com/v1.0/me/drive/items/${folder.id}:/${encodeURIComponent(name)}:/content`;
      const res = await fetch(url, {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': file.type || 'application/octet-stream'
        },
        body: file
      });
      if (!res.ok) throw new Error(`${file.name} の保存に失敗：${res.status} ${await res.text()}`);
      status(`${i + 1}/${files.length} 枚保存しました…`);
    }

    status(`保存完了：OneDrive／冊子電子化／${bookId}（${files.length}枚）`);
  } catch (e) {
    status('保存できませんでした：' + e.message);
  }
}

$('login').addEventListener('click', login);
$('upload').addEventListener('click', uploadFiles);
$('photos').addEventListener('change', () => {
  const files = [...$('photos').files];
  $('preview').textContent = files.length
    ? `選択中：${files.length}枚（${files.map(f => f.name).join('、')}）`
    : '';
});

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('./sw.js').catch(() => {}));
}

initAuth();
