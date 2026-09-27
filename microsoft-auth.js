(function attachPayrollMicrosoftAuth(root, factory) {
  'use strict';

  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./cloud-snapshot'));
    return;
  }

  root.PayrollMicrosoftAuth = factory(root.PayrollCloudSnapshot);
}(typeof globalThis !== 'undefined' ? globalThis : this, function buildPayrollMicrosoftAuth(cloudSnapshot) {
  'use strict';

  if (!cloudSnapshot?.CloudSnapshotError) {
    throw new Error('PayrollCloudSnapshot must be loaded before PayrollMicrosoftAuth.');
  }

  const { CloudSnapshotError } = cloudSnapshot;
  const AUTHORITY = 'https://login.microsoftonline.com/consumers';
  const DEFAULT_SCOPES = Object.freeze(['Files.ReadWrite.AppFolder']);
  const EXPECTED_VENDOR_PATH = './vendor/msal-browser.min.js';

  function fail(message, code, status = 0) {
    throw new CloudSnapshotError(message, code, status);
  }

  function validateClientId(value) {
    const clientId = String(value || '').trim().toLowerCase();
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(clientId)) {
      fail('Microsoft 应用编号格式无效，请在电脑端重新配置。', 'INVALID_MICROSOFT_CLIENT_ID');
    }
    return clientId;
  }

  function normalizeRedirectUri(value, locationObject) {
    let parsed;
    try {
      parsed = new URL(String(value || `${locationObject?.origin || ''}${locationObject?.pathname || '/'}`));
    } catch (_error) {
      fail('Microsoft 登录回调地址格式无效。', 'INVALID_REDIRECT_URI');
    }
    const localDevelopment = ['localhost', '127.0.0.1', '[::1]'].includes(parsed.hostname);
    if ((parsed.protocol !== 'https:' && !(localDevelopment && parsed.protocol === 'http:'))
      || parsed.username || parsed.password || parsed.search || parsed.hash) {
      fail('Microsoft 登录回调必须是已登记的 HTTPS 网页地址。', 'INVALID_REDIRECT_URI');
    }
    return parsed.href;
  }

  function createMicrosoftAuthProvider(options = {}) {
    const msalRuntime = options.msal || (typeof globalThis !== 'undefined' ? globalThis.msal : null);
    const clientId = validateClientId(options.clientId);
    const redirectUri = normalizeRedirectUri(options.redirectUri, options.location || globalThis.location);
    const scopes = Array.isArray(options.scopes) && options.scopes.length
      ? Object.freeze(options.scopes.map((scope) => String(scope || '').trim()).filter(Boolean))
      : DEFAULT_SCOPES;
    let application = null;
    let account = null;
    let initialization = null;

    function isAvailable() {
      return typeof msalRuntime?.PublicClientApplication === 'function';
    }

    async function ensureInitialized() {
      if (!isAvailable()) {
        fail(`手机网页缺少官方 Microsoft 登录组件（${EXPECTED_VENDOR_PATH}），请联系管理员完成配置。`, 'MSAL_NOT_AVAILABLE');
      }
      if (initialization) return initialization;
      initialization = (async () => {
        application = new msalRuntime.PublicClientApplication({
          auth: {
            clientId,
            authority: AUTHORITY,
            redirectUri,
            navigateToLoginRequestUrl: false,
          },
          cache: {
            // Access/ID/refresh tokens must disappear on reload and are never
            // written to localStorage or sessionStorage by this PWA.
            cacheLocation: 'memoryStorage',
            temporaryCacheLocation: 'memoryStorage',
            storeAuthStateInCookie: false,
          },
        });
        if (typeof application.initialize === 'function') await application.initialize();
        if (typeof application.handleRedirectPromise === 'function') {
          const redirectResult = await application.handleRedirectPromise();
          if (redirectResult?.account) account = redirectResult.account;
        }
        if (!account && typeof application.getAllAccounts === 'function') {
          account = application.getAllAccounts()[0] || null;
        }
      })();
      try {
        await initialization;
      } catch (_error) {
        application = null;
        account = null;
        initialization = null;
        fail('Microsoft 登录组件初始化失败，请检查网页地址配置。', 'MSAL_INITIALIZE_FAILED');
      }
    }

    async function signIn(request = {}) {
      await ensureInitialized();
      const requestedScopes = Array.isArray(request.scopes) && request.scopes.length ? request.scopes : scopes;
      const previousAccount = account;
      let result;
      try {
        result = await application.loginPopup({
          scopes: [...requestedScopes],
          prompt: 'select_account',
          redirectUri,
        });
      } catch (_error) {
        fail('Microsoft 登录未完成，请允许弹窗后重试。', 'ONEDRIVE_AUTH_FAILED', 401);
      }
      const nextAccount = result?.account || null;
      if (!nextAccount) fail('Microsoft 登录没有返回账号信息，请重试。', 'ONEDRIVE_AUTH_FAILED', 401);
      if (previousAccount && previousAccount.homeAccountId !== nextAccount.homeAccountId
        && typeof application.clearCache === 'function') {
        try { await application.clearCache({ account: previousAccount }); } catch (_error) { /* Memory is dropped when the page locks. */ }
      }
      account = nextAccount;
      return { signedIn: true };
    }

    async function acquireToken(request = {}) {
      await ensureInitialized();
      if (!account) fail('请先点击“登录 Microsoft 账号”。', 'MICROSOFT_SIGNIN_REQUIRED', 401);
      const requestedScopes = Array.isArray(request.scopes) && request.scopes.length ? request.scopes : scopes;
      let result;
      try {
        result = await application.acquireTokenSilent({
          account,
          scopes: [...requestedScopes],
          redirectUri,
        });
      } catch (_error) {
        const expiredAccount = account;
        account = null;
        if (typeof application.clearCache === 'function') {
          try { await application.clearCache(expiredAccount ? { account: expiredAccount } : undefined); } catch (_cacheError) { /* Memory is dropped when the page locks. */ }
        }
        fail('Microsoft 登录已过期，请重新点击登录。', 'MICROSOFT_REAUTH_REQUIRED', 401);
      }
      const accessToken = String(result?.accessToken || '').trim();
      if (!accessToken) fail('Microsoft 登录没有返回可用权限，请重新登录。', 'TOKEN_UNAVAILABLE', 401);
      return { accessToken };
    }

    async function clearSession() {
      const previousAccount = account;
      account = null;
      if (application && typeof application.clearCache === 'function') {
        try { await application.clearCache(previousAccount ? { account: previousAccount } : undefined); } catch (_error) { /* Memory is dropped below. */ }
      }
      application = null;
      initialization = null;
      return { signedIn: false };
    }

    function hasAccount() {
      return Boolean(account);
    }

    // Initialize while the login screen is being rendered. This keeps the
    // subsequent `loginPopup` directly tied to the user's button click, which
    // avoids popup blockers on stricter Android browsers.
    if (isAvailable() && typeof globalThis.document === 'object') ensureInitialized().catch(() => {});

    return Object.freeze({
      acquireToken,
      clearSession,
      hasAccount,
      isAvailable,
      signIn,
    });
  }

  return Object.freeze({
    AUTHORITY,
    DEFAULT_SCOPES,
    EXPECTED_VENDOR_PATH,
    createMicrosoftAuthProvider,
  });
}));
