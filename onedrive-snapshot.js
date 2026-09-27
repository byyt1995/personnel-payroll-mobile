(function attachPayrollOneDriveSnapshot(root, factory) {
  'use strict';

  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./cloud-snapshot'));
    return;
  }

  root.PayrollOneDriveSnapshot = factory(root.PayrollCloudSnapshot);
}(typeof globalThis !== 'undefined' ? globalThis : this, function buildPayrollOneDriveSnapshot(cloudSnapshot) {
  'use strict';

  if (!cloudSnapshot?.CloudSnapshotError || !cloudSnapshot?.__internal) {
    throw new Error('PayrollCloudSnapshot must be loaded before PayrollOneDriveSnapshot.');
  }

  const { CloudSnapshotError, __internal } = cloudSnapshot;
  const {
    base64UrlToBytes,
    cleanFragment,
    createSnapshotQueryApi,
    cryptoApi,
    decryptSnapshot,
    hasValidStoredSecretRecord,
    publicMetadata,
    readResponsePayload,
    unwrapSecretRecord,
    validatePin,
    verifySnapshotPin,
    wrapSecretRecord,
  } = __internal;

  const SNAPSHOT_FILE_NAME = 'payroll-snapshot.enc.json';
  const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
  const GRAPH_METADATA_URL = `${GRAPH_BASE_URL}/me/drive/special/approot:/${SNAPSHOT_FILE_NAME}`
    + '?$select=id,name,size,eTag,lastModifiedDateTime,@microsoft.graph.downloadUrl';
  const RECEIPT_ASSET_FORMAT = 'personnel-payroll-receipt-asset';
  const RECEIPT_ASSET_VERSION = 1;
  const RECEIPT_ASSET_DIRECTORY = 'payment-receipts';
  const MAX_RECEIPT_BYTES = 25 * 1024 * 1024;
  const MAX_RECEIPT_ENVELOPE_BYTES = 36 * 1024 * 1024;
  const DEFAULT_SCOPES = Object.freeze(['Files.ReadWrite.AppFolder']);
  const STATE_STORAGE_PREFIX = 'payroll-onedrive.snapshot-state.v1.';
  const PAIRING_STORAGE_KEY = 'payroll-onedrive.pairing.v1';
  const PAIRING_HINT_STORAGE_KEY = 'payroll-onedrive.pairing-hint.v1';
  const PAIRING_AAD = 'payroll-onedrive-pairing:v1';
  const MAX_METADATA_BYTES = 128 * 1024;

  function fail(message, code, status = 0) {
    throw new CloudSnapshotError(message, code, status);
  }

  function hasOwn(object, key) {
    return Object.prototype.hasOwnProperty.call(object || {}, key);
  }

  function validateNoClientSecret(options) {
    // Delegated browser auth is supplied by an MSAL/PKCE wrapper. A public
    // client cannot safely hold a client secret, so this adapter rejects one.
    if (hasOwn(options, 'clientSecret') || hasOwn(options, 'client_secret')) {
      fail('手机端不能配置客户端密钥，请使用授权码加 PKCE 登录。', 'CLIENT_SECRET_FORBIDDEN');
    }
  }

  function validateClientId(value) {
    const clientId = String(value || '').trim().toLowerCase();
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(clientId)) {
      fail('Microsoft 应用编号格式无效，请在电脑端重新配置。', 'INVALID_MICROSOFT_CLIENT_ID');
    }
    return clientId;
  }

  function validatePairing(value) {
    const snapshotId = String(value?.snapshotId || '').trim();
    const dataKey = String(value?.dataKey || '').trim();
    if (!/^[A-Za-z0-9_-]{16,128}$/.test(snapshotId)) {
      fail('OneDrive 配对资料编号无效。', 'INVALID_PAIRING');
    }
    const keyBytes = base64UrlToBytes(dataKey, 'OneDrive 配对密钥');
    if (keyBytes.length !== 32) fail('OneDrive 配对密钥长度无效。', 'INVALID_PAIRING');
    keyBytes.fill(0);
    const pairing = { snapshotId, dataKey };
    if (value?.clientId) pairing.clientId = validateClientId(value.clientId);
    return Object.freeze(pairing);
  }

  function detectOneDriveMode(options = {}) {
    const locationObject = options.location || globalThis.location;
    const historyObject = options.history || globalThis.history;
    const hash = String(locationObject?.hash || '');
    if (!hash || hash === '#') return { active: false, pairing: null, clientId: '' };
    const params = new URLSearchParams(hash.slice(1));
    if (String(params.get('provider') || '').toLowerCase() !== 'onedrive') {
      return { active: false, pairing: null, clientId: '' };
    }

    // Remove the key-bearing fragment before validation, network access or UI
    // rendering so it cannot remain in history or a screenshot.
    cleanFragment(locationObject, historyObject);
    if (params.get('v') !== '1') fail('OneDrive 手机配对链接版本不受支持。', 'INVALID_PAIRING');
    const clientId = validateClientId(params.get('client'));
    const pairing = validatePairing({
      snapshotId: params.get('s'),
      dataKey: params.get('k'),
      clientId,
    });
    return { active: true, pairing, clientId };
  }

  function readPairingHint(storage) {
    let raw;
    try { raw = storage?.getItem?.(PAIRING_HINT_STORAGE_KEY); } catch (_error) { return null; }
    if (!raw) return null;
    try {
      const parsed = JSON.parse(raw);
      if (Number(parsed?.version) !== 1 || parsed?.provider !== 'onedrive') return null;
      return Object.freeze({ provider: 'onedrive', clientId: validateClientId(parsed.client_id) });
    } catch (_error) {
      return null;
    }
  }

  function hasStoredPairing(storage = globalThis.localStorage) {
    try {
      const raw = storage?.getItem?.(PAIRING_STORAGE_KEY);
      if (!raw || !readPairingHint(storage)) return false;
      return hasValidStoredSecretRecord(JSON.parse(raw));
    } catch (_error) {
      return false;
    }
  }

  function validateETag(value) {
    const eTag = String(value || '').trim();
    if (!eTag || eTag.length > 1024 || /[\u0000-\u001f\u007f]/.test(eTag)) {
      fail('OneDrive 文件缺少有效版本标识。', 'INVALID_ONEDRIVE_METADATA');
    }
    return eTag;
  }

  function validateReceiptAssetId(value) {
    const assetId = String(value || '').trim();
    if (!/^ra_[A-Za-z0-9_-]{43}$/.test(assetId)) {
      fail('付款凭据编号无效。', 'INVALID_RECEIPT_ASSET_ID');
    }
    return assetId;
  }

  function receiptAssetFileName(assetIdValue) {
    return `${validateReceiptAssetId(assetIdValue)}.enc.json`;
  }

  function receiptAssetGraphMetadataUrl(assetIdValue) {
    const fileName = encodeURIComponent(receiptAssetFileName(assetIdValue));
    return `${GRAPH_BASE_URL}/me/drive/special/approot:/${RECEIPT_ASSET_DIRECTORY}/${fileName}`
      + '?$select=id,name,size,eTag,lastModifiedDateTime,@microsoft.graph.downloadUrl';
  }

  function validateDownloadUrl(value) {
    let parsed;
    try {
      parsed = new URL(String(value || ''));
    } catch (_error) {
      fail('OneDrive 没有返回有效的临时下载地址。', 'INVALID_ONEDRIVE_METADATA');
    }
    if (parsed.protocol !== 'https:' || parsed.username || parsed.password) {
      fail('OneDrive 临时下载地址不安全，已停止读取。', 'INVALID_ONEDRIVE_METADATA');
    }
    return parsed.href;
  }

  function sanitizeFileMetadata(raw, expectedName, maximumBytes) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
      fail('OneDrive 文件信息格式无效。', 'INVALID_ONEDRIVE_METADATA');
    }
    if (String(raw.name || '') !== expectedName) {
      fail('OneDrive 返回的不是约定资料文件。', 'INVALID_ONEDRIVE_METADATA');
    }
    const itemId = String(raw.id || '').trim();
    if (!itemId || itemId.length > 512 || /[\u0000-\u001f\u007f]/.test(itemId)) {
      fail('OneDrive 文件编号无效。', 'INVALID_ONEDRIVE_METADATA');
    }
    const size = Number(raw.size);
    if (!Number.isSafeInteger(size) || size < 1 || size > maximumBytes) {
      fail('OneDrive 资料文件大小无效。', 'INVALID_ONEDRIVE_METADATA');
    }
    const lastModifiedDateTime = String(raw.lastModifiedDateTime || '');
    if (!lastModifiedDateTime || Number.isNaN(Date.parse(lastModifiedDateTime))) {
      fail('OneDrive 文件更新时间无效。', 'INVALID_ONEDRIVE_METADATA');
    }
    return {
      itemId,
      name: expectedName,
      size,
      eTag: validateETag(raw.eTag || raw['@odata.etag']),
      lastModifiedDateTime,
      downloadUrl: validateDownloadUrl(raw['@microsoft.graph.downloadUrl']),
    };
  }

  function sanitizeMetadata(raw) {
    return sanitizeFileMetadata(raw, SNAPSHOT_FILE_NAME, 20 * 1024 * 1024);
  }

  function publicRemoteMetadata(metadata) {
    if (!metadata) return null;
    return {
      item_id: metadata.itemId,
      name: metadata.name,
      size: metadata.size,
      e_tag: metadata.eTag,
      last_modified_at: metadata.lastModifiedDateTime,
    };
  }

  function normalizeToken(result) {
    const token = typeof result === 'string' ? result : result?.accessToken;
    const clean = String(token || '').trim();
    if (clean.length < 16 || clean.length > 16384 || /\s/.test(clean)) {
      fail('Microsoft 登录没有返回有效访问令牌。', 'TOKEN_UNAVAILABLE', 401);
    }
    return clean;
  }

  function resolveAuthProvider(options) {
    const provider = options.authProvider;
    const acquireToken = typeof options.acquireToken === 'function'
      ? options.acquireToken
      : typeof provider === 'function'
        ? provider
        : typeof provider?.acquireToken === 'function'
          ? provider.acquireToken.bind(provider)
          : null;
    if (!acquireToken) fail('尚未配置 Microsoft 登录。', 'AUTH_PROVIDER_REQUIRED');
    return Object.freeze({
      acquireToken,
      signIn: typeof provider?.signIn === 'function' ? provider.signIn.bind(provider) : async () => ({ signedIn: true }),
      clearSession: typeof provider?.clearSession === 'function' ? provider.clearSession.bind(provider) : async () => {},
      hasAccount: typeof provider?.hasAccount === 'function' ? provider.hasAccount.bind(provider) : () => true,
      isAvailable: typeof provider?.isAvailable === 'function' ? provider.isAvailable.bind(provider) : () => true,
    });
  }

  async function parseMetadataResponse(response) {
    if (!response || typeof response.ok !== 'boolean') {
      fail('Microsoft Graph 返回格式无效。', 'INVALID_GRAPH_RESPONSE');
    }
    if (!response.ok) {
      const status = Number(response.status) || 0;
      if (status === 401 || status === 403) {
        fail('Microsoft 登录已失效或没有 OneDrive 应用文件夹权限。', 'ONEDRIVE_AUTH_REJECTED', status);
      }
      if (status === 404) fail('OneDrive 中尚未找到手机查询资料。', 'SNAPSHOT_NOT_FOUND', status);
      fail(`OneDrive 文件信息暂时无法读取（${status}）。`, 'METADATA_FETCH_FAILED', status);
    }
    const lengthHeader = response.headers?.get?.('content-length');
    if (lengthHeader && Number(lengthHeader) > MAX_METADATA_BYTES) {
      fail('OneDrive 文件信息过大，已停止读取。', 'INVALID_GRAPH_RESPONSE');
    }
    let payload;
    if (typeof response.text === 'function') {
      const text = await response.text();
      if (text.length > MAX_METADATA_BYTES) fail('OneDrive 文件信息过大，已停止读取。', 'INVALID_GRAPH_RESPONSE');
      try {
        payload = JSON.parse(text);
      } catch (_error) {
        fail('OneDrive 文件信息不是有效 JSON。', 'INVALID_GRAPH_RESPONSE');
      }
    } else if (typeof response.json === 'function') {
      payload = await response.json();
    } else {
      fail('Microsoft Graph 返回格式无效。', 'INVALID_GRAPH_RESPONSE');
    }
    return payload;
  }

  async function defaultGetMetadata({ accessToken, fetch: fetchImpl, graphUrl }) {
    return fetchImpl(graphUrl, {
      method: 'GET',
      cache: 'no-store',
      credentials: 'omit',
      headers: {
        Accept: 'application/json',
        Authorization: `Bearer ${accessToken}`,
      },
    });
  }

  function normalizeReceiptRecord(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      fail('付款凭据索引格式无效。', 'INVALID_RECEIPT_INDEX');
    }
    const mimeType = String(value.mime_type || '').trim().toLowerCase();
    if (!['image/png', 'image/jpeg', 'application/pdf'].includes(mimeType)) {
      fail('付款凭据类型无效。', 'INVALID_RECEIPT_INDEX');
    }
    const sizeBytes = Number(value.size_bytes);
    const personId = Number(value.person_id);
    const year = Number(value.year);
    const month = Number(value.month);
    const sourceScope = String(value.source_scope || 'exact_person_month');
    if (!Number.isSafeInteger(sizeBytes) || sizeBytes < 1 || sizeBytes > MAX_RECEIPT_BYTES
      || !Number.isSafeInteger(personId) || personId < 1
      || !Number.isSafeInteger(year) || year < 2000 || year > 2200
      || !Number.isSafeInteger(month) || month < 1 || month > 12) {
      fail('付款凭据索引中的人员、月份或大小无效。', 'INVALID_RECEIPT_INDEX');
    }
    if (!['exact_person_month', 'legacy_shared_folder'].includes(sourceScope)) {
      fail('付款凭据来源范围无效。', 'INVALID_RECEIPT_INDEX');
    }
    return Object.freeze({
      asset_id: validateReceiptAssetId(value.asset_id),
      person_id: personId,
      year,
      month,
      display_name: String(value.display_name || '付款凭据')
        .replace(/[\u0000-\u001f\u007f]/g, '').slice(0, 180) || '付款凭据',
      mime_type: mimeType,
      size_bytes: sizeBytes,
      source_scope: sourceScope,
      created_at: String(value.created_at || '').slice(0, 40),
    });
  }

  function receiptAssetAad(metadata) {
    return [
      RECEIPT_ASSET_FORMAT,
      `v=${Number(metadata.version)}`,
      `asset=${metadata.asset_id}`,
      `sha256=${metadata.sha256}`,
      `mime=${metadata.mime_type}`,
      `size=${Number(metadata.size_bytes)}`,
      `key=${metadata.key_id}`,
    ].join('|');
  }

  function assertReceiptSignature(bytes, mimeType) {
    const value = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const matches = mimeType === 'image/png'
      ? value.length >= 8 && [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]
        .every((item, index) => value[index] === item)
      : mimeType === 'image/jpeg'
        ? value.length >= 3 && value[0] === 0xff && value[1] === 0xd8 && value[2] === 0xff
        : value.length >= 5 && value[0] === 0x25 && value[1] === 0x50
          && value[2] === 0x44 && value[3] === 0x46 && value[4] === 0x2d;
    if (!matches) fail('付款凭据内容与文件类型不一致。', 'INVALID_RECEIPT_CONTENT');
  }

  function validateReceiptEnvelope(raw, expectedRecord, expectedKeyId) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
      fail('付款凭据密文格式无效。', 'INVALID_RECEIPT_ENVELOPE');
    }
    const allowed = new Set([
      'format', 'version', 'asset_id', 'key_id', 'sha256', 'mime_type', 'size_bytes',
      'algorithm', 'iv', 'ciphertext',
    ]);
    if (Object.keys(raw).some((key) => !allowed.has(key))
      || raw.format !== RECEIPT_ASSET_FORMAT
      || Number(raw.version) !== RECEIPT_ASSET_VERSION
      || raw.algorithm !== 'AES-256-GCM') {
      fail('付款凭据密文版本或算法无效。', 'INVALID_RECEIPT_ENVELOPE');
    }
    const metadata = {
      format: RECEIPT_ASSET_FORMAT,
      version: RECEIPT_ASSET_VERSION,
      asset_id: validateReceiptAssetId(raw.asset_id),
      key_id: String(raw.key_id || ''),
      sha256: String(raw.sha256 || '').toLowerCase(),
      mime_type: String(raw.mime_type || '').toLowerCase(),
      size_bytes: Number(raw.size_bytes),
    };
    if (metadata.asset_id !== expectedRecord.asset_id
      || !/^[A-Za-z0-9_-]{8,128}$/.test(metadata.key_id)
      || metadata.key_id !== String(expectedKeyId || '')
      || !/^[a-f0-9]{64}$/.test(metadata.sha256)
      || metadata.mime_type !== expectedRecord.mime_type
      || metadata.size_bytes !== expectedRecord.size_bytes) {
      fail('付款凭据密文与索引不一致。', 'INVALID_RECEIPT_ENVELOPE');
    }
    const iv = base64UrlToBytes(raw.iv, '付款凭据随机量');
    const ciphertext = base64UrlToBytes(raw.ciphertext, '付款凭据密文');
    if (iv.length !== 12 || ciphertext.length !== metadata.size_bytes + 16
      || ciphertext.length > MAX_RECEIPT_BYTES + 16) {
      fail('付款凭据密文长度无效。', 'INVALID_RECEIPT_ENVELOPE');
    }
    return { metadata, iv, ciphertext };
  }

  async function decryptReceiptEnvelope(raw, expectedRecord, expectedKeyId, dataKey, selectedCrypto) {
    const validated = validateReceiptEnvelope(raw, expectedRecord, expectedKeyId);
    const rawKey = base64UrlToBytes(dataKey, '手机配对密钥');
    if (rawKey.length !== 32) fail('手机配对密钥长度无效。', 'INVALID_PAIRING');
    let plaintext;
    try {
      const key = await selectedCrypto.subtle.importKey('raw', rawKey, { name: 'AES-GCM' }, false, ['decrypt']);
      const Encoder = globalThis.TextEncoder;
      if (!Encoder) fail('当前浏览器不支持安全解密。', 'UNSUPPORTED_BROWSER');
      try {
        const output = await selectedCrypto.subtle.decrypt({
          name: 'AES-GCM',
          iv: validated.iv,
          additionalData: new Encoder().encode(receiptAssetAad(validated.metadata)),
          tagLength: 128,
        }, key, validated.ciphertext);
        plaintext = new Uint8Array(output);
      } catch (_error) {
        fail('付款凭据校验失败，文件可能已损坏。', 'RECEIPT_DECRYPT_FAILED');
      }
      if (plaintext.length !== validated.metadata.size_bytes) {
        fail('付款凭据解密后大小不一致。', 'INVALID_RECEIPT_CONTENT');
      }
      const digest = new Uint8Array(await selectedCrypto.subtle.digest('SHA-256', plaintext));
      const sha256 = Array.from(digest, (value) => value.toString(16).padStart(2, '0')).join('');
      digest.fill(0);
      if (sha256 !== validated.metadata.sha256) fail('付款凭据完整性校验失败。', 'INVALID_RECEIPT_CONTENT');
      assertReceiptSignature(plaintext, validated.metadata.mime_type);
      const BlobClass = globalThis.Blob;
      if (!BlobClass) fail('当前浏览器不支持安全预览附件。', 'UNSUPPORTED_BROWSER');
      return new BlobClass([plaintext], { type: validated.metadata.mime_type });
    } finally {
      rawKey.fill(0);
      if (plaintext) plaintext.fill(0);
    }
  }

  async function readReceiptEnvelopeResponse(response) {
    if (!response?.ok) {
      const status = Number(response?.status) || 0;
      fail(`付款凭据暂时无法下载（${status}）。`, 'RECEIPT_DOWNLOAD_FAILED', status);
    }
    const length = Number(response.headers?.get?.('content-length') || 0);
    if (length && (!Number.isSafeInteger(length) || length > MAX_RECEIPT_ENVELOPE_BYTES)) {
      fail('付款凭据密文过大，已停止读取。', 'INVALID_RECEIPT_ENVELOPE');
    }
    if (typeof response.text !== 'function') fail('付款凭据下载格式无效。', 'RECEIPT_DOWNLOAD_FAILED');
    const text = await response.text();
    if (!text || text.length > MAX_RECEIPT_ENVELOPE_BYTES) {
      fail('付款凭据密文大小无效。', 'INVALID_RECEIPT_ENVELOPE');
    }
    try { return JSON.parse(text); }
    catch (_error) { fail('付款凭据密文不是有效JSON。', 'INVALID_RECEIPT_ENVELOPE'); }
  }

  function stateStorageKey(snapshotId) {
    return `${STATE_STORAGE_PREFIX}${snapshotId}`;
  }

  function readStoredState(storage, snapshotId) {
    let raw;
    try {
      raw = storage.getItem(stateStorageKey(snapshotId));
    } catch (_error) {
      fail('当前浏览器无法读取 OneDrive 版本记录。', 'STORAGE_UNAVAILABLE');
    }
    if (!raw) return null;
    try {
      const value = JSON.parse(raw);
      if (Number(value?.version) !== 1
        || String(value?.snapshot_id || '') !== snapshotId
        || !Number.isSafeInteger(Number(value?.sequence))
        || Number(value.sequence) < 1) {
        fail('本机保存的 OneDrive 版本记录已损坏。', 'INVALID_ROLLBACK_STATE');
      }
      return {
        sequence: Number(value.sequence),
        eTag: validateETag(value.e_tag),
      };
    } catch (error) {
      if (error instanceof CloudSnapshotError) throw error;
      fail('本机保存的 OneDrive 版本记录已损坏。', 'INVALID_ROLLBACK_STATE');
    }
  }

  function verifyAndStoreVersion(storage, snapshot, eTag) {
    const previous = readStoredState(storage, snapshot.snapshot_id);
    if (previous && Number(snapshot.sequence) < previous.sequence) {
      fail('检测到旧版本 OneDrive 资料，已拒绝打开以防数据回退。', 'SNAPSHOT_ROLLBACK');
    }
    if (previous && Number(snapshot.sequence) === previous.sequence && eTag !== previous.eTag) {
      fail('同一资料版本对应了不同的 OneDrive 文件，已停止读取。', 'SNAPSHOT_ETAG_CONFLICT');
    }
    if (previous && Number(snapshot.sequence) > previous.sequence && eTag === previous.eTag) {
      fail('OneDrive 文件版本与资料版本不一致，已停止读取。', 'SNAPSHOT_ETAG_CONFLICT');
    }
    const persisted = JSON.stringify({
      version: 1,
      snapshot_id: snapshot.snapshot_id,
      sequence: Number(snapshot.sequence),
      e_tag: eTag,
    });
    try {
      storage.setItem(stateStorageKey(snapshot.snapshot_id), persisted);
    } catch (_error) {
      fail('当前浏览器无法保存 OneDrive 版本记录。', 'STORAGE_UNAVAILABLE');
    }
  }

  function createOneDriveClient(options = {}) {
    validateNoClientSecret(options);
    const storage = options.storage || globalThis.localStorage;
    const detected = options.pairing
      ? { active: true, pairing: validatePairing(options.pairing), clientId: options.clientId || options.pairing.clientId || '' }
      : options.detectPairing === false
        ? { active: false, pairing: null, clientId: '' }
        : detectOneDriveMode({ location: options.location, history: options.history });
    let pendingPairing = detected.pairing;
    const hint = readPairingHint(storage);
    const clientId = validateClientId(options.clientId || detected.clientId || pendingPairing?.clientId || hint?.clientId);
    const authProvider = resolveAuthProvider(options);
    const metadataProvider = typeof options.getMetadata === 'function' ? options.getMetadata : defaultGetMetadata;
    const receiptMetadataProvider = typeof options.getReceiptMetadata === 'function'
      ? options.getReceiptMetadata
      : defaultGetMetadata;
    const fetchImpl = options.fetch || globalThis.fetch?.bind(globalThis);
    const selectedCrypto = cryptoApi(options.crypto);
    const scopes = Array.isArray(options.scopes) && options.scopes.length
      ? Object.freeze(options.scopes.map((scope) => String(scope || '').trim()).filter(Boolean))
      : DEFAULT_SCOPES;

    if (typeof fetchImpl !== 'function') fail('当前浏览器无法读取 OneDrive 资料。', 'FETCH_UNAVAILABLE');
    if (!storage || typeof storage.getItem !== 'function' || typeof storage.setItem !== 'function') {
      fail('当前浏览器无法保存 OneDrive 防回滚记录。', 'STORAGE_UNAVAILABLE');
    }
    if (!scopes.length) fail('Microsoft 登录权限配置无效。', 'INVALID_AUTH_SCOPES');

    let currentSnapshot = null;
    let currentRemoteMetadata = null;
    let currentPairing = null;
    const queryApi = createSnapshotQueryApi(() => currentSnapshot);

    function storedPairingEnvelope() {
      let raw;
      try { raw = storage.getItem(PAIRING_STORAGE_KEY); } catch (_error) { fail('当前浏览器无法读取 OneDrive 配对信息。', 'STORAGE_UNAVAILABLE'); }
      if (!raw) return null;
      try {
        const parsed = JSON.parse(raw);
        return hasValidStoredSecretRecord(parsed) ? parsed : null;
      } catch (_error) {
        return null;
      }
    }

    function persistPairing(credentials, pin) {
      return wrapSecretRecord({
        version: 1,
        provider: 'onedrive',
        snapshotId: credentials.snapshotId,
        dataKey: credentials.dataKey,
        clientId,
      }, pin, selectedCrypto, PAIRING_AAD).then((wrapped) => {
        try {
          storage.setItem(PAIRING_STORAGE_KEY, JSON.stringify(wrapped));
          storage.setItem(PAIRING_HINT_STORAGE_KEY, JSON.stringify({
            version: 1,
            provider: 'onedrive',
            client_id: clientId,
          }));
        } catch (_error) {
          fail('当前浏览器无法安全保存 OneDrive 配对信息。', 'STORAGE_UNAVAILABLE');
        }
      });
    }

    async function loadPairing(pin) {
      if (pendingPairing) return pendingPairing;
      const stored = storedPairingEnvelope();
      if (!stored) fail('本手机尚未配对，请用电脑端 OneDrive 二维码重新打开。', 'PAIRING_REQUIRED', 401);
      let parsed;
      try {
        parsed = await unwrapSecretRecord(stored, pin, selectedCrypto, PAIRING_AAD);
        if (Number(parsed?.version) !== 1 || parsed?.provider !== 'onedrive') {
          fail('本机保存的 OneDrive 配对信息已损坏，请重新扫码配对。', 'INVALID_STORED_PAIRING');
        }
        const credentials = validatePairing(parsed);
        if (validateClientId(parsed.clientId) !== clientId) {
          fail('本机保存的 Microsoft 应用编号不一致，请重新扫码配对。', 'INVALID_STORED_PAIRING');
        }
        return credentials;
      } catch (error) {
        if (error instanceof CloudSnapshotError && error.code === 'INVALID_PAIRING') {
          fail('本机保存的 OneDrive 配对信息已损坏，请重新扫码配对。', 'INVALID_STORED_PAIRING');
        }
        throw error;
      }
    }

    async function acquireAccessToken() {
      try {
        return normalizeToken(await authProvider.acquireToken({ scopes: [...scopes] }));
      } catch (error) {
        if (error instanceof CloudSnapshotError) throw error;
        fail('Microsoft 登录未完成，请重新登录后再试。', 'ONEDRIVE_AUTH_FAILED', 401);
      }
    }

    async function loadMetadata(accessToken) {
      let result;
      try {
        result = await metadataProvider({
          accessToken,
          fetch: fetchImpl,
          graphUrl: GRAPH_METADATA_URL,
          fileName: SNAPSHOT_FILE_NAME,
          scopes: [...scopes],
        });
      } catch (error) {
        if (error instanceof CloudSnapshotError) throw error;
        fail('OneDrive 文件信息暂时无法读取。', 'METADATA_FETCH_FAILED');
      }
      let raw;
      try {
        raw = result && typeof result.ok === 'boolean'
          ? await parseMetadataResponse(result)
          : result;
      } catch (error) {
        if (error instanceof CloudSnapshotError) throw error;
        fail('OneDrive 文件信息暂时无法读取。', 'METADATA_FETCH_FAILED');
      }
      return sanitizeMetadata(raw);
    }

    async function downloadEnvelope(downloadUrl) {
      let response;
      try {
        // The URL is already preauthenticated by OneDrive. Deliberately omit
        // Authorization so a Graph bearer token can never reach that host.
        response = await fetchImpl(downloadUrl, {
          method: 'GET',
          cache: 'no-store',
          credentials: 'omit',
          redirect: 'follow',
          referrerPolicy: 'no-referrer',
        });
      } catch (_error) {
        fail('OneDrive 资料暂时无法下载。', 'DOWNLOAD_FAILED');
      }
      if (!response?.ok) {
        const status = Number(response?.status) || 0;
        fail(`OneDrive 资料暂时无法下载（${status}）。`, 'DOWNLOAD_FAILED', status);
      }
      try {
        return await readResponsePayload(response);
      } catch (error) {
        if (error instanceof CloudSnapshotError) throw error;
        fail('OneDrive 资料暂时无法读取。', 'DOWNLOAD_FAILED');
      }
    }

    async function login(pinValue) {
      currentSnapshot = null;
      currentRemoteMetadata = null;
      currentPairing = null;
      const pin = validatePin(pinValue);
      if (!authProvider.hasAccount()) {
        fail('请先登录 Microsoft 个人账号，再输入8位访问码。', 'MICROSOFT_SIGNIN_REQUIRED', 401);
      }
      let accessToken = '';
      try {
        accessToken = await acquireAccessToken();
        const pairing = await loadPairing(pin);
        const metadata = await loadMetadata(accessToken);
        const envelope = await downloadEnvelope(metadata.downloadUrl);
        const snapshot = await decryptSnapshot(envelope, pairing, selectedCrypto);
        await verifySnapshotPin(snapshot, pin, selectedCrypto);
        verifyAndStoreVersion(storage, snapshot, metadata.eTag);
        if (pendingPairing) {
          await persistPairing(pairing, pin);
          pendingPairing = null;
        }
        currentSnapshot = snapshot;
        currentRemoteMetadata = publicRemoteMetadata(metadata);
        currentPairing = pairing;
        return {
          authenticated: true,
          cloud: true,
          provider: 'onedrive',
          metadata: publicMetadata(snapshot),
          remote: { ...currentRemoteMetadata },
        };
      } finally {
        // The token is never copied to storage or retained as client state.
        accessToken = '';
      }
    }

    async function logout() {
      currentSnapshot = null;
      currentRemoteMetadata = null;
      currentPairing = null;
      try { await authProvider.clearSession(); } catch (_error) { /* Local sensitive state is already cleared. */ }
      return { loggedOut: true };
    }

    async function session() {
      return {
        authenticated: Boolean(currentSnapshot),
        cloud: true,
        provider: 'onedrive',
        metadata: publicMetadata(currentSnapshot),
        remote: currentRemoteMetadata ? { ...currentRemoteMetadata } : null,
      };
    }

    function hasPairing() {
      return Boolean(pendingPairing || storedPairingEnvelope());
    }

    async function microsoftLogin() {
      if (!authProvider.isAvailable()) {
        fail('Microsoft 登录组件尚未安装，请联系电脑端管理员完成手机网页配置。', 'MSAL_NOT_AVAILABLE');
      }
      try {
        await authProvider.signIn({ scopes: [...scopes] });
      } catch (error) {
        if (error instanceof CloudSnapshotError) throw error;
        fail('Microsoft 登录未完成，请重试。', 'ONEDRIVE_AUTH_FAILED', 401);
      }
      return { signedIn: Boolean(authProvider.hasAccount()), provider: 'onedrive' };
    }

    function microsoftStatus() {
      return {
        available: Boolean(authProvider.isAvailable()),
        signedIn: Boolean(authProvider.hasAccount()),
        clientId,
      };
    }

    async function receipts(filters = {}) {
      if (!currentSnapshot) fail('登录已过期，请重新输入访问码。', 'NOT_AUTHENTICATED', 401);
      const source = Array.isArray(currentSnapshot.payment_receipts) ? currentSnapshot.payment_receipts : [];
      if (source.length > 100000) fail('付款凭据索引记录过多。', 'RECORD_LIMIT_EXCEEDED');
      const personId = filters.personId == null || filters.personId === '' ? null : Number(filters.personId);
      const year = filters.year == null || filters.year === '' ? null : Number(filters.year);
      const month = filters.month == null || filters.month === '' ? null : Number(filters.month);
      return source.map(normalizeReceiptRecord)
        .filter((item) => personId == null || item.person_id === personId)
        .filter((item) => year == null || item.year === year)
        .filter((item) => month == null || item.month === month)
        .sort((left, right) => right.year - left.year || right.month - left.month
          || right.created_at.localeCompare(left.created_at) || left.display_name.localeCompare(right.display_name))
        .map((item) => ({ ...item }));
    }

    async function receiptAsset(assetIdValue) {
      if (!currentSnapshot || !currentPairing) {
        fail('登录已过期，请重新输入访问码。', 'NOT_AUTHENTICATED', 401);
      }
      const assetId = validateReceiptAssetId(assetIdValue);
      const indexRows = await receipts({});
      const receipt = indexRows.find((item) => item.asset_id === assetId);
      if (!receipt) fail('付款凭据不存在或已取消关联。', 'RECEIPT_NOT_FOUND', 404);
      if (!authProvider.hasAccount()) {
        fail('Microsoft 登录已失效，请重新登录。', 'MICROSOFT_SIGNIN_REQUIRED', 401);
      }
      let accessToken = '';
      let plaintextBlob;
      try {
        accessToken = await acquireAccessToken();
        const graphUrl = receiptAssetGraphMetadataUrl(assetId);
        let rawMetadata;
        try {
          const result = await receiptMetadataProvider({
            accessToken,
            fetch: fetchImpl,
            graphUrl,
            fileName: receiptAssetFileName(assetId),
            assetId,
            scopes: [...scopes],
          });
          rawMetadata = result && typeof result.ok === 'boolean'
            ? await parseMetadataResponse(result)
            : result;
        } catch (error) {
          if (error instanceof CloudSnapshotError) throw error;
          fail('OneDrive 付款凭据信息暂时无法读取。', 'RECEIPT_METADATA_FETCH_FAILED');
        }
        const metadata = sanitizeFileMetadata(
          rawMetadata,
          receiptAssetFileName(assetId),
          MAX_RECEIPT_ENVELOPE_BYTES,
        );
        let response;
        try {
          // OneDrive download URLs are already preauthenticated. Never attach
          // the Microsoft Graph bearer token to this separate host.
          response = await fetchImpl(metadata.downloadUrl, {
            method: 'GET',
            cache: 'no-store',
            credentials: 'omit',
            redirect: 'follow',
            referrerPolicy: 'no-referrer',
          });
        } catch (_error) {
          fail('OneDrive 付款凭据暂时无法下载。', 'RECEIPT_DOWNLOAD_FAILED');
        }
        const envelope = await readReceiptEnvelopeResponse(response);
        plaintextBlob = await decryptReceiptEnvelope(
          envelope,
          receipt,
          currentSnapshot.key_id,
          currentPairing.dataKey,
          selectedCrypto,
        );
        return {
          asset_id: receipt.asset_id,
          display_name: receipt.display_name,
          mime_type: receipt.mime_type,
          size_bytes: receipt.size_bytes,
          blob: plaintextBlob,
        };
      } finally {
        accessToken = '';
      }
    }

    return Object.freeze({
      session,
      login,
      logout,
      ...queryApi,
      hasPairing,
      microsoftLogin,
      microsoftStatus,
      receipts,
      receiptAsset,
    });
  }

  return Object.freeze({
    DEFAULT_SCOPES,
    GRAPH_METADATA_URL,
    PAIRING_HINT_STORAGE_KEY,
    PAIRING_STORAGE_KEY,
    SNAPSHOT_FILE_NAME,
    RECEIPT_ASSET_DIRECTORY,
    MAX_RECEIPT_BYTES,
    MAX_RECEIPT_ENVELOPE_BYTES,
    createOneDriveClient,
    detectOneDriveMode,
    hasStoredPairing,
    readPairingHint,
    receiptAssetGraphMetadataUrl,
  });
}));
