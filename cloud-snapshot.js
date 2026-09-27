(function attachPayrollCloudSnapshot(root, factory) {
  'use strict';

  if (typeof module === 'object' && module.exports) {
    const nodeCrypto = require('node:crypto');
    const nodeZlib = require('node:zlib');
    module.exports = factory({
      globalObject: globalThis,
      crypto: nodeCrypto.webcrypto,
      gunzip(bytes, maximumBytes) {
        return new Uint8Array(nodeZlib.gunzipSync(Buffer.from(bytes), { maxOutputLength: maximumBytes }));
      },
    });
    return;
  }

  root.PayrollCloudSnapshot = factory({
    globalObject: root,
    crypto: root.crypto,
  });
}(typeof globalThis !== 'undefined' ? globalThis : this, function buildPayrollCloudSnapshot(runtime) {
  'use strict';

  const CREDENTIAL_STORAGE_KEY = 'payroll-cloud.credentials.v1';
  const SEQUENCE_STORAGE_PREFIX = 'payroll-cloud.sequence.v1.';
  const CREDENTIAL_AAD = 'payroll-cloud-credentials:v1';
  const SNAPSHOT_FORMAT = 'personnel-payroll-readonly';
  const WRAP_ITERATIONS = 210000;
  const MAX_ENVELOPE_BYTES = 20 * 1024 * 1024;
  const MAX_SNAPSHOT_BYTES = 32 * 1024 * 1024;
  const RECORD_LIMITS = Object.freeze({
    groups: 10000,
    people: 100000,
    declarations: 250000,
    operation_logs: 100000,
  });

  class CloudSnapshotError extends Error {
    constructor(message, code = 'CLOUD_SNAPSHOT_ERROR', status = 0) {
      super(message);
      this.name = 'CloudSnapshotError';
      this.code = code;
      this.status = status;
    }
  }

  function fail(message, code, status) {
    throw new CloudSnapshotError(message, code, status);
  }

  function getTextEncoder() {
    const Encoder = runtime.globalObject.TextEncoder;
    if (!Encoder) fail('当前浏览器不支持安全解密，请更新浏览器。', 'UNSUPPORTED_BROWSER');
    return new Encoder();
  }

  function getTextDecoder() {
    const Decoder = runtime.globalObject.TextDecoder;
    if (!Decoder) fail('当前浏览器不支持安全解密，请更新浏览器。', 'UNSUPPORTED_BROWSER');
    return new Decoder('utf-8', { fatal: true });
  }

  function utf8(value) {
    return getTextEncoder().encode(String(value));
  }

  function decodeUtf8(bytes) {
    return getTextDecoder().decode(bytes);
  }

  function bytesToBase64Url(bytes) {
    let base64;
    if (typeof Buffer !== 'undefined') {
      base64 = Buffer.from(bytes).toString('base64');
    } else {
      let binary = '';
      const source = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      for (let offset = 0; offset < source.length; offset += 0x8000) {
        binary += String.fromCharCode(...source.subarray(offset, offset + 0x8000));
      }
      base64 = runtime.globalObject.btoa(binary);
    }
    return base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
  }

  function base64UrlToBytes(value, label = '编码内容') {
    const clean = String(value || '').trim();
    if (!clean || !/^[A-Za-z0-9_-]+$/.test(clean)) fail(`${label}格式无效。`, 'INVALID_ENCODING');
    const padded = clean.replace(/-/g, '+').replace(/_/g, '/') + '='.repeat((4 - (clean.length % 4)) % 4);
    try {
      if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(padded, 'base64'));
      const binary = runtime.globalObject.atob(padded);
      return Uint8Array.from(binary, (character) => character.charCodeAt(0));
    } catch (_error) {
      fail(`${label}格式无效。`, 'INVALID_ENCODING');
    }
  }

  function constantTimeEqual(left, right) {
    if (!(left instanceof Uint8Array) || !(right instanceof Uint8Array) || left.length !== right.length) return false;
    let difference = 0;
    for (let index = 0; index < left.length; index += 1) difference |= left[index] ^ right[index];
    return difference === 0;
  }

  function cryptoApi(customCrypto) {
    const selected = customCrypto || runtime.crypto;
    if (!selected?.subtle || typeof selected.getRandomValues !== 'function') {
      fail('当前环境不支持本地安全加密，请更新浏览器。', 'UNSUPPORTED_CRYPTO');
    }
    return selected;
  }

  function randomBytes(length, selectedCrypto) {
    const bytes = new Uint8Array(length);
    selectedCrypto.getRandomValues(bytes);
    return bytes;
  }

  async function derivePinBytes(pin, salt, iterations, selectedCrypto) {
    const material = await selectedCrypto.subtle.importKey('raw', utf8(pin), 'PBKDF2', false, ['deriveBits']);
    const bits = await selectedCrypto.subtle.deriveBits({
      name: 'PBKDF2',
      hash: 'SHA-256',
      salt,
      iterations,
    }, material, 256);
    return new Uint8Array(bits);
  }

  async function aesGcmDecrypt(ciphertext, rawKey, iv, aad, selectedCrypto, errorCode) {
    try {
      const key = await selectedCrypto.subtle.importKey('raw', rawKey, { name: 'AES-GCM' }, false, ['decrypt']);
      const plaintext = await selectedCrypto.subtle.decrypt({
        name: 'AES-GCM',
        iv,
        additionalData: aad,
        tagLength: 128,
      }, key, ciphertext);
      return new Uint8Array(plaintext);
    } catch (_error) {
      fail('加密资料校验失败，可能是链接不正确或数据已损坏。', errorCode || 'DECRYPT_FAILED');
    }
  }

  async function aesGcmEncrypt(plaintext, rawKey, iv, aad, selectedCrypto) {
    const key = await selectedCrypto.subtle.importKey('raw', rawKey, { name: 'AES-GCM' }, false, ['encrypt']);
    const ciphertext = await selectedCrypto.subtle.encrypt({
      name: 'AES-GCM',
      iv,
      additionalData: aad,
      tagLength: 128,
    }, key, plaintext);
    return new Uint8Array(ciphertext);
  }

  function cleanFragment(locationObject, historyObject) {
    const pathname = String(locationObject?.pathname || '/');
    const search = String(locationObject?.search || '');
    if (historyObject && typeof historyObject.replaceState === 'function') {
      try {
        historyObject.replaceState(historyObject.state || null, '', `${pathname}${search}`);
        return;
      } catch (_error) {
        // Fall through to clearing location.hash on unusual embedded browsers.
      }
    }
    if (locationObject) locationObject.hash = '';
  }

  function validatePairing(pairing) {
    const snapshotId = String(pairing?.snapshotId || '').trim();
    const readToken = String(pairing?.readToken || '').trim();
    const dataKey = String(pairing?.dataKey || '').trim();
    if (!/^[A-Za-z0-9_-]{16,128}$/.test(snapshotId)) fail('手机配对链接中的资料编号无效。', 'INVALID_PAIRING');
    if (readToken.length < 24 || readToken.length > 512 || /\s/.test(readToken)) fail('手机配对链接中的读取凭据无效。', 'INVALID_PAIRING');
    const keyBytes = base64UrlToBytes(dataKey, '手机配对密钥');
    if (keyBytes.length !== 32) fail('手机配对密钥长度无效。', 'INVALID_PAIRING');
    return { version: 1, snapshotId, readToken, dataKey };
  }

  /**
   * Detects a cloud pairing fragment and removes it from the address bar before
   * any asynchronous work begins. The returned credentials are memory-only.
   */
  function detectCloudMode(options = {}) {
    const locationObject = options.location || runtime.globalObject.location;
    const historyObject = options.history || runtime.globalObject.history;
    const hash = String(locationObject?.hash || '');
    if (!hash || hash === '#') return { active: false, pairing: null };

    const params = new URLSearchParams(hash.slice(1));
    const resemblesCloudPairing = ['v', 's', 'r', 'k'].some((key) => params.has(key));
    if (!resemblesCloudPairing) return { active: false, pairing: null };

    // Do this before validation so malformed or truncated secrets do not stay in
    // browser history or screenshots.
    cleanFragment(locationObject, historyObject);
    if (params.get('v') !== '1') fail('手机配对链接版本不受支持。', 'INVALID_PAIRING');
    const pairing = validatePairing({
      snapshotId: params.get('s'),
      readToken: params.get('r'),
      dataKey: params.get('k'),
    });
    return { active: true, pairing };
  }

  function sequenceStorageKey(snapshotId) {
    return `${SEQUENCE_STORAGE_PREFIX}${snapshotId}`;
  }

  function validatePin(pin) {
    const clean = String(pin || '').trim();
    if (!/^\d{8}$/.test(clean)) fail('请输入8位数字访问码。', 'INVALID_PIN_FORMAT', 400);
    return clean;
  }

  function validateStoredCredentials(value) {
    if (!value || typeof value !== 'object' || Number(value.version) !== 1) return false;
    if (value.kdf?.name !== 'PBKDF2' || value.kdf?.hash !== 'SHA-256') return false;
    if (!Number.isInteger(value.kdf?.iterations) || value.kdf.iterations < 100000 || value.kdf.iterations > 1000000) return false;
    try {
      return base64UrlToBytes(value.kdf.salt).length === 16
        && base64UrlToBytes(value.cipher?.iv).length === 12
        && base64UrlToBytes(value.ciphertext).length >= 17;
    } catch (_error) {
      return false;
    }
  }

  async function wrapSecretRecord(value, pin, selectedCrypto, aadLabel = CREDENTIAL_AAD) {
    const salt = randomBytes(16, selectedCrypto);
    const iv = randomBytes(12, selectedCrypto);
    const wrappingKey = await derivePinBytes(validatePin(pin), salt, WRAP_ITERATIONS, selectedCrypto);
    const plaintext = utf8(JSON.stringify(value));
    const ciphertext = await aesGcmEncrypt(plaintext, wrappingKey, iv, utf8(aadLabel), selectedCrypto);
    wrappingKey.fill(0);
    plaintext.fill(0);
    return {
      version: 1,
      kdf: {
        name: 'PBKDF2',
        hash: 'SHA-256',
        iterations: WRAP_ITERATIONS,
        salt: bytesToBase64Url(salt),
      },
      cipher: { name: 'AES-GCM', iv: bytesToBase64Url(iv) },
      ciphertext: bytesToBase64Url(ciphertext),
    };
  }

  async function unwrapSecretRecord(stored, pin, selectedCrypto, aadLabel = CREDENTIAL_AAD) {
    if (!validateStoredCredentials(stored)) fail('本机保存的配对信息已损坏，请重新扫码配对。', 'INVALID_STORED_PAIRING');
    const salt = base64UrlToBytes(stored.kdf.salt);
    const iv = base64UrlToBytes(stored.cipher.iv);
    const wrappingKey = await derivePinBytes(validatePin(pin), salt, stored.kdf.iterations, selectedCrypto);
    let plaintext;
    try {
      plaintext = await aesGcmDecrypt(
        base64UrlToBytes(stored.ciphertext),
        wrappingKey,
        iv,
        utf8(aadLabel),
        selectedCrypto,
        'WRONG_PIN',
      );
      return JSON.parse(decodeUtf8(plaintext));
    } catch (error) {
      if (error instanceof CloudSnapshotError && error.code === 'WRONG_PIN') {
        fail('访问码不正确，请重新输入。', 'WRONG_PIN', 401);
      }
      if (error instanceof SyntaxError) fail('本机保存的配对信息已损坏，请重新扫码配对。', 'INVALID_STORED_PAIRING');
      throw error;
    } finally {
      wrappingKey.fill(0);
      if (plaintext) plaintext.fill(0);
    }
  }

  async function wrapCredentials(credentials, pin, selectedCrypto) {
    return wrapSecretRecord({
      version: 1,
      snapshotId: credentials.snapshotId,
      readToken: credentials.readToken,
      dataKey: credentials.dataKey,
    }, pin, selectedCrypto, CREDENTIAL_AAD);
  }

  async function unwrapCredentials(stored, pin, selectedCrypto) {
    try {
      const parsed = await unwrapSecretRecord(stored, pin, selectedCrypto, CREDENTIAL_AAD);
      return validatePairing(parsed);
    } catch (error) {
      if (error instanceof CloudSnapshotError && error.code === 'INVALID_PAIRING') {
        fail('本机保存的配对信息已损坏，请重新扫码配对。', 'INVALID_STORED_PAIRING');
      }
      throw error;
    }
  }

  async function readResponsePayload(response) {
    const lengthHeader = response?.headers?.get?.('content-length');
    if (lengthHeader && Number(lengthHeader) > MAX_ENVELOPE_BYTES) fail('云端资料包过大，已停止读取。', 'ENVELOPE_TOO_LARGE');
    let text;
    if (typeof response.text === 'function') {
      text = await response.text();
    } else if (typeof response.json === 'function') {
      return response.json();
    } else {
      fail('云端返回格式无效。', 'INVALID_RESPONSE');
    }
    if (utf8(text).byteLength > MAX_ENVELOPE_BYTES) fail('云端资料包过大，已停止读取。', 'ENVELOPE_TOO_LARGE');
    try {
      return JSON.parse(text);
    } catch (_error) {
      fail('云端返回的资料包不是有效JSON。', 'INVALID_RESPONSE');
    }
  }

  async function gunzipBytes(compressed) {
    if (runtime.gunzip) {
      try {
        return runtime.gunzip(compressed, MAX_SNAPSHOT_BYTES);
      } catch (_error) {
        fail('云端资料解压失败或内容过大。', 'DECOMPRESSION_FAILED');
      }
    }
    const Decompression = runtime.globalObject.DecompressionStream;
    const BlobConstructor = runtime.globalObject.Blob;
    if (!Decompression || !BlobConstructor) fail('当前浏览器不支持安全资料解压，请更新浏览器。', 'UNSUPPORTED_BROWSER');
    try {
      const reader = new BlobConstructor([compressed]).stream().pipeThrough(new Decompression('gzip')).getReader();
      const chunks = [];
      let total = 0;
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        total += value.byteLength;
        if (total > MAX_SNAPSHOT_BYTES) {
          await reader.cancel();
          fail('云端资料解压后过大，已停止读取。', 'SNAPSHOT_TOO_LARGE');
        }
        chunks.push(value);
      }
      const result = new Uint8Array(total);
      let offset = 0;
      for (const chunk of chunks) {
        result.set(chunk, offset);
        offset += chunk.byteLength;
      }
      return result;
    } catch (error) {
      if (error instanceof CloudSnapshotError) throw error;
      fail('云端资料解压失败或内容已损坏。', 'DECOMPRESSION_FAILED');
    }
  }

  function snapshotAad(envelope) {
    return [
      'personnel-payroll-cloud-snapshot',
      `v=${Number(envelope.version)}`,
      `snapshot=${envelope.snapshot_id}`,
      `source=${envelope.source_instance_id}`,
      `sequence=${Number(envelope.sequence)}`,
      `generated=${envelope.generated_at}`,
      `key=${envelope.key_id}`,
    ].join('|');
  }

  function validateEnvelope(rawEnvelope, snapshotId) {
    const envelope = rawEnvelope?.envelope && typeof rawEnvelope.envelope === 'object'
      ? rawEnvelope.envelope
      : rawEnvelope?.data?.envelope && typeof rawEnvelope.data.envelope === 'object'
        ? rawEnvelope.data.envelope
        : rawEnvelope?.data && typeof rawEnvelope.data === 'object'
          ? rawEnvelope.data
          : rawEnvelope;
    if (!envelope || typeof envelope !== 'object' || Array.isArray(envelope)) fail('云端资料包格式无效。', 'INVALID_ENVELOPE');
    if (Number(envelope.version) !== 1
      || envelope.format !== SNAPSHOT_FORMAT
      || Number(envelope.schema_version) !== 1
      || envelope.algorithm !== 'AES-256-GCM'
      || envelope.compression !== 'gzip') {
      fail('云端资料包版本或加密格式不受支持。', 'INVALID_ENVELOPE');
    }
    if (String(envelope.snapshot_id || '') !== snapshotId) fail('云端资料编号与配对信息不一致。', 'INVALID_ENVELOPE');
    if (!/^[A-Za-z0-9_-]{12,128}$/.test(String(envelope.source_instance_id || ''))
      || !/^[A-Za-z0-9_-]{8,128}$/.test(String(envelope.key_id || ''))
      || !envelope.generated_at
      || Number.isNaN(Date.parse(envelope.generated_at))) {
      fail('云端资料包元数据无效。', 'INVALID_ENVELOPE');
    }
    const sequence = Number(envelope.sequence);
    if (!Number.isSafeInteger(sequence) || sequence < 1) fail('云端资料版本号无效。', 'INVALID_ENVELOPE');
    const iv = base64UrlToBytes(envelope.iv, '云端资料随机量');
    const ciphertext = base64UrlToBytes(envelope.ciphertext, '云端加密资料');
    if (iv.length !== 12 || ciphertext.length < 17 || ciphertext.length > MAX_ENVELOPE_BYTES) {
      fail('云端资料包长度无效。', 'INVALID_ENVELOPE');
    }
    return { envelope, sequence, iv, ciphertext, aad: utf8(snapshotAad(envelope)) };
  }

  function validateSnapshot(snapshot, envelope) {
    if (!snapshot || typeof snapshot !== 'object' || Array.isArray(snapshot)) fail('解密后的资料格式无效。', 'INVALID_SNAPSHOT');
    if (snapshot.format !== SNAPSHOT_FORMAT || Number(snapshot.schema_version) !== 1) fail('资料快照版本不受支持。', 'INVALID_SNAPSHOT');
    if (String(snapshot.snapshot_id || '') !== envelope.snapshot_id
      || String(snapshot.source_instance_id || '') !== envelope.source_instance_id
      || String(snapshot.key_id || '') !== envelope.key_id
      || Number(snapshot.sequence) !== Number(envelope.sequence)
      || String(snapshot.generated_at || '') !== envelope.generated_at) {
      fail('资料快照的编号或版本不一致。', 'INVALID_SNAPSHOT');
    }
    if (!snapshot.generated_at || Number.isNaN(Date.parse(snapshot.generated_at))) fail('资料快照缺少有效同步时间。', 'INVALID_SNAPSHOT');
    for (const [field, maximum] of Object.entries(RECORD_LIMITS)) {
      if (!Array.isArray(snapshot[field])) fail(`资料快照缺少${field}列表。`, 'INVALID_SNAPSHOT');
      if (snapshot[field].length > maximum) fail(`资料快照中的${field}记录过多。`, 'RECORD_LIMIT_EXCEEDED');
      for (const row of snapshot[field]) {
        if (!row || typeof row !== 'object' || Array.isArray(row)) fail(`资料快照中的${field}记录格式无效。`, 'INVALID_SNAPSHOT');
      }
    }
    const totalRecords = Object.keys(RECORD_LIMITS).reduce((sum, field) => sum + snapshot[field].length, 0);
    if (totalRecords > 350000) fail('资料快照总记录数过多。', 'RECORD_LIMIT_EXCEEDED');
    return snapshot;
  }

  async function verifySnapshotPin(snapshot, pin, selectedCrypto) {
    const auth = snapshot.auth;
    if (!auth || auth.method !== 'PBKDF2-SHA-256') fail('资料快照缺少访问码校验信息。', 'INVALID_SNAPSHOT_AUTH');
    const iterations = Number(auth.iterations);
    if (!Number.isInteger(iterations) || iterations < 10000 || iterations > 1000000) fail('资料快照访问码参数无效。', 'INVALID_SNAPSHOT_AUTH');
    const salt = base64UrlToBytes(auth.salt, '访问码盐值');
    const expected = base64UrlToBytes(auth.hash, '访问码校验值');
    if (salt.length < 16 || salt.length > 64 || expected.length !== 32) fail('资料快照访问码参数无效。', 'INVALID_SNAPSHOT_AUTH');
    const actual = await derivePinBytes(pin, salt, iterations, selectedCrypto);
    const valid = constantTimeEqual(actual, expected);
    actual.fill(0);
    if (!valid) fail('访问码不正确，请重新输入。', 'WRONG_PIN', 401);
  }

  async function decryptSnapshot(envelopePayload, credentials, selectedCrypto) {
    const validated = validateEnvelope(envelopePayload, credentials.snapshotId);
    const rawKey = base64UrlToBytes(credentials.dataKey, '手机配对密钥');
    if (rawKey.length !== 32) fail('手机配对密钥长度无效。', 'INVALID_PAIRING');
    let compressed;
    let decompressed;
    try {
      compressed = await aesGcmDecrypt(
        validated.ciphertext,
        rawKey,
        validated.iv,
        validated.aad,
        selectedCrypto,
        'SNAPSHOT_DECRYPT_FAILED',
      );
      decompressed = await gunzipBytes(compressed);
      if (decompressed.byteLength > MAX_SNAPSHOT_BYTES) fail('解密后的资料过大，已停止读取。', 'SNAPSHOT_TOO_LARGE');
      let snapshot;
      try {
        snapshot = JSON.parse(decodeUtf8(decompressed));
      } catch (_error) {
        fail('解密后的资料不是有效JSON。', 'INVALID_SNAPSHOT');
      }
      return validateSnapshot(snapshot, validated.envelope);
    } finally {
      rawKey.fill(0);
      if (compressed) compressed.fill(0);
      if (decompressed) decompressed.fill(0);
    }
  }

  function pinyinCollator() {
    try {
      return new Intl.Collator('zh-CN-u-co-pinyin', { sensitivity: 'base', numeric: true });
    } catch (_error) {
      return new Intl.Collator('zh-CN', { sensitivity: 'base', numeric: true });
    }
  }

  const collator = pinyinCollator();
  const groupKindRank = new Map([['分类', 0], ['临时项目用工', 1], ['包工头', 2]]);

  function groupCompare(left, right) {
    if (left.kind !== right.kind) return (groupKindRank.get(left.kind) ?? 99) - (groupKindRank.get(right.kind) ?? 99);
    return collator.compare(String(left.name || ''), String(right.name || '')) || Number(left.id || 0) - Number(right.id || 0);
  }

  function personCompare(left, right) {
    if (left.group_kind !== right.group_kind) return (groupKindRank.get(left.group_kind) ?? 99) - (groupKindRank.get(right.group_kind) ?? 99);
    return collator.compare(String(left.group_name || ''), String(right.group_name || ''))
      || collator.compare(String(left.source_sort_key || left.legal_name || left.display_name || ''), String(right.source_sort_key || right.legal_name || right.display_name || ''))
      || collator.compare(String(left.display_name || ''), String(right.display_name || ''))
      || Number(left.id || left.person_id || 0) - Number(right.id || right.person_id || 0);
  }

  function cleanLeaderPrefix(value) {
    return String(value || '').trim().replace(/^bgt(?:[\s_:\-：]*)/i, '').trim();
  }

  function safeNumber(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number : 0;
  }

  function personIdOf(row) {
    return Number(row.person_id ?? row.id);
  }

  function declarationYear(row) {
    if (Number.isInteger(Number(row.year))) return Number(row.year);
    const match = String(row.month_key || '').match(/^(\d{4})-(\d{2})$/);
    return match ? Number(match[1]) : 0;
  }

  function declarationMonth(row) {
    if (Number.isInteger(Number(row.month))) return Number(row.month);
    const match = String(row.month_key || '').match(/^(\d{4})-(\d{2})$/);
    return match ? Number(match[2]) : 0;
  }

  function publicMetadata(snapshot) {
    if (!snapshot) return null;
    return {
      schema_version: snapshot.schema_version,
      snapshot_id: snapshot.snapshot_id,
      sequence: snapshot.sequence,
      generated_at: snapshot.generated_at,
      group_count: snapshot.groups.length,
      people_count: snapshot.people.length,
      declaration_count: snapshot.declarations.length,
      log_count: snapshot.operation_logs.length,
    };
  }

  /**
   * Builds the read-only query surface around an in-memory decrypted snapshot.
   * Transport adapters (the existing Worker client and the OneDrive client)
   * share this implementation so field filtering and sorting cannot drift.
   */
  function createSnapshotQueryApi(snapshotProvider) {
    if (typeof snapshotProvider !== 'function') fail('资料读取器配置无效。', 'INVALID_SNAPSHOT_PROVIDER');

    function requireSnapshot() {
      const snapshot = snapshotProvider();
      if (!snapshot) fail('登录已过期，请重新输入访问码。', 'NOT_AUTHENTICATED', 401);
      return snapshot;
    }

    async function groups() {
      const snapshot = requireSnapshot();
      return snapshot.groups.map((group) => ({ ...group })).sort(groupCompare);
    }

    async function searchPeople(query, year) {
      const snapshot = requireSnapshot();
      const needle = String(query || '').trim().toLocaleLowerCase('zh-CN');
      if (!needle) return [];
      const declarationsByPerson = new Map();
      for (const declaration of snapshot.declarations) {
        if (declarationYear(declaration) !== Number(year)) continue;
        const id = personIdOf(declaration);
        declarationsByPerson.set(id, (declarationsByPerson.get(id) || 0) + safeNumber(declaration.amount));
      }
      return snapshot.people
        .filter((person) => Number(person.active) !== 0)
        .filter((person) => [
          person.legal_name,
          person.display_name,
          person.id_number,
          person.phone,
          person.bank_number,
          person.group_name,
        ].some((value) => String(value || '').toLocaleLowerCase('zh-CN').includes(needle)))
        .map((person) => ({
          id: person.id,
          person_code: person.person_code,
          legal_name: cleanLeaderPrefix(person.legal_name),
          display_name: cleanLeaderPrefix(person.display_name),
          is_group_leader: Number(person.is_group_leader) ? 1 : 0,
          is_material_manual_payment: Number(person.is_material_manual_payment) ? 1 : 0,
          group_kind: person.group_kind,
          group_name: person.group_name,
          phone: person.phone,
          id_number: person.id_number,
          bank_number: person.bank_number,
          issue_count: person.issue_count,
          annual_total: declarationsByPerson.get(Number(person.id)) || 0,
        }))
        .sort(personCompare);
    }

    async function person(personId) {
      const snapshot = requireSnapshot();
      const source = snapshot.people.find((item) => Number(item.id) === Number(personId));
      if (!source) fail('人员不存在。', 'PERSON_NOT_FOUND', 404);
      const salaryHistory = snapshot.declarations
        .filter((row) => personIdOf(row) === Number(personId))
        .map((row) => ({
          year: declarationYear(row),
          month: declarationMonth(row),
          amount: safeNumber(row.amount),
          status: row.status,
          source: row.source,
          updated_at: row.updated_at,
          receipt_count: Math.max(0, Number(row.receipt_count) || 0),
        }))
        .sort((left, right) => right.year - left.year || right.month - left.month);
      return {
        ...source,
        legal_name: cleanLeaderPrefix(source.legal_name),
        display_name: cleanLeaderPrefix(source.display_name),
        salary_history: salaryHistory,
        salary_trend: [...salaryHistory].reverse(),
      };
    }

    async function declarations(filters = {}) {
      const snapshot = requireSnapshot();
      const year = Number(filters.year);
      const month = Number(filters.month);
      const groupId = filters.groupId === '' || filters.groupId == null ? null : Number(filters.groupId);
      const needle = String(filters.q || '').trim().toLocaleLowerCase('zh-CN');
      const peopleById = new Map(snapshot.people.map((item) => [Number(item.id), item]));
      const rows = snapshot.declarations
        .filter((row) => declarationYear(row) === year && declarationMonth(row) === month)
        .filter((row) => row.status !== '未申报' || safeNumber(row.amount) !== 0)
        .map((row) => ({ person: peopleById.get(personIdOf(row)), declaration: row }))
        .filter(({ person: item }) => item)
        .filter(({ person: item }) => groupId == null || Number(item.group_id) === groupId)
        .filter(({ person: item }) => !needle || [item.legal_name, item.display_name, item.id_number, item.phone, item.group_name]
          .some((value) => String(value || '').toLocaleLowerCase('zh-CN').includes(needle)))
        .map(({ person: item, declaration }) => ({
          ...declaration,
          person_id: item.id,
          person_code: item.person_code,
          legal_name: cleanLeaderPrefix(item.legal_name),
          display_name: cleanLeaderPrefix(item.display_name),
          is_group_leader: Number(item.is_group_leader) ? 1 : 0,
          is_material_manual_payment: Number(item.is_material_manual_payment) ? 1 : 0,
          id_number: item.id_number,
          phone: item.phone,
          bank_number: item.bank_number,
          group_id: item.group_id,
          group_kind: item.group_kind,
          group_name: item.group_name,
          year,
          month,
          month_key: `${year}-${String(month).padStart(2, '0')}`,
          amount: safeNumber(declaration.amount),
        }))
        .sort(personCompare);
      return {
        rows,
        summary: {
          month_key: `${year}-${String(month).padStart(2, '0')}`,
          people_count: rows.length,
          total_amount: rows.reduce((sum, row) => sum + safeNumber(row.amount), 0),
        },
      };
    }

    async function stats(yearValue, monthValue) {
      const snapshot = requireSnapshot();
      const year = Number(yearValue);
      const month = Number(monthValue);
      const activePeople = snapshot.people.filter((item) => Number(item.active) !== 0);
      const activeIds = new Set(activePeople.map((item) => Number(item.id)));
      const monthRows = snapshot.declarations.filter((row) => activeIds.has(personIdOf(row))
        && declarationYear(row) === year && declarationMonth(row) === month);
      const monthly = Array.from({ length: 12 }, (_, index) => {
        const selectedMonth = index + 1;
        const rows = snapshot.declarations.filter((row) => activeIds.has(personIdOf(row))
          && declarationYear(row) === year && declarationMonth(row) === selectedMonth);
        return {
          month: selectedMonth,
          total_amount: rows.reduce((sum, row) => sum + safeNumber(row.amount), 0),
          paid_people: new Set(rows.filter((row) => safeNumber(row.amount) > 0).map(personIdOf)).size,
        };
      });
      const groups = snapshot.groups.map((group) => {
        const members = activePeople.filter((item) => Number(item.group_id) === Number(group.id));
        const memberIds = new Set(members.map((item) => Number(item.id)));
        const selectedRows = monthRows.filter((row) => memberIds.has(personIdOf(row)));
        return {
          id: group.id,
          kind: group.kind,
          name: group.name,
          people_count: members.length,
          paid_people: new Set(selectedRows.filter((row) => safeNumber(row.amount) > 0).map(personIdOf)).size,
          total_amount: selectedRows.reduce((sum, row) => sum + safeNumber(row.amount), 0),
        };
      }).filter((group) => group.people_count > 0)
        .sort((left, right) => right.total_amount - left.total_amount || groupCompare(left, right));
      return {
        summary: {
          total_amount: monthRows.reduce((sum, row) => sum + safeNumber(row.amount), 0),
          paid_people: new Set(monthRows.filter((row) => safeNumber(row.amount) > 0).map(personIdOf)).size,
          active_people: activePeople.length,
          group_count: new Set(activePeople.map((item) => item.group_id).filter((value) => value != null)).size,
        },
        groups,
        monthly,
      };
    }

    async function logs(filters = {}) {
      const snapshot = requireSnapshot();
      const needle = String(filters.q || '').trim().toLocaleLowerCase('zh-CN');
      const limit = Math.min(Math.max(Number(filters.limit) || 100, 1), 2000);
      return snapshot.operation_logs
        .filter((log) => !filters.module || log.module === filters.module)
        .filter((log) => !filters.actionType || log.action_type === filters.actionType)
        .filter((log) => !needle || [log.target_name, log.detail, log.operation_code, log.batch_code]
          .some((value) => String(value || '').toLocaleLowerCase('zh-CN').includes(needle)))
        .slice()
        .sort((left, right) => Number(right.id || 0) - Number(left.id || 0)
          || String(right.created_at || '').localeCompare(String(left.created_at || '')))
        .slice(0, limit)
        .map((log) => ({ ...log }));
    }

    async function metadata() {
      return publicMetadata(requireSnapshot());
    }

    return Object.freeze({ groups, searchPeople, person, declarations, stats, logs, metadata });
  }

  function createCloudClient(options = {}) {
    const storage = options.storage || runtime.globalObject.localStorage;
    const fetchImpl = options.fetch || runtime.globalObject.fetch?.bind(runtime.globalObject);
    const selectedCrypto = cryptoApi(options.crypto);
    if (!storage || typeof storage.getItem !== 'function' || typeof storage.setItem !== 'function') {
      fail('当前浏览器无法安全保存手机配对信息。', 'STORAGE_UNAVAILABLE');
    }
    if (typeof fetchImpl !== 'function') fail('当前浏览器无法读取云端资料。', 'FETCH_UNAVAILABLE');

    const detected = options.pairing
      ? { active: true, pairing: validatePairing(options.pairing) }
      : detectCloudMode({ location: options.location, history: options.history });
    let pendingPairing = detected.pairing;
    let currentSnapshot = null;

    function storedPairingEnvelope() {
      const raw = storage.getItem(CREDENTIAL_STORAGE_KEY);
      if (!raw) return null;
      try {
        const parsed = JSON.parse(raw);
        return validateStoredCredentials(parsed) ? parsed : null;
      } catch (_error) {
        return null;
      }
    }

    function hasPairing() {
      return Boolean(pendingPairing || storedPairingEnvelope());
    }

    async function loadCredentials(pin) {
      if (pendingPairing) return pendingPairing;
      const stored = storedPairingEnvelope();
      if (!stored) fail('尚未配对，请先使用电脑端生成的二维码打开手机查询。', 'PAIRING_REQUIRED', 401);
      return unwrapCredentials(stored, pin, selectedCrypto);
    }

    async function fetchEnvelope(credentials) {
      const baseUrl = String(options.baseUrl || '').replace(/\/$/, '');
      const response = await fetchImpl(`${baseUrl}/api/snapshot/${encodeURIComponent(credentials.snapshotId)}`, {
        method: 'GET',
        cache: 'no-store',
        credentials: 'omit',
        headers: {
          Accept: 'application/json',
          Authorization: `Bearer ${credentials.readToken}`,
        },
      });
      if (!response?.ok) {
        if (Number(response?.status) === 401 || Number(response?.status) === 403 || Number(response?.status) === 404) {
          fail('手机配对已失效，请在电脑端重新生成配对二维码。', 'PAIRING_REJECTED', Number(response.status));
        }
        fail(`云端资料暂时无法读取（${Number(response?.status) || 0}）。`, 'FETCH_FAILED', Number(response?.status) || 0);
      }
      return readResponsePayload(response);
    }

    function rejectRollback(snapshot) {
      const key = sequenceStorageKey(snapshot.snapshot_id);
      const highest = Number(storage.getItem(key) || 0);
      if (Number.isSafeInteger(highest) && snapshot.sequence < highest) {
        fail('检测到旧版本资料，已拒绝打开以防数据回退。', 'SNAPSHOT_ROLLBACK');
      }
      storage.setItem(key, String(Math.max(highest || 0, snapshot.sequence)));
    }

    async function login(pinValue) {
      currentSnapshot = null;
      const pin = validatePin(pinValue);
      const credentials = await loadCredentials(pin);
      const envelope = await fetchEnvelope(credentials);
      const snapshot = await decryptSnapshot(envelope, credentials, selectedCrypto);
      await verifySnapshotPin(snapshot, pin, selectedCrypto);
      rejectRollback(snapshot);

      if (pendingPairing) {
        const wrapped = await wrapCredentials(credentials, pin, selectedCrypto);
        storage.setItem(CREDENTIAL_STORAGE_KEY, JSON.stringify(wrapped));
        pendingPairing = null;
      }
      currentSnapshot = snapshot;
      return { authenticated: true, cloud: true, metadata: publicMetadata(snapshot) };
    }

    async function logout() {
      currentSnapshot = null;
      return { loggedOut: true };
    }

    async function session() {
      return { authenticated: Boolean(currentSnapshot), cloud: true, metadata: publicMetadata(currentSnapshot) };
    }
    const queryApi = createSnapshotQueryApi(() => currentSnapshot);

    return Object.freeze({
      session,
      login,
      logout,
      ...queryApi,
      hasPairing,
    });
  }

  return Object.freeze({
    CloudSnapshotError,
    detectCloudMode,
    createCloudClient,
    __internal: Object.freeze({
      base64UrlToBytes,
      createSnapshotQueryApi,
      cryptoApi,
      decryptSnapshot,
      publicMetadata,
      readResponsePayload,
      cleanFragment,
      hasValidStoredSecretRecord: validateStoredCredentials,
      unwrapSecretRecord,
      validatePin,
      verifySnapshotPin,
      wrapSecretRecord,
    }),
  });
}));
