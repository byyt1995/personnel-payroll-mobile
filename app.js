'use strict';

const PAGE_PARAMS = new URL(window.location.href).searchParams;
const DEMO_MODE = PAGE_PARAMS.get('demo') === '1';
const DEMO_SCREEN = PAGE_PARAMS.get('screen') || 'login';
const HAS_ONEDRIVE_FRAGMENT = (() => {
  const params = new URLSearchParams(String(window.location.hash || '').replace(/^#/, ''));
  return String(params.get('provider') || '').toLowerCase() === 'onedrive';
})();
let oneDriveStartupError = null;
let oneDriveDetection = { active: false, pairing: null, clientId: '' };
if (!DEMO_MODE && HAS_ONEDRIVE_FRAGMENT) {
  try {
    oneDriveDetection = window.PayrollOneDriveSnapshot.detectOneDriveMode({
      location: window.location,
      history: window.history,
    });
  } catch (error) {
    oneDriveStartupError = error;
  }
}
const oneDriveHint = !DEMO_MODE ? window.PayrollOneDriveSnapshot.readPairingHint(window.localStorage) : null;
const ONEDRIVE_MODE = !DEMO_MODE && (HAS_ONEDRIVE_FRAGMENT
  || Boolean(oneDriveHint)
  || window.PayrollOneDriveSnapshot.hasStoredPairing(window.localStorage));
const LEGACY_CLOUD_MODE = !DEMO_MODE && !ONEDRIVE_MODE && (PAGE_PARAMS.get('cloud') === '1'
  || (window.location.protocol === 'https:' && !window.location.hostname.endsWith('.ts.net')));
const CLOUD_MODE = ONEDRIVE_MODE || LEGACY_CLOUD_MODE;
const CURRENT_DATE = new Date();
const DEFAULT_YEAR = CURRENT_DATE.getFullYear();
const DEFAULT_MONTH = CURRENT_DATE.getMonth() + 1;
const CLOUD_BACKGROUND_LOCK_MS = 15 * 60 * 1000;
const MAX_RECEIPT_PREVIEW_BYTES = 25 * 1024 * 1024;
let cloudHiddenAt = 0;
let cloudLockTimer = null;
let activeReceiptObjectUrl = '';
let receiptRequestSequence = 0;

const state = {
  authenticated: false,
  currentView: 'home',
  returnView: 'home',
  groups: [],
  loaded: { stats: false, logs: false },
  retryActions: new Map(),
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

class ApiError extends Error {
  constructor(message, status = 0) {
    super(message);
    this.name = 'ApiError';
    this.status = status;
  }
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function safeNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function formatMoney(value) {
  return safeNumber(value).toLocaleString('zh-CN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatInteger(value) {
  return Math.max(0, Math.round(safeNumber(value))).toLocaleString('zh-CN');
}

function displayValue(value, fallback = '未填写') {
  const text = String(value ?? '').trim();
  return text || fallback;
}

function cleanPersonName(value) {
  return String(value || '').trim().replace(/^bgt(?:[\s_:\-：]*)/i, '').trim();
}

function isGroupLeader(person) {
  return Boolean(Number(person?.is_group_leader)
    || /^bgt(?:[\s_:\-：]*)/i.test(String(person?.legal_name || '').trim())
    || /^bgt(?:[\s_:\-：]*)/i.test(String(person?.display_name || '').trim()));
}

function mobilePersonName(person, fallback = '未填写姓名') {
  return cleanPersonName(person?.display_name || person?.legal_name) || fallback;
}

function leaderBadge(person) {
  return isGroupLeader(person) ? '<span class="leader-badge">包工头本人</span>' : '';
}

function materialPaymentBadge(person) {
  return Number(person?.is_material_manual_payment)
    ? '<span class="material-payment-badge">材料款人工代发</span>' : '';
}

function formatDateTime(value) {
  if (!value) return '时间未记录';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleString('zh-CN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  });
}

function formatMonth(year, month) {
  return `${Number(year)}-${String(Number(month)).padStart(2, '0')}`;
}

function initialFor(person) {
  const name = mobilePersonName(person, '人');
  return name.slice(-1) || '人';
}

function normalizedError(error, fallback = '读取失败，请稍后重试') {
  if (!error) return fallback;
  if (error instanceof TypeError && /fetch/i.test(error.message)) {
    return CLOUD_MODE ? '无法连接云端，请确认手机网络正常后重试。' : '无法连接电脑端，请确认电脑端服务已开启且手机网络正常。';
  }
  return error.message || fallback;
}

async function request(path, options = {}) {
  const requestOptions = {
    method: options.method || 'GET',
    credentials: 'same-origin',
    cache: 'no-store',
    headers: {
      Accept: 'application/json',
      ...(options.body ? { 'Content-Type': 'application/json' } : {}),
      ...(options.headers || {}),
    },
    ...(options.body ? { body: JSON.stringify(options.body) } : {}),
  };

  const response = await fetch(path, requestOptions);
  const contentType = response.headers.get('content-type') || '';
  let payload = null;
  if (contentType.includes('application/json')) {
    payload = await response.json().catch(() => null);
  } else {
    const text = await response.text().catch(() => '');
    payload = text ? { error: text } : null;
  }

  if (!response.ok || payload?.ok === false) {
    const message = payload?.error || payload?.message || (response.status === 401 ? '访问码无效或登录已过期' : `请求失败（${response.status}）`);
    throw new ApiError(message, response.status);
  }

  if (payload && Object.prototype.hasOwnProperty.call(payload, 'data')) return payload.data;
  return payload;
}

function makeQuery(params) {
  const search = new URLSearchParams();
  Object.entries(params).forEach(([key, value]) => {
    if (value !== '' && value !== null && value !== undefined) search.set(key, String(value));
  });
  const query = search.toString();
  return query ? `?${query}` : '';
}

const demoGroups = [
  { id: 1, kind: '包工头', name: '张建国', person_count: 2 },
  { id: 2, kind: '包工头', name: '刘师傅班组', person_count: 2 },
  { id: 3, kind: '分类', name: '点工', person_count: 1 },
];

const demoPeople = [
  {
    id: 101, person_code: 'P000101', group_id: 1, group_kind: '包工头', group_name: '张建国', is_group_leader: 1,
    legal_name: '张建国', display_name: '张建国', gender: '男', id_number: '310101197406120018',
    phone: '13800001021', bank_number: '6222023100001021001', bank_name: '中国工商银行',
    bank_branch: '上海海复路支行', bank_branch_code: '102290012345', address: '上海市宝山区示例路18号',
    birth_date: '1974-06-12', issuing_authority: '上海市公安局宝山分局', id_valid_period: '2019-06-12 至 2039-06-12',
    notes: '本人为该班组包工头，联系时优先使用手机。', active: 1,
    salary_history: [
      { year: 2026, month: 9, amount: 8650, status: '已申报', source: '月度申报导入', updated_at: '2026-09-22T09:18:00+08:00', receipt_count: 2 },
      { year: 2026, month: 8, amount: 7920, status: '已申报', source: '月度申报导入', updated_at: '2026-08-23T10:03:00+08:00' },
      { year: 2026, month: 7, amount: 8350, status: '已申报', source: '月度申报导入', updated_at: '2026-07-24T11:08:00+08:00' },
      { year: 2026, month: 6, amount: 7680, status: '已申报', source: '月度申报导入', updated_at: '2026-06-24T08:42:00+08:00' },
      { year: 2026, month: 5, amount: 8100, status: '已申报', source: '月度申报导入', updated_at: '2026-05-23T15:11:00+08:00' },
      { year: 2026, month: 4, amount: 7420, status: '已申报', source: '月度申报导入', updated_at: '2026-04-22T14:05:00+08:00' },
    ],
  },
  {
    id: 102, person_code: 'P000102', group_id: 1, group_kind: '包工头', group_name: '张建国', is_group_leader: 0,
    legal_name: '李秀兰', display_name: '李秀兰', gender: '女', id_number: '320102197809260027',
    phone: '13900001036', bank_number: '6217003100001036008', bank_name: '中国建设银行',
    bank_branch: '上海共和新路支行', bank_branch_code: '105290077777', address: '上海市静安区演示路6号',
    birth_date: '1978-09-26', issuing_authority: '南京市公安局玄武分局', id_valid_period: '2020-09-26 至 2040-09-26',
    notes: '工资卡信息已于2026年3月核对。', active: 1,
    salary_history: [
      { year: 2026, month: 9, amount: 7280, status: '已申报', source: '月度申报导入', updated_at: '2026-09-22T09:18:00+08:00' },
      { year: 2026, month: 8, amount: 7150, status: '已申报', source: '月度申报导入', updated_at: '2026-08-23T10:03:00+08:00' },
      { year: 2026, month: 7, amount: 6890, status: '已申报', source: '月度申报导入', updated_at: '2026-07-24T11:08:00+08:00' },
      { year: 2026, month: 6, amount: 7010, status: '已申报', source: '月度申报导入', updated_at: '2026-06-24T08:42:00+08:00' },
      { year: 2026, month: 5, amount: 6640, status: '已申报', source: '月度申报导入', updated_at: '2026-05-23T15:11:00+08:00' },
    ],
  },
  {
    id: 103, person_code: 'P000103', group_id: 2, group_kind: '包工头', group_name: '刘师傅班组',
    legal_name: '陈国强', display_name: '陈国强（刘）', gender: '男', id_number: '330106196912080031',
    phone: '13600001058', bank_number: '6228483100001058016', bank_name: '中国农业银行',
    bank_branch: '上海大场支行', bank_branch_code: '103290088888', address: '上海市宝山区样例街28号',
    birth_date: '1969-12-08', issuing_authority: '杭州市公安局西湖分局', id_valid_period: '长期',
    notes: '同名人员较多，以显示名称、完整手机号或身份证号识别。', active: 1,
    salary_history: [
      { year: 2026, month: 9, amount: 9100, status: '已申报', source: '月度申报导入', updated_at: '2026-09-22T09:18:00+08:00' },
      { year: 2026, month: 8, amount: 8840, status: '已申报', source: '月度申报导入', updated_at: '2026-08-23T10:03:00+08:00' },
      { year: 2026, month: 7, amount: 9320, status: '已申报', source: '月度申报导入', updated_at: '2026-07-24T11:08:00+08:00' },
      { year: 2026, month: 6, amount: 8590, status: '已申报', source: '月度申报导入', updated_at: '2026-06-24T08:42:00+08:00' },
    ],
  },
  {
    id: 104, person_code: 'P000104', group_id: 2, group_kind: '包工头', group_name: '刘师傅班组',
    legal_name: '周桂芳', display_name: '周桂芳', gender: '女', id_number: '340103197203170046',
    phone: '13700001072', bank_number: '6212263100001072005', bank_name: '中国工商银行',
    bank_branch: '上海沪太路支行', bank_branch_code: '102290066666', address: '上海市普陀区测试路39号',
    birth_date: '1972-03-17', issuing_authority: '合肥市公安局庐阳分局', id_valid_period: '2018-03-17 至 2038-03-17',
    notes: '', active: 1,
    salary_history: [
      { year: 2026, month: 9, amount: 6450, status: '已申报', source: '月度申报导入', updated_at: '2026-09-22T09:18:00+08:00' },
      { year: 2026, month: 8, amount: 6210, status: '已申报', source: '月度申报导入', updated_at: '2026-08-23T10:03:00+08:00' },
      { year: 2026, month: 7, amount: 5980, status: '已申报', source: '月度申报导入', updated_at: '2026-07-24T11:08:00+08:00' },
    ],
  },
  {
    id: 105, person_code: 'P000105', group_id: 3, group_kind: '分类', group_name: '点工',
    legal_name: '赵海波', display_name: '赵海波', gender: '男', id_number: '310110198105200052',
    phone: '13500001089', bank_number: '6225883100001089012', bank_name: '招商银行',
    bank_branch: '上海五角场支行', bank_branch_code: '308290055555', address: '上海市杨浦区示范路12号',
    birth_date: '1981-05-20', issuing_authority: '上海市公安局杨浦分局', id_valid_period: '2021-05-20 至 2041-05-20',
    notes: '按点工分类统计。', active: 1,
    salary_history: [
      { year: 2026, month: 9, amount: 4350, status: '已申报', source: '前端手动修改', updated_at: '2026-09-23T16:45:00+08:00' },
      { year: 2026, month: 8, amount: 3890, status: '已申报', source: '月度申报导入', updated_at: '2026-08-23T10:03:00+08:00' },
      { year: 2026, month: 7, amount: 4210, status: '已申报', source: '月度申报导入', updated_at: '2026-07-24T11:08:00+08:00' },
    ],
  },
];

const demoLogs = [
  {
    id: 5001, operation_code: 'OP202609230001', module: '工资表', action_type: 'UPDATE_SALARY',
    action_label: '手动修改申报金额', target_name: '赵海波', batch_code: '',
    detail: '赵海波 2026-09：4,100.00 元改为 4,350.00 元',
    before_json: '{"month":"2026-09","amount":4100}', after_json: '{"month":"2026-09","amount":4350}',
    created_at: '2026-09-23T16:45:00+08:00',
  },
  {
    id: 5002, operation_code: 'OP202609220001', module: '申报管理', action_type: 'IMPORT_DECLARATION_BATCH',
    action_label: '导入月度申报', target_name: '2026年9月申报表.xlsx', batch_code: 'B20260922001',
    detail: '共5行，匹配5人，申报总金额35,830.00元', before_json: '',
    after_json: '{"period":"2026-09","row_count":5,"total_amount":35830}', created_at: '2026-09-22T09:18:00+08:00',
  },
  {
    id: 5003, operation_code: 'OP202609180001', module: '人员管理', action_type: 'UPDATE_PERSON',
    action_label: '修改人员资料', target_name: '李秀兰', batch_code: '',
    detail: '更新工资卡开户行信息', before_json: '{"bank_branch":"上海共和路支行"}',
    after_json: '{"bank_branch":"上海共和新路支行"}', created_at: '2026-09-18T14:12:00+08:00',
  },
  {
    id: 5004, operation_code: 'OP202609010001', module: '数据维护', action_type: 'CREATE_BACKUP',
    action_label: '创建数据备份', target_name: '人员工资数据', batch_code: '',
    detail: '已创建9月月初数据备份', before_json: '', after_json: '', created_at: '2026-09-01T08:30:00+08:00',
  },
];

function cloneDemo(value) {
  return JSON.parse(JSON.stringify(value));
}

function demoPause() {
  return new Promise((resolve) => window.setTimeout(resolve, 180));
}

function demoSalary(person, year, month) {
  return (person.salary_history || []).find((item) => Number(item.year) === Number(year) && Number(item.month) === Number(month));
}

function demoReceiptPngBlob() {
  return new Promise((resolve, reject) => {
    const canvas = document.createElement('canvas');
    canvas.width = 960;
    canvas.height = 620;
    const context = canvas.getContext('2d');
    if (!context) {
      reject(new ApiError('当前浏览器无法生成演示凭据'));
      return;
    }
    context.fillStyle = '#f7f9fc';
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.fillStyle = '#155eef';
    context.fillRect(0, 0, canvas.width, 92);
    context.fillStyle = '#ffffff';
    context.font = '700 34px sans-serif';
    context.fillText('付款凭据 · 演示图片', 48, 59);
    context.fillStyle = '#172033';
    context.font = '700 29px sans-serif';
    context.fillText('2026年9月工资付款回单', 48, 158);
    context.font = '24px sans-serif';
    const rows = [
      ['收款人', '张建国（演示）'],
      ['付款日期', '2026-09-22'],
      ['付款金额', '8,650.00 元'],
      ['状态', '付款成功'],
    ];
    rows.forEach(([label, value], index) => {
      const y = 232 + index * 72;
      context.fillStyle = '#5f6b7c';
      context.fillText(label, 52, y);
      context.fillStyle = '#172033';
      context.fillText(value, 250, y);
      context.strokeStyle = '#dce3ed';
      context.beginPath();
      context.moveTo(48, y + 25);
      context.lineTo(912, y + 25);
      context.stroke();
    });
    context.fillStyle = '#16805a';
    context.font = '700 22px sans-serif';
    context.fillText('演示资料 · 不连接正式数据', 48, 558);
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new ApiError('演示图片生成失败'));
    }, 'image/png');
  });
}

function demoReceiptPdfBlob() {
  const encoder = new TextEncoder();
  const content = 'BT\n/F1 22 Tf\n70 760 Td\n(DEMO PAYMENT RECEIPT) Tj\n0 -48 Td\n/F1 14 Tf\n(Period: 2026-09) Tj\n0 -28 Td\n(Payee: Zhang Jianguo - Demo) Tj\n0 -28 Td\n(Amount: CNY 8,650.00) Tj\n0 -28 Td\n(Status: Paid) Tj\nET\n';
  const objects = [
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n',
    `4 0 obj\n<< /Length ${encoder.encode(content).byteLength} >>\nstream\n${content}endstream\nendobj\n`,
    '5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n',
  ];
  let pdf = '%PDF-1.4\n';
  const offsets = [0];
  objects.forEach((object) => {
    offsets.push(encoder.encode(pdf).byteLength);
    pdf += object;
  });
  const xrefOffset = encoder.encode(pdf).byteLength;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.slice(1).forEach((offset) => { pdf += `${String(offset).padStart(10, '0')} 00000 n \n`; });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`;
  return new Blob([pdf], { type: 'application/pdf' });
}

const localApi = {
  async session() {
    if (DEMO_MODE) return { authenticated: false };
    return request('/api/session');
  },

  async login(accessCode) {
    if (DEMO_MODE) {
      await demoPause();
      if (!String(accessCode || '').trim()) throw new ApiError('请输入访问码', 400);
      return { authenticated: true, demo: true };
    }
    return request('/api/login', { method: 'POST', body: { accessCode } });
  },

  async logout() {
    if (DEMO_MODE) return { loggedOut: true };
    return request('/api/logout', { method: 'POST' });
  },

  async groups() {
    if (DEMO_MODE) {
      await demoPause();
      return cloneDemo(demoGroups);
    }
    return request('/api/groups');
  },

  async searchPeople(q, year) {
    if (DEMO_MODE) {
      await demoPause();
      const needle = String(q || '').toLowerCase();
      return demoPeople
        .filter((person) => [person.legal_name, person.display_name, person.id_number, person.phone, person.bank_number, person.group_name]
          .some((value) => String(value || '').toLowerCase().includes(needle)))
        .map((person) => ({
          id: person.id,
          person_code: person.person_code,
          legal_name: person.legal_name,
          display_name: person.display_name,
          is_group_leader: person.is_group_leader,
          group_kind: person.group_kind,
          group_name: person.group_name,
          phone: person.phone,
          id_number: person.id_number,
          bank_number: person.bank_number,
          annual_total: person.salary_history
            .filter((item) => Number(item.year) === Number(year))
            .reduce((sum, item) => sum + safeNumber(item.amount), 0),
        }));
    }
    return request(`/api/people/search${makeQuery({ q, year })}`);
  },

  async person(id) {
    if (DEMO_MODE) {
      await demoPause();
      const person = demoPeople.find((item) => Number(item.id) === Number(id));
      if (!person) throw new ApiError('人员不存在', 404);
      const result = cloneDemo(person);
      result.salary_trend = [...result.salary_history].sort((a, b) => Number(a.year) - Number(b.year) || Number(a.month) - Number(b.month));
      return result;
    }
    return request(`/api/people/${encodeURIComponent(id)}`);
  },

  async declarations(filters) {
    if (DEMO_MODE) {
      await demoPause();
      const needle = String(filters.q || '').toLowerCase();
      const rows = demoPeople
        .filter((person) => !filters.groupId || Number(person.group_id) === Number(filters.groupId))
        .filter((person) => !needle || [person.legal_name, person.display_name, person.id_number, person.phone, person.group_name]
          .some((value) => String(value || '').toLowerCase().includes(needle)))
        .map((person) => ({ person, salary: demoSalary(person, filters.year, filters.month) }))
        .filter((item) => item.salary)
        .map(({ person, salary }) => ({
          person_id: person.id,
          person_code: person.person_code,
          legal_name: person.legal_name,
          display_name: person.display_name,
          is_group_leader: person.is_group_leader,
          id_number: person.id_number,
          phone: person.phone,
          bank_number: person.bank_number,
          group_id: person.group_id,
          group_kind: person.group_kind,
          group_name: person.group_name,
          year: Number(filters.year),
          month: Number(filters.month),
          month_key: formatMonth(filters.year, filters.month),
          amount: salary.amount,
          status: salary.status,
          source: salary.source,
          updated_at: salary.updated_at,
          receipt_count: Math.max(0, Number(salary.receipt_count) || 0),
          batch_code: salary.source === '月度申报导入' ? `B${filters.year}${String(filters.month).padStart(2, '0')}DEMO` : '',
        }));
      return {
        rows,
        summary: {
          month_key: formatMonth(filters.year, filters.month),
          people_count: rows.length,
          total_amount: rows.reduce((sum, row) => sum + safeNumber(row.amount), 0),
        },
      };
    }
    return request(`/api/declarations${makeQuery({
      year: filters.year,
      month: filters.month,
      groupId: filters.groupId,
      q: filters.q,
    })}`);
  },

  async receipts(filters) {
    if (DEMO_MODE) {
      await demoPause();
      if (Number(filters.personId) === 101 && Number(filters.year) === 2026 && Number(filters.month) === 9) {
        return [
          { asset_id: 'demo-receipt-image', person_id: 101, year: 2026, month: 9, display_name: '2026年9月付款截图.png', mime_type: 'image/png', size_bytes: 68432, created_at: '2026-09-22T09:20:00+08:00', source_scope: 'legacy_shared_folder' },
          { asset_id: 'demo-receipt-pdf', person_id: 101, year: 2026, month: 9, display_name: '银行付款回单.pdf', mime_type: 'application/pdf', size_bytes: 1024, created_at: '2026-09-22T09:21:00+08:00' },
        ];
      }
      return [];
    }
    return request(`/api/receipts${makeQuery({
      personId: filters.personId,
      year: filters.year,
      month: filters.month,
    })}`);
  },

  async receiptAsset(assetId) {
    if (!DEMO_MODE) throw new ApiError('本地手机查询暂不提供付款凭据预览，请在电脑端查看。', 501);
    await demoPause();
    let blob;
    let displayName;
    if (assetId === 'demo-receipt-image') {
      blob = await demoReceiptPngBlob();
      displayName = '2026年9月付款截图.png';
    } else if (assetId === 'demo-receipt-pdf') {
      blob = demoReceiptPdfBlob();
      displayName = '银行付款回单.pdf';
    } else {
      throw new ApiError('演示付款凭据不存在', 404);
    }
    return {
      asset_id: assetId,
      display_name: displayName,
      mime_type: blob.type,
      size_bytes: blob.size,
      blob,
    };
  },

  async stats(year, month) {
    if (DEMO_MODE) {
      await demoPause();
      const selected = demoPeople.map((person) => ({ person, salary: demoSalary(person, year, month) }));
      const groups = demoGroups.map((group) => {
        const members = selected.filter(({ person }) => Number(person.group_id) === Number(group.id));
        return {
          id: group.id,
          kind: group.kind,
          name: group.name,
          people_count: members.length,
          paid_people: members.filter(({ salary }) => salary && safeNumber(salary.amount) > 0).length,
          total_amount: members.reduce((sum, { salary }) => sum + safeNumber(salary?.amount), 0),
        };
      }).sort((a, b) => b.total_amount - a.total_amount);
      const monthly = Array.from({ length: 12 }, (_, index) => {
        const selectedMonth = index + 1;
        const salaries = demoPeople.map((person) => demoSalary(person, year, selectedMonth)).filter(Boolean);
        return {
          month: selectedMonth,
          total_amount: salaries.reduce((sum, salary) => sum + safeNumber(salary.amount), 0),
          paid_people: salaries.filter((salary) => safeNumber(salary.amount) > 0).length,
        };
      });
      return {
        summary: {
          total_amount: selected.reduce((sum, { salary }) => sum + safeNumber(salary?.amount), 0),
          paid_people: selected.filter(({ salary }) => salary && safeNumber(salary.amount) > 0).length,
          active_people: demoPeople.length,
          group_count: demoGroups.length,
        },
        groups,
        monthly,
      };
    }
    return request(`/api/stats${makeQuery({ year, month })}`);
  },

  async logs(filters) {
    if (DEMO_MODE) {
      await demoPause();
      const needle = String(filters.q || '').toLowerCase();
      return cloneDemo(demoLogs.filter((log) => (!filters.module || log.module === filters.module)
        && (!filters.actionType || log.action_type === filters.actionType)
        && (!needle || [log.target_name, log.detail, log.operation_code, log.batch_code]
          .some((value) => String(value || '').toLowerCase().includes(needle))))
        .slice(0, Number(filters.limit) || 100));
    }
    return request(`/api/logs${makeQuery({
      module: filters.module,
      actionType: filters.actionType,
      q: filters.q,
      limit: filters.limit,
    })}`);
  },
};

function unavailableOneDriveApi(error) {
  const reject = async () => { throw error; };
  return Object.freeze({
    session: async () => ({ authenticated: false, cloud: true, provider: 'onedrive' }),
    login: reject,
    logout: async () => ({ loggedOut: true }),
    groups: reject,
    searchPeople: reject,
    person: reject,
    declarations: reject,
    receipts: reject,
    receiptAsset: reject,
    stats: reject,
    logs: reject,
    metadata: reject,
    hasPairing: () => false,
    microsoftLogin: reject,
    microsoftStatus: () => ({ available: false, signedIn: false, clientId: '' }),
  });
}

function createOneDriveApi() {
  if (oneDriveStartupError) return unavailableOneDriveApi(oneDriveStartupError);
  const pairing = oneDriveDetection.pairing;
  const clientId = oneDriveDetection.clientId || oneDriveHint?.clientId || '';
  try {
    const authProvider = window.PayrollMicrosoftAuth.createMicrosoftAuthProvider({
      clientId,
      location: window.location,
      msal: window.msal,
      scopes: window.PayrollOneDriveSnapshot.DEFAULT_SCOPES,
    });
    return window.PayrollOneDriveSnapshot.createOneDriveClient({
      pairing,
      clientId,
      authProvider,
      detectPairing: false,
      location: window.location,
      history: window.history,
      storage: window.localStorage,
      fetch: window.fetch.bind(window),
      crypto: window.crypto,
    });
  } catch (error) {
    oneDriveStartupError = error;
    return unavailableOneDriveApi(error);
  } finally {
    // The adapter owns the pending pairing until it has been PIN-wrapped. Do
    // not retain a second data-key reference in the page controller.
    oneDriveDetection = { active: false, pairing: null, clientId: '' };
  }
}

const api = ONEDRIVE_MODE
  ? createOneDriveApi()
  : LEGACY_CLOUD_MODE
    ? window.PayrollCloudSnapshot.createCloudClient({
      location: window.location,
      history: window.history,
      storage: window.localStorage,
      fetch: window.fetch.bind(window),
      crypto: window.crypto,
      baseUrl: window.location.origin,
    })
    : localApi;

function setLoading(visible, message = '正在读取，请稍候') {
  $('#loading-text').textContent = message;
  $('#loading-overlay').classList.toggle('hidden', !visible);
  $('#loading-overlay').setAttribute('aria-hidden', visible ? 'false' : 'true');
}

function toast(message, type = 'normal') {
  const element = $('#toast');
  element.textContent = message;
  element.className = `toast${type === 'error' ? ' error' : ''}`;
  window.clearTimeout(element._hideTimer);
  element._hideTimer = window.setTimeout(() => element.classList.add('hidden'), 3600);
}

function setLoginError(message = '') {
  const element = $('#login-error');
  element.textContent = message;
  element.classList.toggle('hidden', !message);
}

function updateMicrosoftLoginUi() {
  const step = $('#microsoft-login-step');
  const stepHeading = $('#access-code-step-heading');
  const plainLabel = $('#access-code-label');
  if (!ONEDRIVE_MODE) {
    step.classList.add('hidden');
    stepHeading.classList.add('hidden');
    plainLabel.classList.remove('hidden');
    return;
  }

  step.classList.remove('hidden');
  stepHeading.classList.remove('hidden');
  plainLabel.classList.add('hidden');
  const status = api.microsoftStatus();
  const statusElement = $('#microsoft-login-status');
  const loginButton = $('#microsoft-login-button');
  const pinInput = $('#access-code');
  const submitButton = $('#login-button');
  statusElement.className = 'field-hint';

  if (!status.available) {
    statusElement.textContent = oneDriveStartupError?.message
      || '缺少官方 Microsoft 登录组件（mobile/vendor/msal-browser.min.js），请联系管理员完成手机网页配置。';
    statusElement.classList.add('error');
    loginButton.textContent = 'Microsoft 登录暂不可用';
    loginButton.disabled = true;
    pinInput.disabled = true;
    submitButton.disabled = true;
    return;
  }
  if (status.signedIn) {
    statusElement.textContent = 'Microsoft 账号已登录，请继续输入8位访问码。';
    statusElement.classList.add('ready');
    loginButton.textContent = '重新选择 Microsoft 账号';
    pinInput.disabled = false;
    submitButton.disabled = false;
    return;
  }
  statusElement.textContent = '尚未登录。手机端不会保存 Microsoft 登录令牌。';
  loginButton.textContent = '登录 Microsoft 账号';
  loginButton.disabled = false;
  pinInput.disabled = true;
  submitButton.disabled = true;
}

function fillDateSelects() {
  const startYear = Math.min(2020, DEFAULT_YEAR - 6);
  const endYear = DEFAULT_YEAR + 1;
  const years = [];
  for (let year = endYear; year >= startYear; year -= 1) years.push(year);
  const yearOptions = years.map((year) => `<option value="${year}">${year}年</option>`).join('');
  ['monthly-year', 'stats-year'].forEach((id) => {
    $(`#${id}`).innerHTML = yearOptions;
    $(`#${id}`).value = String(DEFAULT_YEAR);
  });
  const monthOptions = Array.from({ length: 12 }, (_, index) => `<option value="${index + 1}">${index + 1}月</option>`).join('');
  ['monthly-month', 'stats-month'].forEach((id) => {
    $(`#${id}`).innerHTML = monthOptions;
    $(`#${id}`).value = String(DEFAULT_MONTH);
  });
}

function populateGroupSelect() {
  const options = state.groups.map((group) => (
    `<option value="${Number(group.id)}">${escapeHtml(group.kind)}｜${escapeHtml(group.name)}（${formatInteger(group.person_count)}人）</option>`
  )).join('');
  $('#monthly-group').innerHTML = `<option value="">全部班组</option>${options}`;
}

function showAuthenticatedApp() {
  state.authenticated = true;
  cloudHiddenAt = 0;
  if (cloudLockTimer) window.clearTimeout(cloudLockTimer);
  cloudLockTimer = null;
  $('#login-screen').classList.add('hidden');
  $('#app-shell').classList.remove('hidden');
  $('#demo-app-notice').classList.toggle('hidden', !DEMO_MODE);
  navigate('home');
  $('#main-content').focus({ preventScroll: true });
}

function renderCloudSnapshotMetadata(metadata) {
  const strip = $('#cloud-sync-strip');
  if (!CLOUD_MODE || !metadata) {
    strip.classList.add('hidden');
    return;
  }
  const generatedAt = new Date(metadata.generated_at);
  const validDate = !Number.isNaN(generatedAt.getTime());
  const ageDays = validDate ? Math.floor((Date.now() - generatedAt.getTime()) / 86400000) : 0;
  strip.textContent = validDate
    ? `云端资料更新至 ${formatDateTime(metadata.generated_at)} · 第 ${formatInteger(metadata.sequence)} 版${ageDays >= 30 ? ` · 已超过${ageDays}天未同步` : ''}`
    : `云端资料第 ${formatInteger(metadata.sequence)} 版`;
  strip.classList.toggle('stale', ageDays >= 30);
  strip.classList.remove('hidden');
}

function showLogin() {
  state.authenticated = false;
  $('#app-shell').classList.add('hidden');
  $('#login-screen').classList.remove('hidden');
  $('#access-code').value = DEMO_MODE ? '12345678' : '';
  updateMicrosoftLoginUi();
  if (ONEDRIVE_MODE && !api.microsoftStatus().signedIn) {
    $('#microsoft-login-button').focus({ preventScroll: true });
  } else {
    $('#access-code').focus({ preventScroll: true });
  }
}

function clearSensitiveViews() {
  state.groups = [];
  state.loaded = { stats: false, logs: false };
  state.retryActions.clear();
  $('#monthly-group').innerHTML = '<option value="">全部班组</option>';
  $('#search-results').replaceChildren();
  $('#search-results-section').classList.add('hidden');
  $('#search-welcome').classList.remove('hidden');
  $('#monthly-results').replaceChildren();
  $('#monthly-summary').replaceChildren();
  $('#monthly-summary').classList.add('hidden');
  $('#monthly-placeholder').classList.remove('hidden');
  $('#stats-cards').replaceChildren();
  $('#monthly-chart').replaceChildren();
  $('#group-stats').replaceChildren();
  $('#stats-content').classList.add('hidden');
  $('#stats-placeholder').classList.remove('hidden');
  $('#logs-results').replaceChildren();
  $('#logs-count').textContent = '0 条';
  $('#logs-placeholder').classList.remove('hidden');
  $('#person-detail').replaceChildren();
  $('#people-search').value = '';
  $('#monthly-search').value = '';
  $('#logs-search').value = '';
  closeReceiptDialog();
}

async function handleLogin(event) {
  event.preventDefault();
  setLoginError();
  const accessCode = $('#access-code').value.trim();
  if (!accessCode) {
    setLoginError('请输入访问码。');
    $('#access-code').focus();
    return;
  }

  $('#login-button').disabled = true;
  setLoading(true, '正在验证访问码');
  try {
    await api.login(accessCode);
    state.groups = await api.groups();
    populateGroupSelect();
    showAuthenticatedApp();
    if (CLOUD_MODE) renderCloudSnapshotMetadata(await api.metadata());
  } catch (error) {
    const pairingProblem = ['PAIRING_REQUIRED', 'PAIRING_REJECTED', 'INVALID_STORED_PAIRING'].includes(error.code);
    const microsoftProblem = [
      'MICROSOFT_SIGNIN_REQUIRED',
      'MICROSOFT_REAUTH_REQUIRED',
      'ONEDRIVE_AUTH_FAILED',
      'ONEDRIVE_AUTH_REJECTED',
      'MSAL_NOT_AVAILABLE',
      'MSAL_INITIALIZE_FAILED',
      'TOKEN_UNAVAILABLE',
    ].includes(error.code);
    setLoginError(pairingProblem || microsoftProblem
      ? error.message
      : (error.status === 401 ? '访问码不正确，请重新输入。' : normalizedError(error, '暂时无法登录，请稍后重试。')));
    if (microsoftProblem) updateMicrosoftLoginUi();
    else $('#access-code').select();
  } finally {
    setLoading(false);
    $('#login-button').disabled = false;
    updateMicrosoftLoginUi();
  }
}

async function handleMicrosoftLogin() {
  if (!ONEDRIVE_MODE) return;
  setLoginError();
  $('#microsoft-login-button').disabled = true;
  setLoading(true, '正在打开 Microsoft 登录');
  try {
    await api.microsoftLogin();
    updateMicrosoftLoginUi();
    $('#access-code').focus({ preventScroll: true });
  } catch (error) {
    setLoginError(normalizedError(error, 'Microsoft 登录未完成，请重试。'));
  } finally {
    setLoading(false);
    updateMicrosoftLoginUi();
  }
}

async function handleLogout() {
  if (cloudLockTimer) window.clearTimeout(cloudLockTimer);
  cloudLockTimer = null;
  setLoading(true, '正在安全退出');
  try {
    await api.logout();
  } catch (_error) {
    // The local screen is still cleared if the server is temporarily unavailable.
  } finally {
    clearSensitiveViews();
    setLoading(false);
    showLogin();
  }
}

async function lockCloudSessionAfterBackground() {
  if (!CLOUD_MODE || !state.authenticated) return;
  if (cloudLockTimer) window.clearTimeout(cloudLockTimer);
  cloudLockTimer = null;
  try { await api.logout(); } catch (_error) { /* Clearing the local view is sufficient. */ }
  clearSensitiveViews();
  showLogin();
  setLoginError('为保护完整证件信息，应用退到后台超过15分钟后已自动锁定。');
}

function navigate(view) {
  if (view !== 'person') state.currentView = view;
  $$('.view').forEach((element) => element.classList.toggle('hidden', element.dataset.view !== view));
  const highlighted = view === 'person' ? state.returnView : view;
  $$('.nav-button').forEach((button) => {
    const active = button.dataset.nav === highlighted;
    button.classList.toggle('active', active);
    if (active) button.setAttribute('aria-current', 'page');
    else button.removeAttribute('aria-current');
  });
  window.scrollTo({ top: 0, behavior: 'auto' });

  if (view === 'stats' && !state.loaded.stats) loadStats();
  if (view === 'logs' && !state.loaded.logs) loadLogs();
}

function renderError(container, message, retryKey, retryAction) {
  state.retryActions.set(retryKey, retryAction);
  container.innerHTML = `
    <div class="error-panel" role="alert">
      <h3>暂时无法读取</h3>
      <p>${escapeHtml(message)}</p>
      <button type="button" data-retry="${escapeHtml(retryKey)}">重新读取</button>
    </div>`;
}

function peopleResultCard(person) {
  const displayName = mobilePersonName(person);
  const legalName = cleanPersonName(person.legal_name);
  const legalSuffix = displayName && legalName && displayName !== legalName
    ? `（姓名：${escapeHtml(legalName)}）` : '';
  const leader = isGroupLeader(person);
  return `
    <button class="person-result-card ${leader ? 'leader-person-card' : ''}" type="button" data-person-id="${Number(person.id)}"
      aria-label="查看${escapeHtml(displayName)}${leader ? '，包工头本人' : ''}的人员详情">
      <span class="result-card-top">
        <span class="result-name ${leader ? 'leader-person-name' : ''}">${escapeHtml(displayName)}${legalSuffix}${leaderBadge(person)}${materialPaymentBadge(person)}</span>
        <span class="count-pill">查看详情</span>
      </span>
      <span class="result-group">${escapeHtml(displayValue(person.group_kind, '未分组'))}｜${escapeHtml(displayValue(person.group_name, '未分组'))}</span>
      <span class="result-details">
        <span class="result-detail-row"><span class="result-label">身份证</span><span class="result-value">${escapeHtml(displayValue(person.id_number))}</span></span>
        <span class="result-detail-row"><span class="result-label">手机号</span><span class="result-value">${escapeHtml(displayValue(person.phone))}</span></span>
        <span class="result-detail-row"><span class="result-label">银行卡</span><span class="result-value">${escapeHtml(displayValue(person.bank_number))}</span></span>
      </span>
      <span class="result-amount">本年申报合计 <strong>${formatMoney(person.annual_total)} 元</strong></span>
    </button>`;
}

async function searchPeople() {
  const query = $('#people-search').value.trim();
  if (!query) {
    toast('请先输入姓名、身份证、手机号或包工头。', 'error');
    $('#people-search').focus();
    return;
  }

  $('#search-welcome').classList.add('hidden');
  $('#search-results-section').classList.remove('hidden');
  setLoading(true, '正在查找人员');
  try {
    const rows = await api.searchPeople(query, DEFAULT_YEAR);
    $('#search-result-count').textContent = `${formatInteger(rows.length)} 人`;
    if (!rows.length) {
      $('#search-results').innerHTML = `
        <div class="state-panel compact-state">
          <div class="state-icon search-state-icon" aria-hidden="true"></div>
          <h3>没有找到人员</h3>
          <p>请检查姓名或号码，也可改用包工头名称。</p>
        </div>`;
    } else {
      $('#search-results').innerHTML = rows.map(peopleResultCard).join('');
    }
    $('#search-results-title').focus?.();
  } catch (error) {
    if (error.status === 401) return handleExpiredSession();
    $('#search-result-count').textContent = '读取失败';
    renderError($('#search-results'), normalizedError(error), 'people-search', searchPeople);
  } finally {
    setLoading(false);
  }
}

function personDetailField(label, value, className = '') {
  return `<div><dt>${escapeHtml(label)}</dt><dd class="${escapeHtml(className)}">${escapeHtml(displayValue(value))}</dd></div>`;
}

function renderPersonTrend(person) {
  const source = Array.isArray(person.salary_trend) && person.salary_trend.length
    ? person.salary_trend
    : (person.salary_history || []).slice().sort((a, b) => Number(a.year) - Number(b.year) || Number(a.month) - Number(b.month));
  const trend = source.filter((item) => item.status !== '未申报' || safeNumber(item.amount) !== 0).slice(-12);
  if (!trend.length) return '<div class="empty-inline">暂无申报趋势</div>';
  const maximum = Math.max(1, ...trend.map((item) => safeNumber(item.amount)));
  return `
    <div class="person-trend" role="img" aria-label="最近${trend.length}个月申报金额趋势">
      ${trend.map((item) => {
        const amount = safeNumber(item.amount);
        const height = Math.max(3, Math.min(100, (amount / maximum) * 100));
        const month = `${String(item.year).slice(-2)}-${String(item.month).padStart(2, '0')}`;
        return `
          <div class="person-trend-item" title="${escapeHtml(formatMonth(item.year, item.month))}：${formatMoney(amount)}元">
            <div class="person-trend-bar-wrap"><span class="person-trend-bar" style="height:${height.toFixed(2)}%"></span></div>
            <span class="person-trend-label">${escapeHtml(month)}</span>
          </div>`;
      }).join('')}
    </div>`;
}

function receiptActionButton(personId, year, month, count, className = '', personName = '') {
  const total = Math.max(0, Number(count) || 0);
  if (!total || (!ONEDRIVE_MODE && !DEMO_MODE)) return '';
  return `<button class="receipt-action ${escapeHtml(className)}" type="button"
    data-receipts data-receipt-person-id="${Number(personId)}" data-receipt-year="${Number(year)}"
    data-receipt-month="${Number(month)}" data-receipt-person-name="${escapeHtml(personName)}">查看凭据（${formatInteger(total)}）</button>`;
}

function clearReceiptDialogContents() {
  receiptRequestSequence += 1;
  clearReceiptViewer();
  $('#receipt-results').replaceChildren();
  $('#receipt-dialog-subtitle').textContent = '';
}

function clearReceiptViewer() {
  if (activeReceiptObjectUrl) {
    window.URL.revokeObjectURL(activeReceiptObjectUrl);
    activeReceiptObjectUrl = '';
  }
  const image = $('#receipt-image-viewer');
  const pdf = $('#receipt-pdf-viewer');
  image.removeAttribute('src');
  pdf.removeAttribute('src');
  image.classList.add('hidden');
  pdf.classList.add('hidden');
  $('#receipt-viewer-error').classList.add('hidden');
  $('#receipt-viewer-error').textContent = '';
  $('#receipt-viewer-title').textContent = '';
  $('#receipt-viewer').classList.add('hidden');
  $('#receipt-results').classList.remove('hidden');
}

function closeReceiptDialog() {
  const dialog = $('#receipt-dialog');
  if (typeof dialog.close === 'function' && dialog.open) dialog.close();
  else dialog.removeAttribute('open');
  clearReceiptDialogContents();
}

function receiptDisplayName(receipt) {
  return displayValue(receipt?.display_name || receipt?.file_name || receipt?.original_name || receipt?.name, '付款凭据');
}

function formatFileSize(value) {
  const bytes = Math.max(0, Number(value) || 0);
  if (!bytes) return '';
  if (bytes < 1024) return `${formatInteger(bytes)} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function renderReceiptList(records) {
  const container = $('#receipt-results');
  if (!records.length) {
    container.innerHTML = '<div class="empty-inline">该月份暂无可查看的付款凭据</div>';
    return;
  }
  container.innerHTML = records.map((receipt, index) => {
    const type = displayValue(receipt.mime_type || receipt.content_type || receipt.file_type, '附件');
    const time = receipt.uploaded_at || receipt.created_at || receipt.updated_at;
    const size = formatFileSize(receipt.size_bytes || receipt.file_size);
    const assetId = String(receipt.asset_id || receipt.id || '').trim();
    const sharedLegacy = receipt.source_scope === 'legacy_shared_folder';
    const countLabel = `${index + 1}`.padStart(2, '0');
    return `<article class="receipt-list-item">
      <span class="receipt-file-index" aria-hidden="true">${countLabel}</span>
      <div>
        <strong>${escapeHtml(receiptDisplayName(receipt))}</strong>
        <p>${escapeHtml(type)}${size ? ` · ${escapeHtml(size)}` : ''}${time ? ` · ${escapeHtml(formatDateTime(time))}` : ''}</p>
        ${sharedLegacy ? '<p class="warning-text">原总表关联的多人/班组共享凭据，请打开后核对姓名与金额。</p>' : ''}
      </div>
      ${assetId && typeof api.receiptAsset === 'function'
        ? `<button class="receipt-open-button" type="button" data-receipt-asset-id="${escapeHtml(assetId)}">打开</button>`
        : ''}
    </article>`;
  }).join('');
}

async function openReceiptAsset(assetId) {
  if (!assetId || typeof api.receiptAsset !== 'function') return;
  clearReceiptViewer();
  const requestSequence = ++receiptRequestSequence;
  setLoading(true, '正在安全解密付款凭据');
  try {
    const asset = await api.receiptAsset(assetId);
    if (requestSequence !== receiptRequestSequence || !$('#receipt-dialog').open) return;
    if (!(asset?.blob instanceof Blob)) throw new ApiError('付款凭据内容格式无效');
    const mimeType = String(asset.mime_type || asset.blob.type || '').toLowerCase();
    const declaredSize = Number(asset.size_bytes);
    if (!Number.isSafeInteger(declaredSize) || declaredSize < 1 || declaredSize > MAX_RECEIPT_PREVIEW_BYTES
      || asset.blob.size !== declaredSize) {
      throw new ApiError('付款凭据大小校验失败');
    }
    if (asset.blob.type && String(asset.blob.type).toLowerCase() !== mimeType) {
      throw new ApiError('付款凭据类型校验失败');
    }
    if (!(mimeType === 'application/pdf' || mimeType === 'image/png' || mimeType === 'image/jpeg')) {
      throw new ApiError('该付款凭据格式暂不支持手机查看');
    }
    activeReceiptObjectUrl = window.URL.createObjectURL(asset.blob);
    $('#receipt-results').classList.add('hidden');
    $('#receipt-viewer').classList.remove('hidden');
    $('#receipt-viewer-title').textContent = receiptDisplayName(asset);
    if (mimeType === 'application/pdf') {
      $('#receipt-pdf-viewer').src = activeReceiptObjectUrl;
      $('#receipt-pdf-viewer').classList.remove('hidden');
    } else {
      $('#receipt-image-viewer').src = activeReceiptObjectUrl;
      $('#receipt-image-viewer').alt = receiptDisplayName(asset);
      $('#receipt-image-viewer').classList.remove('hidden');
    }
  } catch (error) {
    if (requestSequence !== receiptRequestSequence) return;
    clearReceiptViewer();
    if (error.status === 401) {
      closeReceiptDialog();
      return handleExpiredSession();
    }
    toast(normalizedError(error, '付款凭据打开失败'), 'error');
  } finally {
    if (requestSequence === receiptRequestSequence || !$('#receipt-dialog').open) setLoading(false);
  }
}

async function openReceiptList(button) {
  const personId = Number(button.dataset.receiptPersonId);
  const year = Number(button.dataset.receiptYear);
  const month = Number(button.dataset.receiptMonth);
  const personName = String(button.dataset.receiptPersonName || '').trim();
  if (!personId || !year || !month) return;
  const dialog = $('#receipt-dialog');
  const requestSequence = ++receiptRequestSequence;
  clearReceiptViewer();
  $('#receipt-dialog-subtitle').textContent = `${personName ? `${personName} · ` : ''}${formatMonth(year, month)} · 只读查看`;
  $('#receipt-results').innerHTML = '<div class="empty-inline">正在读取付款凭据…</div>';
  if (typeof dialog.showModal === 'function') dialog.showModal();
  else dialog.setAttribute('open', '');

  if (typeof api.receipts !== 'function') {
    $('#receipt-results').innerHTML = '<div class="empty-inline">当前版本尚未启用手机端凭据读取</div>';
    return;
  }
  try {
    const result = await api.receipts({ personId, year, month });
    if (requestSequence !== receiptRequestSequence || !dialog.open) return;
    const records = Array.isArray(result) ? result : Array.isArray(result?.rows) ? result.rows : [];
    renderReceiptList(records);
  } catch (error) {
    if (requestSequence !== receiptRequestSequence) return;
    if (error.status === 401) {
      closeReceiptDialog();
      return handleExpiredSession();
    }
    $('#receipt-results').innerHTML = `<div class="error-panel compact-receipt-error" role="alert"><p>${escapeHtml(normalizedError(error, '付款凭据读取失败'))}</p></div>`;
  }
}

function renderPersonHistory(person) {
  const history = (person.salary_history || [])
    .filter((item) => item.status !== '未申报' || safeNumber(item.amount) !== 0)
    .slice()
    .sort((a, b) => Number(b.year) - Number(a.year) || Number(b.month) - Number(a.month));
  if (!history.length) return '<div class="empty-inline">暂无历月申报记录</div>';
  const total = history.reduce((sum, item) => sum + safeNumber(item.amount), 0);
  return `
    <p class="history-summary">共 ${formatInteger(history.length)} 个月有记录，合计 ${formatMoney(total)} 元</p>
    <div class="history-list">
      ${history.map((item) => `
        <div class="history-row" title="来源：${escapeHtml(displayValue(item.source))}｜更新时间：${escapeHtml(formatDateTime(item.updated_at))}">
          <time datetime="${escapeHtml(formatMonth(item.year, item.month))}">${escapeHtml(formatMonth(item.year, item.month))}</time>
          <strong>${formatMoney(item.amount)} 元</strong>
          <span class="status-pill">${escapeHtml(displayValue(item.status, '有记录'))}${Number(item.receipt_count) ? ` · 凭据${formatInteger(item.receipt_count)}份` : ''}</span>
          ${receiptActionButton(person.id, item.year, item.month, item.receipt_count, '', mobilePersonName(person))}
        </div>`).join('')}
    </div>`;
}

function renderPersonDetail(person) {
  const name = mobilePersonName(person);
  const group = [person.group_kind, person.group_name].filter(Boolean).join('｜') || '未分组';
  const bankAndBranch = [person.bank_name, person.bank_branch].filter(Boolean).join('｜');
  $('#person-detail').innerHTML = `
    <article>
      <div class="person-hero ${isGroupLeader(person) ? 'leader-person-hero' : ''}">
        <div class="detail-identity">
          <div class="person-initial" aria-hidden="true">${escapeHtml(initialFor(person))}</div>
          <div>
            <h2 id="person-detail-title">${escapeHtml(name)}${leaderBadge(person)}${materialPaymentBadge(person)}</h2>
            <p>${escapeHtml(group)}｜${escapeHtml(displayValue(person.person_code, '无内部编号'))}</p>
          </div>
        </div>
        <p class="readonly-detail-note">本页为只读资料，不能编辑；有付款凭据的月份可安全查看。</p>
      </div>

      <section class="detail-section" aria-labelledby="identity-section-title">
        <h3 id="identity-section-title">身份与联系方式</h3>
        <dl class="detail-list">
          ${personDetailField('法定姓名', cleanPersonName(person.legal_name))}
          ${personDetailField('显示名称', cleanPersonName(person.display_name))}
          ${personDetailField('性别', person.gender)}
          ${personDetailField('身份证号', person.id_number, 'sensitive-value')}
          ${personDetailField('手机号', person.phone, 'sensitive-value')}
          ${personDetailField('出生日期', person.birth_date)}
          ${personDetailField('归属', group)}
          ${personDetailField('状态', Number(person.active) === 0 ? '已停用' : '在用')}
        </dl>
      </section>

      <section class="detail-section" aria-labelledby="bank-section-title">
        <h3 id="bank-section-title">银行卡与银行</h3>
        <dl class="detail-list">
          ${personDetailField('银行卡号', displayValue(person.bank_number), 'sensitive-value')}
          ${personDetailField('银行 / 开户行', bankAndBranch)}
          ${personDetailField('银行行号', person.bank_branch_code)}
        </dl>
      </section>

      <section class="detail-section" aria-labelledby="other-section-title">
        <h3 id="other-section-title">其他资料</h3>
        <dl class="detail-list">
          ${personDetailField('地址', person.address)}
          ${personDetailField('签发机关', person.issuing_authority)}
          ${personDetailField('证件有效期', person.id_valid_period)}
          ${personDetailField('备注', person.notes)}
        </dl>
      </section>

      <section class="detail-section" aria-labelledby="trend-section-title">
        <h3 id="trend-section-title">历月金额趋势</h3>
        ${renderPersonTrend(person)}
      </section>

      <section class="detail-section" aria-labelledby="history-section-title">
        <h3 id="history-section-title">历月金额列表</h3>
        ${renderPersonHistory(person)}
      </section>
    </article>`;
}

async function openPerson(personId) {
  if (!personId) return;
  const previousView = $('#person-detail-view').classList.contains('hidden') ? state.currentView : state.returnView;
  setLoading(true, '正在读取人员详情');
  try {
    const person = await api.person(personId);
    state.returnView = previousView;
    renderPersonDetail(person);
    navigate('person');
    $('#person-detail-title').focus?.();
  } catch (error) {
    if (error.status === 401) return handleExpiredSession();
    toast(normalizedError(error, '人员详情读取失败'), 'error');
  } finally {
    setLoading(false);
  }
}

function monthlyResultCard(row) {
  const name = mobilePersonName(row);
  const leader = isGroupLeader(row);
  const source = row.batch_code || row.source || '来源未记录';
  return `
    <div class="monthly-result-item">
      <button class="monthly-result-card ${leader ? 'leader-person-card' : ''}" type="button" data-person-id="${Number(row.person_id)}"
      aria-label="查看${escapeHtml(name)}${leader ? '，包工头本人' : ''}的人员详情">
        <span class="monthly-card-top">
          <span class="monthly-card-identity">
            <span class="result-name ${leader ? 'leader-person-name' : ''}">${escapeHtml(name)}${leaderBadge(row)}${materialPaymentBadge(row)}</span>
            <span class="monthly-meta"><span>${escapeHtml(displayValue(row.group_name, '未分组'))}</span><span>${escapeHtml(displayValue(row.status, '有记录'))}</span></span>
          </span>
          <span class="amount">${formatMoney(row.amount)} 元</span>
        </span>
        <span class="monthly-meta"><span>身份证：${escapeHtml(displayValue(row.id_number))}</span><span>手机：${escapeHtml(displayValue(row.phone))}</span></span>
        <span class="monthly-source">来源：${escapeHtml(source)}｜付款凭据：${formatInteger(row.receipt_count || 0)} 份｜更新：${escapeHtml(formatDateTime(row.updated_at))}</span>
      </button>
      ${receiptActionButton(row.person_id, row.year, row.month, row.receipt_count, 'monthly-receipt-action', name)}
    </div>`;
}

async function loadMonthly() {
  const filters = {
    year: Number($('#monthly-year').value),
    month: Number($('#monthly-month').value),
    groupId: $('#monthly-group').value,
    q: $('#monthly-search').value.trim(),
  };
  $('#monthly-placeholder').classList.add('hidden');
  setLoading(true, '正在读取月度明细');
  try {
    const result = await api.declarations(filters);
    const rows = Array.isArray(result?.rows) ? result.rows : [];
    const summary = result?.summary || {};
    $('#monthly-summary').innerHTML = `
      <div><span>申报人数</span><strong>${formatInteger(summary.people_count ?? rows.length)} 人</strong></div>
      <div><span>申报合计</span><strong>${formatMoney(summary.total_amount)} 元</strong></div>`;
    $('#monthly-summary').classList.remove('hidden');
    if (rows.length) {
      $('#monthly-results').innerHTML = rows.map(monthlyResultCard).join('');
    } else {
      $('#monthly-results').innerHTML = `
        <div class="state-panel compact-state">
          <div class="state-icon calendar-state-icon" aria-hidden="true"></div>
          <h3>该条件下暂无记录</h3>
          <p>可切换月份、班组，或清空搜索词后重试。</p>
        </div>`;
    }
  } catch (error) {
    if (error.status === 401) return handleExpiredSession();
    $('#monthly-summary').classList.add('hidden');
    renderError($('#monthly-results'), normalizedError(error), 'monthly', loadMonthly);
  } finally {
    setLoading(false);
  }
}

function renderStats(result) {
  const summary = result?.summary || {};
  const monthly = Array.isArray(result?.monthly) ? result.monthly : [];
  const groups = Array.isArray(result?.groups) ? result.groups : [];
  const year = Number($('#stats-year').value);
  const month = Number($('#stats-month').value);
  $('#stats-period').textContent = `${year}年${month}月`;
  $('#stats-cards').innerHTML = `
    <div class="metric-card primary"><span>本月申报金额</span><strong>${formatMoney(summary.total_amount)} 元</strong></div>
    <div class="metric-card"><span>有申报人数</span><strong>${formatInteger(summary.paid_people)} 人</strong></div>
    <div class="metric-card"><span>在用人员</span><strong>${formatInteger(summary.active_people)} 人</strong></div>
    <div class="metric-card"><span>班组 / 包工头</span><strong>${formatInteger(summary.group_count)} 个</strong></div>`;

  const maximum = Math.max(1, ...monthly.map((item) => safeNumber(item.total_amount)));
  $('#monthly-chart').innerHTML = monthly.length ? monthly.map((item) => {
    const amount = safeNumber(item.total_amount);
    const height = Math.max(2, Math.min(100, (amount / maximum) * 100));
    return `
      <div class="chart-bar-item" title="${Number(item.month)}月：${formatMoney(amount)}元，${formatInteger(item.paid_people)}人">
        <div class="chart-bar-track"><span class="chart-bar" style="height:${height.toFixed(2)}%"></span></div>
        <span class="chart-label">${Number(item.month)}月</span>
      </div>`;
  }).join('') : '<div class="empty-inline">暂无月度趋势</div>';

  $('#group-stats').innerHTML = groups.length ? groups.map((group) => `
    <div class="group-summary-card">
      <div class="summary-row">
        <h4><span class="type-pill">${escapeHtml(displayValue(group.kind, '班组'))}</span>${escapeHtml(displayValue(group.name, '未命名'))}</h4>
        <span class="amount">${formatMoney(group.total_amount)} 元</span>
      </div>
      <div class="group-summary-meta">
        <span>班组人数 ${formatInteger(group.people_count)} 人</span>
        <span>本月有申报 ${formatInteger(group.paid_people)} 人</span>
      </div>
    </div>`).join('') : '<div class="empty-inline">暂无班组汇总</div>';
  $('#stats-placeholder').classList.add('hidden');
  $('#stats-content').classList.remove('hidden');
}

async function loadStats() {
  setLoading(true, '正在读取月度统计');
  try {
    const result = await api.stats(Number($('#stats-year').value), Number($('#stats-month').value));
    renderStats(result);
    state.loaded.stats = true;
  } catch (error) {
    if (error.status === 401) return handleExpiredSession();
    $('#stats-content').classList.add('hidden');
    $('#stats-placeholder').classList.remove('hidden');
    renderError($('#stats-placeholder'), normalizedError(error), 'stats', loadStats);
  } finally {
    setLoading(false);
  }
}

function formatSnapshot(value) {
  const text = String(value || '').trim();
  if (!text) return '无';
  try {
    const parsed = JSON.parse(text);
    if (!parsed || typeof parsed !== 'object') return String(parsed);
    return Object.entries(parsed)
      .map(([key, item]) => `${key}：${item === null || item === '' ? '空' : typeof item === 'object' ? JSON.stringify(item) : String(item)}`)
      .join('；');
  } catch (_error) {
    return text;
  }
}

function logCard(log) {
  const extraRows = [
    ['操作编号', log.operation_code],
    ['对象', log.target_name || log.target_id],
    ['批次', log.batch_code],
    ['修改前', formatSnapshot(log.before_json)],
    ['修改后', formatSnapshot(log.after_json)],
  ].filter(([, value]) => value && value !== '无');
  return `
    <article class="log-card">
      <div class="log-card-header">
        <h3>${escapeHtml(displayValue(log.action_label, '操作记录'))}</h3>
        <time class="log-time" datetime="${escapeHtml(log.created_at || '')}">${escapeHtml(formatDateTime(log.created_at))}</time>
      </div>
      <span class="log-module">${escapeHtml(displayValue(log.module, '其他'))}</span>
      <p class="log-detail">${escapeHtml(displayValue(log.detail, '未记录说明'))}</p>
      ${extraRows.length ? `
        <details class="log-extra">
          <summary>查看记录详情</summary>
          <dl>${extraRows.map(([label, value]) => `<dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd>`).join('')}</dl>
        </details>` : ''}
    </article>`;
}

async function loadLogs() {
  setLoading(true, '正在读取操作日志');
  try {
    const logs = await api.logs({
      module: $('#logs-module').value,
      actionType: '',
      q: $('#logs-search').value.trim(),
      limit: 100,
    });
    const rows = Array.isArray(logs) ? logs : [];
    $('#logs-count').textContent = `${formatInteger(rows.length)} 条`;
    $('#logs-results').innerHTML = rows.length ? rows.map(logCard).join('') : `
      <div class="state-panel compact-state">
        <div class="state-icon log-state-icon" aria-hidden="true"></div>
        <h3>暂无操作日志</h3>
        <p>更换模块或清空关键词后再试。</p>
      </div>`;
    $('#logs-placeholder').classList.add('hidden');
    state.loaded.logs = true;
  } catch (error) {
    if (error.status === 401) return handleExpiredSession();
    $('#logs-placeholder').classList.add('hidden');
    renderError($('#logs-results'), normalizedError(error), 'logs', loadLogs);
  } finally {
    setLoading(false);
  }
}

function handleExpiredSession() {
  clearSensitiveViews();
  setLoading(false);
  showLogin();
  setLoginError('登录已过期，请重新输入访问码。');
}

function attachEvents() {
  $('#login-form').addEventListener('submit', handleLogin);
  $('#microsoft-login-button').addEventListener('click', handleMicrosoftLogin);
  $('#toggle-code').addEventListener('click', () => {
    const input = $('#access-code');
    const showing = input.type === 'text';
    input.type = showing ? 'password' : 'text';
    $('#toggle-code').textContent = showing ? '显示' : '隐藏';
    $('#toggle-code').setAttribute('aria-label', showing ? '显示访问码' : '隐藏访问码');
    $('#toggle-code').setAttribute('aria-pressed', String(!showing));
    input.focus();
  });
  $('#logout-button').addEventListener('click', handleLogout);

  $$('.nav-button').forEach((button) => button.addEventListener('click', () => navigate(button.dataset.nav)));
  $('#person-back-button').addEventListener('click', () => navigate(state.returnView || 'home'));

  $('#people-search-form').addEventListener('submit', (event) => {
    event.preventDefault();
    searchPeople();
  });
  $('#people-search').addEventListener('input', () => {
    $('#clear-people-search').classList.toggle('hidden', !$('#people-search').value);
  });
  $('#clear-people-search').addEventListener('click', () => {
    $('#people-search').value = '';
    $('#clear-people-search').classList.add('hidden');
    $('#search-results-section').classList.add('hidden');
    $('#search-welcome').classList.remove('hidden');
    $('#people-search').focus();
  });

  $('#monthly-filter-form').addEventListener('submit', (event) => {
    event.preventDefault();
    loadMonthly();
  });
  $('#stats-filter-form').addEventListener('submit', (event) => {
    event.preventDefault();
    loadStats();
  });
  $('#logs-filter-form').addEventListener('submit', (event) => {
    event.preventDefault();
    loadLogs();
  });

  $('#main-content').addEventListener('click', (event) => {
    const receiptsButton = event.target.closest('[data-receipts]');
    if (receiptsButton) {
      openReceiptList(receiptsButton);
      return;
    }
    const personButton = event.target.closest('[data-person-id]');
    if (personButton) {
      openPerson(Number(personButton.dataset.personId));
      return;
    }
    const retryButton = event.target.closest('[data-retry]');
    if (retryButton) {
      const retry = state.retryActions.get(retryButton.dataset.retry);
      if (retry) retry();
    }
  });
  $('#receipt-dialog-close').addEventListener('click', closeReceiptDialog);
  $('#receipt-dialog').addEventListener('close', clearReceiptDialogContents);
  $('#receipt-dialog').addEventListener('click', (event) => {
    const openButton = event.target.closest('[data-receipt-asset-id]');
    if (openButton) openReceiptAsset(openButton.dataset.receiptAssetId);
  });
  $('#receipt-viewer-back').addEventListener('click', () => {
    receiptRequestSequence += 1;
    clearReceiptViewer();
  });

  window.addEventListener('offline', () => toast(CLOUD_MODE ? '当前网络已断开，暂时无法读取云端资料。' : '当前网络已断开，查询功能暂不可用。', 'error'));
  window.addEventListener('online', () => toast('网络已恢复，可以继续查询。'));
  document.addEventListener('visibilitychange', () => {
    if (!CLOUD_MODE) return;
    if (document.hidden) {
      cloudHiddenAt = Date.now();
      if (cloudLockTimer) window.clearTimeout(cloudLockTimer);
      cloudLockTimer = window.setTimeout(() => {
        cloudLockTimer = null;
        lockCloudSessionAfterBackground();
      }, CLOUD_BACKGROUND_LOCK_MS);
      return;
    }
    if (cloudLockTimer) window.clearTimeout(cloudLockTimer);
    cloudLockTimer = null;
    if (cloudHiddenAt && Date.now() - cloudHiddenAt >= CLOUD_BACKGROUND_LOCK_MS) {
      lockCloudSessionAfterBackground();
    }
    cloudHiddenAt = 0;
  });
  window.addEventListener('pagehide', () => {
    if (!CLOUD_MODE) return;
    if (cloudLockTimer) window.clearTimeout(cloudLockTimer);
    cloudLockTimer = null;
    clearSensitiveViews();
    // `logout` clears the decrypted snapshot synchronously before its optional
    // auth-provider cleanup awaits. Page teardown then discards all closures.
    api.logout().catch(() => {});
    showLogin();
  });
}

function registerServiceWorker() {
  if (!('serviceWorker' in navigator) || window.location.protocol === 'file:') return;
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').catch(() => {
      // The query UI remains available even if PWA installation is unsupported.
    });
  });
}

async function restoreSession() {
  if (DEMO_MODE) return;
  setLoading(true, '正在检查登录状态');
  try {
    const session = await api.session();
    if (!session?.authenticated) return;
    state.groups = await api.groups();
    populateGroupSelect();
    showAuthenticatedApp();
  } catch (_error) {
    // A missing or unavailable session simply leaves the access-code screen open.
  } finally {
    setLoading(false);
  }
}

async function openDemoPreview() {
  if (!DEMO_MODE || DEMO_SCREEN === 'login') return;
  state.groups = cloneDemo(demoGroups);
  populateGroupSelect();
  showAuthenticatedApp();

  if (DEMO_SCREEN === 'home') {
    $('#people-search').value = '张建国';
    $('#clear-people-search').classList.remove('hidden');
    await searchPeople();
    return;
  }
  if (DEMO_SCREEN === 'person') {
    await openPerson(101);
    return;
  }
  if (DEMO_SCREEN === 'monthly') {
    $('#monthly-year').value = '2026';
    $('#monthly-month').value = '9';
    navigate('monthly');
    await loadMonthly();
    return;
  }
  if (DEMO_SCREEN === 'stats') {
    $('#stats-year').value = '2026';
    $('#stats-month').value = '9';
    state.loaded.stats = true;
    navigate('stats');
    await loadStats();
    return;
  }
  if (DEMO_SCREEN === 'logs') {
    state.loaded.logs = true;
    navigate('logs');
    await loadLogs();
  }
}

function initialize() {
  fillDateSelects();
  attachEvents();
  $('#demo-notice').classList.toggle('hidden', !DEMO_MODE);
  if (DEMO_MODE) $('#access-code').value = '12345678';
  if (ONEDRIVE_MODE) {
    $('.login-lead').textContent = '先登录 Microsoft 个人账号，再输入8位访问码查询 OneDrive 中的最新资料。';
    $('#login-hint').textContent = api.hasPairing()
      ? '请输入电脑端显示的8位访问码。配对密钥只会以访问码加密后保存在本机。'
      : '本手机尚未配对，请用电脑端 OneDrive 二维码重新打开。';
    updateMicrosoftLoginUi();
    if (oneDriveStartupError) setLoginError(oneDriveStartupError.message);
  } else if (CLOUD_MODE) {
    $('.login-lead').textContent = '输入访问码，在手机本地解密并查询最后一次同步的资料。';
    $('#login-hint').textContent = api.hasPairing()
      ? '请输入电脑端显示的8位访问码。'
      : '本手机尚未配对，请先用电脑端二维码打开此页面。';
  }
  registerServiceWorker();
  if (DEMO_MODE) openDemoPreview();
  else restoreSession();
}

initialize();
