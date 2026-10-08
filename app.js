/* SchoolCollect · Local-first. Student and payment data never leave this browser. */
const DB_NAME = 'so-thu-hoc-sinh-hieugiang-so2-local';
const DB_VERSION = 2;
const FEES = [
  { key: 'insurance', label: 'Bảo hiểm y tế (BHYT)', short: 'BHYT' },
  { key: 'mandatory', label: 'Bảo hiểm thân thể (BHTT)', short: 'BHTT' },
  { key: 'service', label: 'Dịch vụ khác', short: 'Dịch vụ' },
  { key: 'other', label: 'Chưa phân loại', short: 'Chưa rõ' }
];
let db;
let activeImport = null;
let toastTimer;
let showIncompleteOnly=false;

const $ = (s, root = document) => root.querySelector(s);
const $$ = (s, root = document) => [...root.querySelectorAll(s)];
const num = value => Math.round(Number(value) || 0);
const money = n => new Intl.NumberFormat('vi-VN').format(num(n)) + ' ₫';
const dateTime = value => value ? new Intl.DateTimeFormat('vi-VN', { dateStyle: 'short', timeStyle: 'short' }).format(new Date(value)) : '—';
const escapeHTML = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const slug = value => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/đ/g, 'd').replace(/[^a-z0-9]+/g, ' ').trim();

function openDatabase() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const d = req.result;
      if (!d.objectStoreNames.contains('students')) d.createObjectStore('students', { keyPath: 'code' });
      if (!d.objectStoreNames.contains('transactions')) d.createObjectStore('transactions', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('history')) d.createObjectStore('history', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('meta')) d.createObjectStore('meta', { keyPath: 'key' });
      if (!d.objectStoreNames.contains('receipts')) d.createObjectStore('receipts', { keyPath: 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
function request(store, method, value) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, method === 'clear' ? 'readwrite' : value === undefined ? 'readonly' : 'readwrite');
    const req = value === undefined ? tx.objectStore(store)[method]() : tx.objectStore(store)[method](value);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
function all(store) { return request(store, 'getAll'); }
async function putMany(store, entries) {
  if (!entries.length) return;
  await new Promise((resolve, reject) => {
    const tx = db.transaction(store, 'readwrite');
    const objectStore = tx.objectStore(store);
    entries.forEach(entry => objectStore.put(entry));
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}
async function clearAll() {
  await new Promise((resolve, reject) => {
    const tx = db.transaction(['students', 'transactions', 'history', 'meta', 'receipts'], 'readwrite');
    ['students', 'transactions', 'history', 'meta', 'receipts'].forEach(name => tx.objectStore(name).clear());
    tx.oncomplete = resolve; tx.onerror = () => reject(tx.error);
  });
}

function parseCsv(text, delimiter) {
  const rows = []; let row = []; let value = ''; let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) { if (ch === '"' && text[i + 1] === '"') { value += '"'; i++; } else if (ch === '"') quoted = false; else value += ch; }
    else if (ch === '"') quoted = true;
    else if (ch === delimiter) { row.push(value); value = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(value); rows.push(row); row = []; value = ''; }
    else value += ch;
  }
  if (value.length || row.length) { row.push(value); rows.push(row); }
  while (rows.length && rows[rows.length - 1].every(cell => !String(cell).trim())) rows.pop();
  return rows;
}
function bestDelimiter(text) {
  const first = text.replace(/^\uFEFF/, '').split(/\r?\n/, 1)[0] || '';
  return [',', ';', '\t'].map(d => [d, first.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
}
async function readZipEntry(bytes, entry) {
  let offset = entry.localOffset;
  if (bytes[offset] !== 0x50 || bytes[offset + 1] !== 0x4b) throw new Error('Tệp Excel không đúng định dạng ZIP/XLSX.');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const nameLength = view.getUint16(offset + 26, true), extraLength = view.getUint16(offset + 28, true);
  offset += 30 + nameLength + extraLength;
  const data = bytes.slice(offset, offset + entry.compressedSize);
  if (entry.method === 0) return data;
  if (entry.method !== 8 || !('DecompressionStream' in window)) throw new Error('Trình duyệt chưa hỗ trợ đọc XLSX ngoại tuyến. Hãy dùng Excel .xlsx hoặc xuất CSV.');
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function parseXlsx(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65558); i--) if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('Không tìm thấy cấu trúc XLSX trong tệp.');
  const count = view.getUint16(eocd + 10, true), cdOffset = view.getUint32(eocd + 16, true);
  const entries = new Map(); let p = cdOffset;
  for (let i = 0; i < count; i++) {
    if (view.getUint32(p, true) !== 0x02014b50) break;
    const method = view.getUint16(p + 10, true), compressedSize = view.getUint32(p + 20, true);
    const nameLength = view.getUint16(p + 28, true), extraLength = view.getUint16(p + 30, true), commentLength = view.getUint16(p + 32, true);
    const localOffset = view.getUint32(p + 42, true);
    const name = new TextDecoder().decode(bytes.slice(p + 46, p + 46 + nameLength));
    entries.set(name, { method, compressedSize, localOffset }); p += 46 + nameLength + extraLength + commentLength;
  }
  const getText = async name => entries.has(name) ? new TextDecoder().decode(await readZipEntry(bytes, entries.get(name))) : '';
  const workbook = new DOMParser().parseFromString(await getText('xl/workbook.xml'), 'application/xml');
  const relations = new DOMParser().parseFromString(await getText('xl/_rels/workbook.xml.rels'), 'application/xml');
  const sharedXml = await getText('xl/sharedStrings.xml');
  const sharedDoc = sharedXml ? new DOMParser().parseFromString(sharedXml, 'application/xml') : null;
  const shared = sharedDoc ? [...sharedDoc.querySelectorAll('si')].map(si => [...si.querySelectorAll('t')].map(t => t.textContent).join('')) : [];
  const headerWords=['ma hoc sinh','ma hs theo khoan nop','ma khach hang','ma moet','ma so cd','ho va ten','ho ten','ten khach hang','ten lop hoc','lop','phan loai hs','bao hiem y te','bao hiem than the','ghi chu','khoan nop','so tien','so hoa don','ngay giao dich','trang thai giao dich','ngay thang nam sinh','ngay sinh','di dong'];
  const candidates=[];
  for (const sheetNode of workbook.querySelectorAll('sheet')) {
    const relId=sheetNode.getAttribute('r:id');
    const rel=[...relations.querySelectorAll('Relationship')].find(node=>node.getAttribute('Id')===relId);
    const target=rel?.getAttribute('Target')||'worksheets/sheet1.xml';
    const sheetPath=target.startsWith('/')?target.slice(1):`xl/${target.replace(/^\.\//,'')}`;
    const sheetXml=await getText(sheetPath);if(!sheetXml)continue;
    const doc=new DOMParser().parseFromString(sheetXml,'application/xml');
    const rows=[...doc.querySelectorAll('sheetData row')].map(rowNode=>{
      const cells=[];
      for(const cell of rowNode.querySelectorAll(':scope > c')){
        const ref=cell.getAttribute('r')||'';const col=[...ref.matchAll(/[A-Z]+/g)][0]?.[0]||'A';
        let index=0;for(const char of col)index=index*26+char.charCodeAt(0)-64;index--;
        const type=cell.getAttribute('t');let val=cell.querySelector('v')?.textContent??'';
        if(type==='s')val=shared[Number(val)]??'';else if(type==='inlineStr')val=[...cell.querySelectorAll('is t')].map(t=>t.textContent).join('');
        cells[index]=val;
      }
      return cells.map(value=>value??'');
    });
    let bestHeaderIndex=0,bestScore=-1;
    rows.slice(0,25).forEach((row,index)=>{
      const headers=(row||[]).map(x=>slug(x));
      const score=headers.reduce((n,h)=>n+(h&&headerWords.some(w=>h===w||h.includes(w)||w.includes(h))?1:0),0);
      if(score>bestScore){bestScore=score;bestHeaderIndex=index;}
    });
    const dataRows=rows.slice(bestHeaderIndex+1).filter(row=>row.some(value=>String(value??'').trim())).length;
    candidates.push({name:sheetNode.getAttribute('name')||'Trang tính',rows,score:bestScore,dataRows,headerIndex:bestHeaderIndex});
  }
  const chosen=candidates.sort((a,b)=>b.score-a.score||b.dataRows-a.dataRows)[0];
  if(!chosen)throw new Error('Không tìm thấy trang tính có dữ liệu trong file Excel.');
  const normalizedRows=chosen.rows.slice(chosen.headerIndex||0);
  normalizedRows.sourceSheet=chosen.name;
  normalizedRows.headerRow=(chosen.headerIndex||0)+1;
  return normalizedRows;
}
async function readRows(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'xlsx') return parseXlsx(file);
  const buffer = await file.arrayBuffer(); let text;
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(buffer); }
  catch { text = new TextDecoder('windows-1258').decode(buffer); }
  const parsed=parseCsv(text.replace(/^\uFEFF/, ''), bestDelimiter(text));
  let best=0,bestScore=-1;
  parsed.slice(0,25).forEach((row,index)=>{const hs=(row||[]).map(slug);let score=0;if(hs.some(h=>h&&['ho ten','ho va ten','ten hoc sinh'].some(x=>h===x||h.includes(x))))score+=4;if(hs.some(h=>h&&['lop','khoi lop','ten lop hoc'].some(x=>h===x||h.includes(x))))score+=3;if(hs.some(h=>h&&['ma moet','ma hoc sinh','ma hs','ma dinh danh','ma so cd'].some(x=>h===x||h.includes(x))))score+=4;if(hs.some(h=>h&&['bao hiem y te','bao hiem than the'].some(x=>h===x||h.includes(x))))score+=2;if(score>bestScore){bestScore=score;best=index;}});
  const normalized=bestScore>=5?parsed.slice(best):parsed;normalized.headerRow=(bestScore>=5?best:0)+1;return normalized;
}
function guessColumn(headers, field) {
  const normal = headers.map(h => slug(h));
  const patterns = {
    code: ['ma moet','ma hoc sinh','ma hs','student code','student id','ma dinh danh','ma so hs'],
    externalCode: ['ma so cd','ma so cong dan','ma cd','ma so cd neu co khong co bo trong'],
    name: ['ho va ten','ho ten','ten hoc sinh','student name','ten'],
    className: ['ten lop hoc','lop','khoi lop','class','ma lop'],
    studentType: ['phan loai hs','phan loai hoc sinh','doi tuong hs'],
    registrationInfo: ['dang ky neu co cac khoan hs dang ky hoc hoac dung de thu','cac khoan hs dang ky hoc hoac dung de thu','dang ky'],
    note: ['ghi chu','note','notes'],
    personalId: ['sdd ca nhan','so dinh danh ca nhan','cccd','cmnd','ma dinh danh ca nhan'],
    gender: ['gioi tinh','gender'],
    birthDate: ['ngay thang nam sinh','ngay sinh','nam sinh','date of birth'],
    ethnicity: ['dan toc','ethnicity'],
    fatherName: ['ten cha','ho ten cha','cha'],
    motherName: ['ten me','ho ten me','me'],
    phone: ['di dong','dien thoai','so dien thoai','sdt','phone'],
    due: ['so tien phai thu', 'phai thu', 'hoc phi', 'muc thu', 'tien thu'],
    dueInsurance: ['bhyt', 'bao hiem y te', 'so tien bao hiem y te', 'bao hiem y te phai thu', 'insurance'],
    dueMandatory: ['bhtt', 'bao hiem than the', 'so tien bao hiem than the', 'bao hiem than the phai thu', 'bao hiem bat buoc', 'bat buoc'],
    dueService: ['dich vu khac phai thu', 'so tien dich vu khac', 'dich vu khac', 'service due'],
    dueParking: ['gui xe phai thu', 'so tien gui xe', 'phi gui xe', 'gui xe'],
    dueWater: ['nuoc uong phai thu', 'so tien nuoc uong', 'phi nuoc uong', 'nuoc uong'],
    studentFeeAmount: ['so tien', 'so tien phai thu', 'fee amount'],
    studentFeeCategory: ['khoan nop', 'loai khoan nop', 'fee item'],
    paymentCode: ['ma hs theo khoan nop', 'ma khach hang', 'ma thanh toan', 'customer code'],
    reportPaymentCode: ['ma khach hang','so tai khoan dinh danh','ma thanh toan'],
    reportMoet: ['thong tin bo sung 2','ma moet','ma hoc sinh','ma hs'],
    reportClass: ['dia chi','lop','khoi lop'],
    reportPersonalId: ['thong tin bo sung 1','sdd ca nhan','so dinh danh ca nhan','ma dinh danh ca nhan'],
    reportCustomerName: ['ten khach hang','ten tai khoan dinh danh'],
    serviceLevel1: ['dich vu cap 1'],
    serviceLevel2: ['dich vu cap 2','khoan nop','khoan thu'],
    invoiceId: ['ma hoa don','so hoa don'],
    amount: ['so tien hoa don','so tien giao dich','amount','credit','ghi co','so tien'],
    txnId: ['so tham chieu','ma tham chieu','ma giao dich','transaction id','reference','trace','so but toan','ma hoa don','so hoa don'],
    date: ['ngay thanh toan','ngay giao dich','thoi gian','transaction date','ngay hach toan','ngay'],
    content: ['noi dung thanh toan','noi dung chuyen khoan','dien giai','transaction content','description','chi tiet','ten khach hang'],
    studentCode: ['ma moet','ma hoc sinh','ma hs theo khoan nop','ma khach hang','ma hs','student code','student id','ma dinh danh'],
    feeCategory: ['dich vu cap 2','khoan nop','khoan thu','loai khoan thu','danh muc thu','fee category','fee type'],
    bankStatus: ['trang thai','trang thai giao dich','status']
  }[field] || [];
  for(const pattern of patterns){const i=normal.indexOf(pattern);if(i>=0)return String(i);}
  for(const pattern of patterns){const i=normal.findIndex(h=>h&&(h.includes(pattern)||pattern.includes(h)));if(i>=0)return String(i);}
  return '-1';
}
function fieldSelect(id, label, headers, field, required = false) {
  const guessed = guessColumn(headers, field);
  return `<div class="mapping-field"><label for="${id}">${label}${required ? ' <em>*</em>' : ''}</label><select id="${id}" data-field="${field}"><option value="-1">— Không chọn —</option>${headers.map((h, i) => `<option value="${i}" ${String(i) === guessed ? 'selected' : ''}>${escapeHTML(h || `Cột ${i + 1}`)}</option>`).join('')}</select></div>`;
}
function previewHtml(headers, rows) {
  return `<div class="preview-box"><strong>Xem trước 3 dòng đầu</strong><div class="table-wrap"><table><thead><tr>${headers.map(h => `<th>${escapeHTML(h)}</th>`).join('')}</tr></thead><tbody>${rows.slice(1, 4).map(row => `<tr>${headers.map((_, i) => `<td>${escapeHTML(row[i] || '')}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;
}
function openImportModal(kind, file, rows) {
  if (!rows.length || rows.length < 2) return toast('File không có dòng tiêu đề hoặc dữ liệu.', true);
  const headers = rows[0].map((h, i) => String(h || `Cột ${i + 1}`).trim());
  const longFormat=kind==='students'&&['code','paymentCode','studentFeeCategory','studentFeeAmount'].every(field=>Number(guessColumn(headers,field))>=0);
  activeImport = { kind, file, rows, headers, longFormat };
  $('#modalEyebrow').textContent = kind === 'students' ? 'DANH SÁCH HỌC SINH' : 'BÁO CÁO THU';
  $('#modalTitle').textContent = kind === 'students' ? 'Ghép cột danh sách học sinh' : 'Ghép cột báo cáo thu';
  const summary = `<div class="file-summary"><span class="file-badge">${file.name.toLowerCase().endsWith('.xlsx') ? 'XLSX' : 'CSV'}</span><div><strong>${escapeHTML(file.name)}</strong><small>${rows.length - 1} dòng dữ liệu · ${headers.length} cột${rows.sourceSheet?` · Trang tính: ${escapeHTML(rows.sourceSheet)}`:''}</small></div></div>`;
  const fields = kind === 'students'
    ? longFormat
      ? `<div class="mapping-grid">${fieldSelect('map-code','Mã học sinh',headers,'code',true)}${fieldSelect('map-payment-code','Mã HS theo khoản nộp',headers,'paymentCode',true)}${fieldSelect('map-name','Họ và tên',headers,'name',true)}${fieldSelect('map-class','Lớp',headers,'className',true)}${fieldSelect('map-fee-category','Khoản nộp (BHYT/BHTT)',headers,'studentFeeCategory',true)}${fieldSelect('map-fee-amount','Số tiền phải thu',headers,'studentFeeAmount',true)}</div><p class="mapping-intro">Mỗi học sinh có một dòng BHYT và một dòng BHTT. Web tự ghép hai dòng theo mã học sinh và giữ mã từng khoản để đối soát với cột “Mã khách hàng” của ngân hàng.</p>`
      : `<div class="mapping-grid">${fieldSelect('map-code','Mã học sinh nội bộ (nếu file đã có)',headers,'code')}${fieldSelect('map-external-code','Mã số CĐ / mã ngoài (nếu có)',headers,'externalCode')}${fieldSelect('map-name','Họ và tên',headers,'name',true)}${fieldSelect('map-class','Lớp',headers,'className',true)}${fieldSelect('map-gender','Giới tính',headers,'gender')}${fieldSelect('map-birth-date','Ngày sinh',headers,'birthDate')}${fieldSelect('map-student-type','Phân loại HS',headers,'studentType')}${fieldSelect('map-registration-info','Đăng ký khoản khác (nếu có)',headers,'registrationInfo')}${fieldSelect('map-due-insurance','Bảo hiểm y tế (BHYT)',headers,'dueInsurance')}${fieldSelect('map-due-mandatory','Bảo hiểm thân thể (BHTT)',headers,'dueMandatory')}${fieldSelect('map-note','Ghi chú',headers,'note')}</div><p class="mapping-intro"><strong>Mẫu THCS Số 2 Hiếu Giang:</strong> web chỉ nhận các trường thực tế có trong file trường: Mã số CĐ, Họ tên, Lớp, Giới tính, Ngày sinh, Phân loại HS, Đăng ký, BHYT, BHTT và Ghi chú. Không tự suy đoán phụ huynh, điện thoại, CCCD hoặc dữ liệu khác.</p>`
    : `<div class="mapping-grid">${fieldSelect('map-amount','Số tiền hóa đơn / giao dịch',headers,'amount',true)}${fieldSelect('map-report-moet','Mã học sinh / mã ngoài',headers,'reportMoet')}${fieldSelect('map-report-payment','Mã khách hàng / TK định danh',headers,'reportPaymentCode')}${fieldSelect('map-report-name','Tên khách hàng',headers,'reportCustomerName')}${fieldSelect('map-report-class','Lớp / địa chỉ',headers,'reportClass')}${fieldSelect('map-report-personal','SĐD cá nhân',headers,'reportPersonalId')}${fieldSelect('map-service-2','Dịch vụ cấp 2 / khoản thu',headers,'serviceLevel2')}${fieldSelect('map-content','Nội dung thanh toán',headers,'content')}${fieldSelect('map-txn','Số tham chiếu / mã giao dịch',headers,'txnId')}${fieldSelect('map-invoice','Mã hóa đơn',headers,'invoiceId')}${fieldSelect('map-date','Ngày thanh toán',headers,'date')}${fieldSelect('map-bank-status','Trạng thái giao dịch',headers,'bankStatus')}</div><p class="mapping-intro"><strong>Mẫu BIDV Hiếu Giang:</strong> web ưu tiên “Thông tin bổ sung 2” để ghép mã học sinh, dùng “Dịch vụ cấp 2”/nội dung để xác định khoản thu và “Số tham chiếu” làm mã giao dịch. “Mã khách hàng” được lưu riêng để tra cứu, không bắt buộc phải trùng mã học sinh.</p>`;
  $('#modalBody').innerHTML = `${summary}${fields}${previewHtml(headers, rows)}`;
  $('#modalConfirm').textContent = kind === 'students' ? 'Nhập danh sách' : 'Nhập & đối soát';
  $('#modalBackdrop').classList.add('open');
}
function closeModal() { $('#modalBackdrop').classList.remove('open'); activeImport = null; }
function getMap() { return Object.fromEntries($$('#modalBody select').map(s => [s.dataset.field, Number(s.value)])); }
function cell(row, map, field) { const index = map[field]; return index === undefined || index < 0 ? '' : String(row[index] ?? '').trim(); }
function parseAmount(value) {
  let s = String(value || '').replace(/[^\d,.-]/g, '').trim();
  if (!s) return 0;
  if (s.includes(',') && s.includes('.')) s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (s.includes(',')) s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
  else if ((s.match(/\./g) || []).length > 1 || /^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  return Math.round(Number(s) || 0);
}
function normalizeDate(value) {
  if (!value) return '';
  if (/^\d+(\.\d+)?$/.test(value) && Number(value) > 15000 && Number(value) < 90000) return new Date(Date.UTC(1899, 11, 30) + Number(value) * 86400000).toISOString().slice(0, 10);
  const m = value.match(/^(\d{1,2})[/. -](\d{1,2})[/. -](\d{2,4})/);
  if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  const d = new Date(value); return Number.isNaN(d.getTime()) ? value : d.toISOString().slice(0, 10);
}
function isSuccessfulBankStatus(value) {
  const status=slug(value);
  if(!status)return true;
  if(/khong|that bai|huy|tu choi|failed|cancel|pending/.test(status))return false;
  return /thanh cong|success|successful|completed|hoan tat/.test(status);
}
function getFeeKey(value, content = '') {
  const raw = slug(value || '');
  if (/bhtt|bao hiem than the|than the|bao hiem bat buoc|bat buoc|mandatory/.test(raw)) return 'mandatory';
  if (/bhyt|bao hiem y te|bao hiem|insurance|y te/.test(raw)) return 'insurance';
  if (/dich vu|gui xe|nuoc uong|ban tru|parking|service|an uong/.test(raw)) return 'service';
  const s = slug(content || '');
  if (/bhtt|bao hiem than the|than the|bao hiem bat buoc|bat buoc|mandatory/.test(s)) return 'mandatory';
  if (/bhyt|bao hiem y te|bao hiem|insurance|y te/.test(s)) return 'insurance';
  if (/dich vu|gui xe|nuoc uong|ban tru|parking|service|an uong/.test(s)) return 'service';
  return 'other';
}
function feeCategoryFromPaymentCode(value) {
  const code=String(value||'').trim().toUpperCase().replace(/\s+/g,'');
  if(!code)return 'other';
  // Mã BIDV thực tế có dạng ...BHTTHG...TT6D38 / ...BHYTHG...YT6A08,
  // vì vậy không thể chỉ kiểm tra phần kết thúc chuỗi.
  if(code.includes('BHTT')||/TT(?=[0-9A-Z])/.test(code))return 'mandatory';
  if(code.includes('BHYT')||/YT(?=[0-9A-Z])/.test(code))return 'insurance';
  return 'other';
}
function getFeeLabel(key) { return FEES.find(f => f.key === key)?.label || 'Chưa phân loại'; }
function serviceDetail(raw, content) {
  const rawText = String(raw || '').trim();
  if (rawText && !/^(dich vu khac|dich vu|bao hiem)$/i.test(slug(rawText))) return rawText;
  let detail = String(content || '').trim().replace(/^(thu|nop|chuyen khoan|thanh toan)\s+/i, '');
  detail = detail.replace(/\s*[-–—]\s*(ma\s*)?hs\s*[a-z0-9-]+.*$/i, '').replace(/\s*[-–—]\s*hs\d+.*$/i, '').trim();
  return detail ? detail.charAt(0).toLocaleUpperCase('vi-VN') + detail.slice(1) : 'Dịch vụ khác';
}
function studentFeesFromRow(row, map) {
  const hasServiceDetails=map.dueParking>=0||map.dueWater>=0;
  const fields=[['insurance',getFeeLabel('insurance'),'dueInsurance'],['mandatory',getFeeLabel('mandatory'),'dueMandatory']];
  if(hasServiceDetails){fields.push(['service','Gửi xe','dueParking'],['service','Nước uống','dueWater']);}
  else fields.push(['service','Dịch vụ khác','dueService']);
  const dueItems=fields.filter(([, ,field])=>map[field]>=0).map(([category,name,field])=>({id:`${category}:${slug(name)}`,category,name,amount:parseAmount(cell(row,map,field))})).filter(item=>item.amount>0);
  const hasBreakdown=fields.some(([, ,field])=>map[field]>=0);
  if(hasServiceDetails&&map.dueService>=0){
    const serviceTotal=parseAmount(cell(row,map,'dueService'));
    const detailedService=dueItems.filter(item=>item.category==='service').reduce((sum,item)=>sum+item.amount,0);
    if(serviceTotal>detailedService)dueItems.push({id:'service:other',category:'service',name:'Dịch vụ khác',amount:serviceTotal-detailedService});
  }
  const statedDue=parseAmount(cell(row,map,'due'));
  const detailedTotal=dueItems.reduce((sum,item)=>sum+item.amount,0);
  if(hasBreakdown&&statedDue>detailedTotal)dueItems.push({id:'other:unclassified',category:'other',name:'Chưa phân loại',amount:statedDue-detailedTotal});
  if(!hasBreakdown&&statedDue>0)dueItems.push({id:'other:unclassified',category:'other',name:'Chưa phân loại',amount:statedDue});
  const due=hasBreakdown?Math.max(statedDue,detailedTotal):statedDue;
  const dueByCategory=dueItems.reduce((out,item)=>(out[item.category]=(out[item.category]||0)+item.amount,out),{});
  return {due,dueItems,dueByCategory,hasFeeBreakdown:hasBreakdown};
}
function studentIdentity(s){
  const external=slug(String(s?.externalCode||'').replace(/\s+/g,''));
  if(external)return `ext|${external}`;
  const name=slug(s?.name||''),birth=String(s?.birthDate||''),gender=slug(s?.gender||'');
  if(name&&birth)return `nb|${name}|${birth}|${gender}`;
  return '';
}
function uniqueStudentIndex(items,keyFn){
  const out=new Map(),duplicates=new Set();
  for(const item of items){const key=keyFn(item);if(!key)continue;if(out.has(key)){out.delete(key);duplicates.add(key);}else if(!duplicates.has(key))out.set(key,item);}
  return out;
}
function studentFallbackIdentity(s){
  const name=slug(s?.name||''),birth=String(s?.birthDate||''),gender=slug(s?.gender||''),cls=slug(s?.className||'');
  if(name&&birth)return `nb|${name}|${birth}`;
  if(name&&gender&&cls)return `ngc|${name}|${gender}|${cls}`;
  return '';
}
function compareStudentImportRows(a,b,map){
  const av={
    className:cell(a,map,'className'),name:cell(a,map,'name'),
    birth:normalizeDate(cell(a,map,'birthDate')),gender:cell(a,map,'gender'),
    external:cell(a,map,'externalCode')
  };
  const bv={
    className:cell(b,map,'className'),name:cell(b,map,'name'),
    birth:normalizeDate(cell(b,map,'birthDate')),gender:cell(b,map,'gender'),
    external:cell(b,map,'externalCode')
  };
  return String(av.className||'').localeCompare(String(bv.className||''),'vi',{numeric:true,sensitivity:'base'})
    ||String(av.name||'').localeCompare(String(bv.name||''),'vi',{sensitivity:'base'})
    ||String(av.birth||'').localeCompare(String(bv.birth||''))
    ||String(av.gender||'').localeCompare(String(bv.gender||''))
    ||String(av.external||'').localeCompare(String(bv.external||''));
}
function createStudentCodeAllocator(existingList){
  const used=new Set(existingList.map(s=>slug(s.code)).filter(Boolean));let max=0;
  existingList.forEach(s=>{const m=String(s.code||'').toUpperCase().match(/^HG2-(\d+)$/);if(m)max=Math.max(max,Number(m[1])||0);});
  return {
    reserve(code){if(code)used.add(slug(code));},
    next(){let code;do{max++;code=`HG2-${String(max).padStart(5,'0')}`;}while(used.has(slug(code)));used.add(slug(code));return code;}
  };
}
function classCodeToken(className){
  return feeSafeCode(className,8)||'LOP';
}
function createClassStudentCodeAllocator(existingList){
  const used=new Set(),maxByClass=new Map();
  for(const s of existingList){
    const value=String(s.classStudentCode||'').toUpperCase();
    if(value)used.add(slug(value));
    const m=value.match(/^HG2-([A-Z0-9]+)-(\d{2})$/);
    if(m)maxByClass.set(m[1],Math.max(maxByClass.get(m[1])||0,Number(m[2])||0));
  }
  return {
    next(className){
      const cls=classCodeToken(className);let n=maxByClass.get(cls)||0,code;
      do{
        n++;
        if(n>99)throw new Error(`Lớp ${className} có hơn 99 học sinh; cần mở rộng cấu trúc mã.`);
        code=`HG2-${cls}-${String(n).padStart(2,'0')}`;
      }while(used.has(slug(code)));
      maxByClass.set(cls,n);used.add(slug(code));return code;
    },
    validForClass(code,className){
      const cls=classCodeToken(className);
      return new RegExp(`^HG2-${cls}-\\d{2}$`,'i').test(String(code||''));
    }
  };
}
function studentWarnings(s,conflicts=[]){
  const out=[];if(!s.name||/^Chưa cập nhật/.test(s.name))out.push('Thiếu họ tên');if(!s.className||s.className==='Chưa xếp lớp')out.push('Thiếu lớp');
  if(conflicts.length)out.push('Thông tin trùng/chưa thống nhất');
  return [...new Set(out)];
}
function studentsFromFeeRows(dataRows,map,now) {
  const grouped=new Map(),paymentOwners=new Map();
  for(let i=0;i<dataRows.length;i++) {
    const row=dataRows[i],code=cell(row,map,'code'),name=cell(row,map,'name'),className=cell(row,map,'className');
    const paymentCode=cell(row,map,'paymentCode').toUpperCase().replace(/\s+/g,'');
    const rawCategory=cell(row,map,'studentFeeCategory'),category=getFeeKey(rawCategory),amount=parseAmount(cell(row,map,'studentFeeAmount'));
    if(!code||!name||!className||!paymentCode||!amount||!['insurance','mandatory'].includes(category))
      throw new Error(`Dòng ${i+2}: thiếu mã, tên, lớp, khoản BHYT/BHTT hoặc số tiền hợp lệ.`);
    if(feeCategoryFromPaymentCode(paymentCode)!==category)throw new Error(`Dòng ${i+2}: mã khoản nộp ${paymentCode} không khớp loại ${rawCategory}.`);
    let student=grouped.get(slug(code));
    if(!student){student={code,name,className,due:0,dueItems:[],dueByCategory:{insurance:0,mandatory:0,service:0,other:0},updatedAt:now,hasFeeBreakdown:true};grouped.set(slug(code),student);}
    if(student.name!==name||student.className!==className)throw new Error(`Mã học sinh ${code} có tên hoặc lớp không thống nhất trong file.`);
    const normalizedPaymentCode=slug(paymentCode);
    const owner=paymentOwners.get(normalizedPaymentCode);
    if(owner&&owner!==slug(code))throw new Error(`Mã HS theo khoản nộp ${paymentCode} bị gán cho nhiều học sinh.`);
    paymentOwners.set(normalizedPaymentCode,slug(code));
    if(student.dueItems.some(item=>slug(item.paymentCode)===normalizedPaymentCode))throw new Error(`Mã HS theo khoản nộp ${paymentCode} bị lặp. Hãy kiểm tra file trước khi nhập.`);
    student.dueItems.push({id:`payment:${normalizedPaymentCode}`,paymentCode,category,name:getFeeLabel(category),amount});
    student.due+=amount;student.dueByCategory[category]+=amount;
  }
  for(const student of grouped.values()){
    const counts=student.dueItems.reduce((out,item)=>(out[item.category]=(out[item.category]||0)+1,out),{});
    if(counts.insurance!==1||counts.mandatory!==1)throw new Error(`Học sinh ${student.code} cần đúng một dòng BHYT và một dòng BHTT trong file.`);
  }
  if(!grouped.size)throw new Error('Không tìm thấy dòng học sinh hợp lệ trong file.');
  return [...grouped.values()];
}
async function confirmImport() {
  if (!activeImport) return;
  const { kind, file, rows, longFormat } = activeImport; const map = getMap();
  if (kind === 'students' && (map.name < 0 || map.className < 0 || (longFormat && (map.code < 0 || [map.paymentCode,map.studentFeeCategory,map.studentFeeAmount].some(x=>x<0))))) return toast(longFormat?'Hãy chọn đủ cột mã, tên, lớp, mã khoản nộp, loại khoản và số tiền.':'Hãy chọn ít nhất cột Họ tên và Lớp.', true);
  if (kind === 'bank' && map.amount < 0) return toast('Hãy chọn cột số tiền giao dịch.', true);
  const dataRows = rows.slice(1).filter(row => row.some(v => String(v ?? '').trim()));
  const now = new Date().toISOString(); let summary;
  if (kind === 'students') {
    let items;
    try {
      if(longFormat) items=studentsFromFeeRows(dataRows,map,now);
      else {
        const existingList=await all('students');
        const existing=new Map(existingList.map(s=>[slug(s.code),s]));
        const byIdentity=uniqueStudentIndex(existingList,studentIdentity);
        const byFallbackIdentity=uniqueStudentIndex(existingList,studentFallbackIdentity);
        const byExternal=uniqueStudentIndex(existingList,s=>slug(String(s.externalCode||'').replace(/\s+/g,'')));
        const allocator=createStudentCodeAllocator(existingList);
        const classAllocator=createClassStudentCodeAllocator(existingList);
        dataRows.forEach(row=>allocator.reserve(cell(row,map,'code')));
        const feeMapped=['due','dueInsurance','dueMandatory','dueService','dueParking','dueWater'].some(field=>Number(map[field])>=0);
        const merged=new Map(),warnings=[];const tempToDelete=new Set();
        const orderedRows=[...dataRows].sort((a,b)=>compareStudentImportRows(a,b,map));
        orderedRows.forEach((row,index)=>{
          let code=cell(row,map,'code'),name=cell(row,map,'name'),className=cell(row,map,'className');
          const profile={
            externalCode:cell(row,map,'externalCode'),
            gender:cell(row,map,'gender'),
            birthDate:normalizeDate(cell(row,map,'birthDate')),
            studentType:cell(row,map,'studentType'),
            registrationInfo:cell(row,map,'registrationInfo'),
            note:cell(row,map,'note')
          };
          if(!code&&!name&&!className&&!Object.values(profile).some(Boolean))return;
          const candidate={name,className,birthDate:profile.birthDate,gender:profile.gender,externalCode:profile.externalCode};
          const identity=studentIdentity(candidate),fallbackIdentity=studentFallbackIdentity(candidate);
          const externalKey=slug(String(profile.externalCode||'').replace(/\s+/g,''));
          let old=code?existing.get(slug(code)):null;
          if(!old&&externalKey)old=byExternal.get(externalKey)||null;
          if(!old&&identity)old=byIdentity.get(identity)||null;
          if(!old&&fallbackIdentity)old=byFallbackIdentity.get(fallbackIdentity)||null;
          if(!code&&old?.code&&!String(old.code).startsWith('TMP-'))code=old.code;
          if(!code)code=allocator.next();
          if(old&&old.code!==code&&String(old.code).startsWith('TMP-'))tempToDelete.add(old.code);
          const missing=[];if(!name)missing.push('họ tên');if(!className)missing.push('lớp');
          if(!name)name=old?.name||'Chưa cập nhật họ tên';
          if(!className)className=old?.className||'Chưa xếp lớp';
          const key=slug(code),prior=merged.get(key)||old||{};
          const conflicts=[];
          if(prior.name&&name&&prior.name!==name&&!/^Chưa cập nhật/.test(prior.name))conflicts.push(`họ tên: “${prior.name}” / “${name}”`);
          if(prior.externalCode&&profile.externalCode&&slug(prior.externalCode)!==slug(profile.externalCode))conflicts.push(`Mã số CĐ: “${prior.externalCode}” / “${profile.externalCode}”`);
          if(merged.has(key))warnings.push(`Dòng ${index+2}: trùng mã ${code}, web đã gộp dữ liệu.`);
          if(conflicts.length)warnings.push(`Mã ${code}: thông tin chưa thống nhất (${conflicts.join('; ')}).`);
          if(missing.length)warnings.push(`Dòng ${index+2}: còn thiếu ${missing.join(', ')}; vẫn được nhập để bổ sung sau.`);
          const choose=(fresh,oldValue)=>fresh||oldValue||'';
          const fees=feeMapped?studentFeesFromRow(row,map):null;
          const effectiveClass=choose(className,prior.className);
          const classStudentCode=classAllocator.validForClass(prior.classStudentCode,effectiveClass)?prior.classStudentCode:classAllocator.next(effectiveClass);
          const next={
            ...prior,code,classStudentCode,
            name:choose(name,prior.name),
            className:effectiveClass,
            externalCode:choose(profile.externalCode,prior.externalCode),
            gender:choose(profile.gender,prior.gender),
            birthDate:choose(profile.birthDate,prior.birthDate),
            studentType:choose(profile.studentType,prior.studentType),
            registrationInfo:choose(profile.registrationInfo,prior.registrationInfo),
            note:choose(profile.note,prior.note),
            personalId:'',ethnicity:'',fatherName:'',motherName:'',phone:'',
            due:feeMapped?fees.due:(num(prior.due)||0),
            dueItems:feeMapped?fees.dueItems:(Array.isArray(prior.dueItems)?prior.dueItems:[]),
            dueByCategory:feeMapped?fees.dueByCategory:(prior.dueByCategory||{insurance:0,mandatory:0,service:0,other:0}),
            hasFeeBreakdown:feeMapped?fees.hasFeeBreakdown:!!prior.hasFeeBreakdown,
            identityKey:studentIdentity({name:choose(name,prior.name),className:effectiveClass,birthDate:choose(profile.birthDate,prior.birthDate),gender:choose(profile.gender,prior.gender),externalCode:choose(profile.externalCode,prior.externalCode)}),
            identityLockedAt:prior.identityLockedAt||now,
            updatedAt:now
          };
          next.dataWarnings=studentWarnings(next,conflicts);
          merged.set(key,next);
          existing.set(key,next);
          if(externalKey)byExternal.set(externalKey,next);
          const nextIdentity=studentIdentity(next);if(nextIdentity)byIdentity.set(nextIdentity,next);
          const nextFallback=studentFallbackIdentity(next);if(nextFallback)byFallbackIdentity.set(nextFallback,next);
        });
        for(const oldCode of tempToDelete)await request('students','delete',oldCode);
        items=[...merged.values()];
        activeImport.warnings=warnings;
      }
    }
    catch(error) { return toast(error.message||'Danh sách học sinh chưa đúng định dạng.',true); }
    await putMany('students', items);
    const currentStudents=await all('students');
    const codes = new Map(currentStudents.map(s => [slug(s.code), s]));
    const paymentOwners=new Map();currentStudents.forEach(s=>studentDueItems(s).forEach(item=>{if(item.paymentCode)paymentOwners.set(slug(item.paymentCode),s);}));
    const knownTransactions = await all('transactions');
    const rematched = knownTransactions.map(t => {
      const reported=t.reportedPaymentCode||t.reportedStudentCode||'';
      const student = reported ? codes.get(slug(reported))||paymentOwners.get(slug(reported)) : null;
      return student ? { ...t, studentCode:student.code, studentName:student.name, matched:true } : { ...t, studentCode:'', studentName:'', matched:false };
    });
    await putMany('transactions', rematched);
    const feeItems=items.reduce((sum,s)=>sum+studentDueItems(s).length,0);
    const warningCount=(activeImport.warnings||[]).length;
    summary = { rows:dataRows.length, imported:items.length, detail:`${items.length.toLocaleString('vi-VN')} học sinh được nhập/cập nhật${warningCount?` · ${warningCount} cảnh báo cần rà soát`:''}`, warnings:activeImport.warnings||[] };
    $('#studentLastImport').textContent = `Gần nhất: ${file.name} · ${items.length.toLocaleString('vi-VN')} học sinh`;
  } else {
    const students = await all('students'); const byCode = new Map(students.map(s => [slug(s.code), s]));
    const byPaymentCode=new Map();students.forEach(s=>studentDueItems(s).forEach(item=>{if(item.paymentCode)byPaymentCode.set(slug(item.paymentCode),s);if(item.qrAccountNumber)byPaymentCode.set(slug(item.qrAccountNumber),s);}));
    const items = dataRows.map((row, i) => {
      const amount = parseAmount(cell(row, map, 'amount')); if (!amount) return null;
      const reportedMoet=cell(row,map,'reportMoet')||cell(row,map,'studentCode');
      const paymentCode=cell(row,map,'reportPaymentCode')||cell(row,map,'studentCode');
      const reportName=cell(row,map,'reportCustomerName');
      const reportClass=cell(row,map,'reportClass');
      const reportPersonalId=cell(row,map,'reportPersonalId');
      const student = (reportedMoet?byCode.get(slug(reportedMoet)):null) || (paymentCode?byPaymentCode.get(slug(paymentCode)):null) || (paymentCode?byCode.get(slug(paymentCode)):null);
      const date = normalizeDate(cell(row, map, 'date')); const content = cell(row, map, 'content');
      const service2=cell(row,map,'serviceLevel2'); const rawCategory=service2||cell(row,map,'feeCategory');
      const codeCategory=feeCategoryFromPaymentCode(paymentCode); const feeCategory = getFeeKey(rawCategory,content)!=='other'?getFeeKey(rawCategory,content):(codeCategory!=='other'?codeCategory:getFeeKey('',content));
      const feeDetail = feeCategory === 'service' ? serviceDetail(rawCategory, content) : getFeeLabel(feeCategory);
      const bankStatus=cell(row,map,'bankStatus');
      const invoiceId=cell(row,map,'invoiceId');
      const ref = cell(row, map, 'txnId')||invoiceId; const normalizedRef = slug(ref);
      const id = normalizedRef ? `ref:${normalizedRef}` : `row:${slug(date)}:${amount}:${slug(paymentCode)}:${slug(reportedMoet)}:${slug(content)}`;
      return { id, ref, invoiceId, date, content, amount, feeCategory, feeDetail, reportedStudentCode:reportedMoet, reportedPaymentCode:paymentCode, reportCustomerName:reportName, reportClass, reportPersonalId, serviceLevel2:service2, bankStatus, studentCode:student?.code || '', studentName:student?.name || reportName || '', sourceFile:file.name, importedAt:now, matched:!!student, paymentChannel:'transfer' };
    }).filter(Boolean);
    // Mỗi món thu có thể được BIDV xuất thành một file riêng (BHYT, BHTT, dịch vụ...).
    // Vì vậy không được xóa toàn bộ giao dịch ngân hàng khi nhập file tiếp theo.
    // Số tham chiếu/mã giao dịch là khóa chính: giao dịch đã có được cập nhật,
    // giao dịch mới được bổ sung, dữ liệu các món thu khác được giữ nguyên.
    const uniqueById=new Map();
    let duplicates=0;
    for(const t of items){
      if(uniqueById.has(t.id)) duplicates++;
      uniqueById.set(t.id,t);
    }
    const incoming=[...uniqueById.values()];
    const existing=await all('transactions');
    const existingBank=existing.filter(t=>t.sourceType!=='cash');
    const manualCash=existing.filter(t=>t.sourceType==='cash');
    const oldById=new Map(existingBank.map(t=>[t.id,t]));
    let added=0,updated=0,changed=0;
    for(const t of incoming){
      const old=oldById.get(t.id);
      if(!old){added++;continue;}
      updated++;
      const compareFields=['ref','invoiceId','date','amount','reportedStudentCode','reportedPaymentCode','serviceLevel2','bankStatus','content'];
      if(compareFields.some(k=>String(old[k]??'')!==String(t[k]??'')))changed++;
    }
    // putMany ghi đè đúng record có cùng id nhưng không đụng đến record của file/món thu khác.
    await putMany('transactions',incoming);
    const after=await all('transactions');
    const bankTotal=after.filter(t=>t.sourceType!=='cash').length;
    const feeLabels=[...new Set(incoming.map(t=>String(t.serviceLevel2||t.feeDetail||getFeeLabel(t.feeCategory)||'').trim()).filter(Boolean))];
    const feeText=feeLabels.length===1?feeLabels[0]:(feeLabels.length>1?`${feeLabels.length} nhóm khoản thu`:'báo cáo thu');
    summary = {
      rows:dataRows.length,
      imported:incoming.length,
      detail:`${incoming.length} giao dịch ${feeText} được kiểm tra · ${added} mới · ${updated} đã có/cập nhật${changed?` (${changed} thay đổi)`:''} · giữ ${Math.max(0,existingBank.length-updated)} giao dịch từ các lần nhập trước · tổng ${bankTotal} giao dịch ngân hàng${manualCash.length?` + ${manualCash.length} tiền mặt`:''}${duplicates?` · ${duplicates} dòng trùng trong file được gộp`:''}`
    };
    $('#bankLastImport').textContent = `Gần nhất: ${file.name} · tổng ${bankTotal.toLocaleString('vi-VN')} giao dịch ngân hàng`;
  }
  await request('history', 'put', { id:crypto.randomUUID(), kind:kind === 'students' ? 'Danh sách học sinh' : 'Báo cáo thu', fileName:file.name, rows:summary.rows, imported:summary.imported, detail:summary.detail, at:now });
  closeModal(); await refresh(); toast(summary.detail);
}
function toast(message, error = false) {
  const el = $('#toast'); el.textContent = message; el.classList.toggle('error', error); el.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
}
async function chooseFile(kind, file) {
  if (!file) return;
  try { const rows = await readRows(file); openImportModal(kind, file, rows); }
  catch (error) { console.error(error); toast(error.message || 'Không đọc được file này.', true); }
}
function setPage(page) {
  if (page === 'insurance') page = 'dashboard';
  $$('.page').forEach(p => p.classList.toggle('active', p.id === `page-${page}`));
  $$('.nav-item[data-page]').forEach(b => b.classList.toggle('active', b.dataset.page === page));
  const labels = {
    dashboard:['Tổng quan','Theo dõi tiến độ thu theo thời gian thực trên thiết bị này'],
    students:['Học sinh','Danh sách và số phải thu chi tiết theo từng học sinh'],
    fees:['Khoản thu','Theo dõi riêng bảo hiểm, dịch vụ khác và từng nội dung dịch vụ'],
    'fee-setup':['Thiết lập khoản thu','Tạo và phân giao khoản thu theo toàn trường, khối, lớp hoặc học sinh'],
    'bidv-export':['Bảng kê BIDV','Tạo mã khách hàng duy nhất và xuất XLSX nhập chương trình thu hộ'],
    notices:['Thông báo nộp tiền','Xuất A4/PDF và ảnh QR hàng loạt gửi phụ huynh'],
    receipts:['Phiếu thu / Xác nhận','Phát hành chứng từ từ các món đã đối soát thành công'],
    qr:['Tạo mã QR','Tạo QR thanh toán theo từng món thu của từng học sinh'],
    imports:['Nhập dữ liệu','Cập nhật cộng dồn nhiều báo cáo thu theo từng khoản, tự kiểm tra trùng giao dịch'],
    history:['Tra cứu & báo cáo','Lịch sử các lần nhập dữ liệu trên thiết bị này'],
    settings:['Sao lưu & cài đặt','Bảo vệ và chuyển dữ liệu theo quy trình của trường']
  };
  $('#topTitle').textContent = labels[page]?.[0] || 'SchoolCollect';
  $('#topSubtitle').textContent = labels[page]?.[1] || '';
  $('#sidebar').classList.remove('open'); window.scrollTo({ top:0, behavior:'smooth' });
}
function transactionCategory(t) {
  const codeCategory=feeCategoryFromPaymentCode(t.reportedPaymentCode||t.reportedStudentCode);
  if(codeCategory!=='other')return codeCategory;
  return ['insurance','mandatory','service','other'].includes(t.feeCategory) ? t.feeCategory : getFeeKey('', t.content);
}
function studentDueItems(student) {
  if (Array.isArray(student.dueItems)) return student.dueItems.filter(x=>num(x.amount)>0).map(x=>({...x,amount:num(x.amount)}));
  if (student.dueByCategory && Object.keys(student.dueByCategory).length) {
    const items=[];
    for (const key of ['insurance','mandatory','service','other']) { const amount=num(student.dueByCategory[key]); if(amount) items.push({id:`legacy:${key}`,category:key,name:key==='other'?'Chưa phân loại':getFeeLabel(key),amount,legacy:true}); }
    const gap=Math.max(0,num(student.due)-items.reduce((sum,x)=>sum+x.amount,0));
    if(gap)items.push({id:'legacy:unclassified',category:'other',name:'Chưa phân loại',amount:gap});
    return items;
  }
  return num(student.due)>0?[{id:'legacy:unclassified',category:'other',name:'Chưa phân loại',amount:num(student.due)}]:[];
}
function studentDueByCategory(student) {
  return studentDueItems(student).reduce((out,item)=>(out[item.category]=(out[item.category]||0)+item.amount,out),{insurance:0,mandatory:0,service:0,other:0});
}
let noticeBundleCache=new Map();
function setNoticeBundleCache(record){
  const items=Array.isArray(record?.items)?record.items:[];
  noticeBundleCache=new Map(items.map(x=>[slug(x.remark),x]).filter(([k])=>k));
}
function transactionMatchedItemIds(t){
  const ids=Array.isArray(t?.matchedDueItemIds)?t.matchedDueItemIds.filter(Boolean):[];
  if(t?.matchedDueItemId&&!ids.includes(t.matchedDueItemId))ids.push(t.matchedDueItemId);
  return ids;
}
function transactionMatchesItem(t,itemId){return transactionMatchedItemIds(t).includes(itemId);}
function transactionPaidKeys(t){
  return transactionMatchedItemIds(t).map(id=>`${t.studentCode}|${id}`);
}
function findUniqueSubsetByAmount(items,target,maxSolutions=2){
  const usable=items.filter(x=>num(x.amount)>0).slice(0,14),solutions=[];
  function walk(index,sum,picked){
    if(solutions.length>=maxSolutions||sum>target)return;
    if(sum===target&&picked.length){solutions.push([...picked]);return;}
    if(index>=usable.length)return;
    walk(index+1,sum,picked);
    picked.push(usable[index]);walk(index+1,sum+num(usable[index].amount),picked);picked.pop();
  }
  walk(0,0,[]);
  return solutions.length===1?solutions[0]:null;
}
function noticeBundleForTransaction(t,student){
  const haystack=slug(`${t.reportedPaymentCode||''} ${t.content||''} ${t.ref||''}`);
  for(const [key,bundle] of noticeBundleCache){
    if(!key||!haystack.includes(key))continue;
    if(bundle.studentCode&&slug(bundle.studentCode)!==slug(student.code))continue;
    if(num(bundle.amount)!==num(t.amount))continue;
    return bundle;
  }
  return null;
}
function buildStudentReportIndexes(students){
  const byNameClass=new Map(),nameGroups=new Map(),byPersonalId=new Map();
  students.forEach(student=>{
    const nameKey=slug(student.name||''),classKey=slug(student.className||'');
    if(nameKey&&classKey)byNameClass.set(slug(`${student.name} ${student.className}`),student);
    if(nameKey){
      if(!nameGroups.has(nameKey))nameGroups.set(nameKey,[]);
      nameGroups.get(nameKey).push(student);
    }
    const pid=String(student.personalId||'').replace(/\s+/g,'');
    if(pid)byPersonalId.set(pid,student);
  });
  const byUniqueName=new Map([...nameGroups.entries()].filter(([,arr])=>arr.length===1).map(([k,arr])=>[k,arr[0]]));
  return {byNameClass,byUniqueName,byPersonalId};
}
function studentFromBankIdentity(t,indexes){
  const reportName=String(t.reportCustomerName||t.studentName||'').trim();
  const reportClass=String(t.reportClass||'').trim();
  const personalId=String(t.reportPersonalId||'').replace(/\s+/g,'');
  if(personalId&&indexes.byPersonalId.has(personalId))return indexes.byPersonalId.get(personalId);

  if(reportName&&reportClass){
    const direct=indexes.byNameClass.get(slug(`${reportName} ${reportClass}`));
    if(direct)return direct;
  }
  if(reportName){
    // BIDV thường trả "TEN HOC SINH 6A/7B..." trong cột Tên khách hàng.
    const withClass=indexes.byNameClass.get(slug(reportName));
    if(withClass)return withClass;

    // Nếu tên báo cáo không kèm lớp, chỉ tự ghép khi họ tên là duy nhất trong toàn trường.
    const unique=indexes.byUniqueName.get(slug(reportName));
    if(unique)return unique;
  }
  return null;
}
function reconcileTransactions(students, transactions) {
  const byCode=new Map(students.map(s=>[slug(s.code),s]));const byPaymentCode=new Map();
  students.forEach(student=>studentDueItems(student).forEach(item=>{if(item.paymentCode)byPaymentCode.set(slug(item.paymentCode),{student,item});if(item.qrAccountNumber)byPaymentCode.set(slug(item.qrAccountNumber),{student,item});}));
  const reportIndexes=buildStudentReportIndexes(students);
  const claimed=new Set();
  return [...transactions].sort((a,b)=>(a.date||a.importedAt||'').localeCompare(b.date||b.importedAt||'')).map(t=>{
    const insuranceResult=reconcileInsuranceTransaction(t,byCode,claimed);
    if(insuranceResult)return insuranceResult;
    const paymentReported=t.reportedPaymentCode||'';
    const studentReported=t.reportedStudentCode||t.studentCode||'';
    const alias=paymentReported?byPaymentCode.get(slug(paymentReported)):null;
    const identityStudent=studentFromBankIdentity(t,reportIndexes);
    const student=alias?.student||(studentReported?byCode.get(slug(studentReported)):null)||(paymentReported?byCode.get(slug(paymentReported)):null)||identityStudent;
    const reportedCode=studentReported||paymentReported||t.reportCustomerName||'';
    if(!isSuccessfulBankStatus(t.bankStatus))return {...t,studentCode:student?.code||'',studentName:student?.name||t.reportCustomerName||'',matched:false,paymentStatus:'bank_not_successful',feeCategory:alias?.item.category||transactionCategory(t)};
    if(!student)return {...t,studentCode:'',studentName:t.reportCustomerName||'',matched:false,paymentStatus:reportedCode?'unmatched':'missing_code'};
    const studentItems=studentDueItems(student), category=alias?.item.category||transactionCategory(t), amount=num(t.amount);
    if(t.sourceType==='cash'&&t.manualDueItemId){
      const item=studentItems.find(x=>x.id===t.manualDueItemId);
      if(!item)return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'no_due',feeCategory:category};
      if(item.amount!==amount)return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'amount_mismatch',feeCategory:item.category};
      const key=`${student.code}|${item.id}`;
      if(claimed.has(key))return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'duplicate',feeCategory:item.category};
      claimed.add(key);
      return {...t,studentCode:student.code,studentName:student.name,matched:true,paymentStatus:'valid',feeCategory:item.category,feeDetail:item.name,matchedDueItem:item.name,matchedDueItemId:item.id,matchedDueItemIds:[item.id],paymentChannel:'cash'};
    }
    if(!alias){
      const bundle=noticeBundleForTransaction(t,student);
      const bundleItems=bundle?(bundle.itemIds||[]).map(id=>studentItems.find(x=>x.id===id)).filter(Boolean):[];
      const available=studentItems.filter(item=>!item.legacy&&item.amount>0&&!claimed.has(`${student.code}|${item.id}`));
      const matchedBundle=bundleItems.length&&bundleItems.reduce((s,x)=>s+num(x.amount),0)===amount?bundleItems:findUniqueSubsetByAmount(available,amount);
      if(matchedBundle&&matchedBundle.length>1){
        const keys=matchedBundle.map(item=>`${student.code}|${item.id}`);
        if(keys.some(key=>claimed.has(key)))return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'duplicate',feeCategory:'other'};
        keys.forEach(key=>claimed.add(key));
        const cats=[...new Set(matchedBundle.map(x=>x.category))];
        return {...t,studentCode:student.code,studentName:student.name,matched:true,paymentStatus:'valid',feeCategory:cats.length===1?cats[0]:'other',feeDetail:matchedBundle.map(x=>x.name).join(' + '),matchedDueItem:matchedBundle.map(x=>x.name).join(' + '),matchedDueItemId:matchedBundle[0].id,matchedDueItemIds:matchedBundle.map(x=>x.id),paymentChannel:t.paymentChannel||'transfer',noticeBundleRemark:bundle?.remark||'',matchedBy:alias?'payment_code':(identityStudent?'name_class':'derived')};
      }
    }
    const candidates=alias?[alias.item]:category==='other'?[]:studentItems.filter(item=>!item.legacy&&item.category===category&&item.amount===amount);
    if(!candidates.length) {
      const categoryItems=studentItems.filter(item=>item.category===category);
      const status=categoryItems.some(item=>item.legacy)?'missing_category':categoryItems.length?'amount_mismatch':studentItems.length?'missing_category':'no_due';
      return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:status,feeCategory:category};
    }
    if(candidates.length===1&&candidates[0].amount!==amount)return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'amount_mismatch',feeCategory:category};
    const reportedDetail=slug(t.feeDetail||'');
    const byName=reportedDetail?candidates.filter(item=>slug(item.name)===reportedDetail):[];
    const possible=byName.length?byName:candidates;
    if(possible.length!==1)return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'ambiguous',feeCategory:category};
    if(possible[0].amount!==amount)return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'amount_mismatch',feeCategory:category};
    const item=possible[0], key=`${student.code}|${item.id}`;
    if(claimed.has(key))return {...t,studentCode:student.code,studentName:student.name,matched:false,paymentStatus:'duplicate',feeCategory:item.category};
    claimed.add(key);
    return {...t,studentCode:student.code,studentName:student.name,matched:true,paymentStatus:'valid',feeCategory:item.category,feeDetail:item.name,matchedDueItem:item.name,matchedDueItemId:item.id,matchedDueItemIds:[item.id],paymentChannel:t.paymentChannel||'transfer',matchedBy:alias?'payment_code':(studentReported&&byCode.get(slug(studentReported))?'student_code':(identityStudent?'name_class':'derived'))};
  });
}
function feeSummaries(students, transactions) {
  const summaries=Object.fromEntries(FEES.map(f=>[f.key,{...f,due:0,paid:0,remain:0,dueItems:0,paidItems:0,pct:0}]));
  const paidByStudent=new Map(),paidKeys=new Set(transactions.filter(t=>t.paymentStatus==='valid').flatMap(transactionPaidKeys));
  students.forEach(student=>{
    studentDueItems(student).forEach(item=>{
      const summary=summaries[item.category]||summaries.other,key=`${student.code}|${item.id}`;
      summary.due+=item.amount;summary.dueItems++;
      if(paidKeys.has(key)){summary.paid+=item.amount;summary.paidItems++;if(!paidByStudent.has(student.code))paidByStudent.set(student.code,{});const map=paidByStudent.get(student.code);map[item.category]=(map[item.category]||0)+item.amount;}
    });
  });
  Object.values(summaries).forEach(s=>{s.remain=Math.max(0,s.due-s.paid);s.pct=s.due?Math.min(100,Math.round(s.paid/s.due*100)):0;});
  return {summaries,paidByStudent};
}
function totals(students, transactions) {
  const {summaries,paidByStudent}=feeSummaries(students,transactions);
  const due=Object.values(summaries).reduce((sum,s)=>sum+s.due,0),paid=Object.values(summaries).reduce((sum,s)=>sum+s.paid,0);
  const dueItems=Object.values(summaries).reduce((sum,s)=>sum+s.dueItems,0),paidItems=Object.values(summaries).reduce((sum,s)=>sum+s.paidItems,0);
  return {due,paid,remain:Math.max(0,due-paid),dueItems,paidItems,unpaidItems:Math.max(0,dueItems-paidItems),paidByStudent,summaries,pct:due?Math.min(100,Math.round(paid/due*100)):0};
}
function renderStudents(students, transactions) {
  const reconciled=transactions.some(t=>t.paymentStatus)?transactions:reconcileTransactions(students,transactions);
  const paidByItem=new Set(reconciled.filter(t=>t.paymentStatus==='valid').flatMap(transactionPaidKeys));
  const query = slug($('#studentSearch')?.value || '');
  const filtered = students.filter(s => {
    const warningCount=studentWarnings(s).length;
    return (!showIncompleteOnly||warningCount>0)&&(!query || slug(`${s.code} ${s.classStudentCode||''} ${s.externalCode||''} ${s.name} ${s.className} ${s.phone||''} ${s.personalId||''} ${s.studentType||''} ${s.note||''}`).includes(query));
  });
  const incomplete=students.filter(s=>studentWarnings(s).length>0);
  $('#studentCountLabel').textContent = `${students.length.toLocaleString('vi-VN')} học sinh`;
  if($('#studentWarningPanel')){
    $('#studentWarningPanel').hidden=!incomplete.length;
    $('#studentWarningTitle').textContent=incomplete.length?`${incomplete.length} hồ sơ cần bổ sung / rà soát`:'';
    $('#studentWarningText').textContent=incomplete.length?'Các hồ sơ này vẫn được lưu và vẫn có thể cập nhật lại từ file sau. Web không chặn toàn bộ danh sách vì dữ liệu thiếu.':'';
    $('#showIncompleteStudents').textContent=showIncompleteOnly?'Xem tất cả học sinh':'Xem hồ sơ cần bổ sung';
  }
  $('#studentsTable').innerHTML = filtered.length ? filtered.map(s => {
    const items=studentDueItems(s);const due=items.reduce((sum,x)=>sum+x.amount,0);
    const paid=items.filter(item=>paidByItem.has(`${s.code}|${item.id}`)).reduce((sum,x)=>sum+x.amount,0);
    const sourceNote=[s.studentType,s.registrationInfo,s.note].filter(Boolean).join(' · ');
    const warnings=studentWarnings(s);const warn=warnings.length?`<span class="student-warning-badge" title="${escapeHTML(warnings.join(' · '))}">⚠ ${warnings.length}</span>`:'';
    return `<tr class="student-master-row ${warnings.length?'has-warning':''}" data-student-code="${escapeHTML(s.code)}"><td><strong>${escapeHTML(s.classStudentCode||s.code)}</strong> ${warn}</td><td><strong>${escapeHTML(s.name)}</strong></td><td>${escapeHTML(s.className || '—')}</td><td>${escapeHTML(s.gender||'—')}</td><td>${escapeHTML(s.birthDate||'—')}</td><td>${escapeHTML(sourceNote||'—')}</td><td>${items.length}</td><td>${money(due)}</td><td>${money(paid)}</td><td class="remain-cell"><strong>${money(Math.max(0,due-paid))}</strong></td><td><button class="text-button student-detail-button" data-student-code="${escapeHTML(s.code)}">Chi tiết ›</button></td></tr>`;
  }).join('') : `<tr><td colspan="11" class="empty-cell">${students.length ? 'Không tìm thấy học sinh phù hợp.' : 'Chưa có học sinh. Hãy tải file danh sách gốc của trường.'}</td></tr>`;
}
function itemMatchesFeeFilter(item,filter){
  if(filter==='all')return true;
  if(filter.startsWith('catalog:'))return item.catalogId===filter.slice(8);
  return item.category===filter;
}
async function populateFeeReportFilters(catalog){
  const select=$('#classFeeFilter');if(!select)return;const current=select.value||'all';
  const categoryOptions=[['all','Tất cả khoản thu'],['mandatory','Nhóm BHTT'],['insurance','Nhóm BHYT'],['service','Nhóm dịch vụ'],['other','Nhóm khác']];
  select.innerHTML=categoryOptions.map(([v,l])=>`<option value="${v}">${l}</option>`).join('')+(catalog.length?'<optgroup label="Từng khoản đã tạo">'+catalog.sort((a,b)=>a.name.localeCompare(b.name,'vi')).map(f=>`<option value="catalog:${f.id}">${escapeHTML(f.name)} · ${money(f.amount)}</option>`).join('')+'</optgroup>':'');
  if([...select.options].some(o=>o.value===current))select.value=current;
}
function renderClasses(students, transactions) {
  const filter=$('#classFeeFilter')?.value||'all';
  const groups = new Map();
  students.forEach(s => { const key=s.className||'Chưa xếp lớp'; const g=groups.get(key)||{students:[],count:0,due:0,paid:0,dueItems:0,paidItems:0};g.students.push(s);g.count++;groups.set(key,g); });
  // Render every class. Do not cap the dashboard list: a 30-row limit hid grade 9
  // whenever the school had more than 30 classes. numeric:true keeps 9A2 before 9A10.
  const rows = [...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0],'vi',{numeric:true,sensitivity:'base'}));
  $('#classTable').innerHTML = rows.length ? rows.map(([name,g])=>{
    const codes=new Set(g.students.map(s=>s.code));const classTransactions=transactions.filter(t=>codes.has(t.studentCode));
    let m;
    if(filter==='all') m=totals(g.students,classTransactions);
    else {
      const dueItems=g.students.flatMap(s=>studentDueItems(s).filter(item=>itemMatchesFeeFilter(item,filter)));
      const due=dueItems.reduce((sum,item)=>sum+item.amount,0);
      const dueKeys=new Set(g.students.flatMap(s=>studentDueItems(s).filter(item=>itemMatchesFeeFilter(item,filter)).map(item=>`${s.code}|${item.id}`)));
      const paidKeys=new Set(classTransactions.filter(t=>t.paymentStatus==='valid').flatMap(t=>transactionPaidKeys(t).filter(key=>dueKeys.has(key))));
      const paidItems=paidKeys.size;const paid=g.students.flatMap(s=>studentDueItems(s).map(item=>({s,item}))).filter(x=>itemMatchesFeeFilter(x.item,filter)&&paidKeys.has(`${x.s.code}|${x.item.id}`)).reduce((sum,x)=>sum+x.item.amount,0);
      m={dueItems:dueItems.length,paidItems,due,paid,remain:Math.max(0,due-paid)};
    }
    const pct=m.due?Math.min(100,Math.round(m.paid/m.due*100)):0;
    return `<tr class="class-summary-row" data-class-name="${escapeHTML(name)}" tabindex="0" role="button" aria-label="Xem chi tiết lớp ${escapeHTML(name)}"><td><strong class="class-link">${escapeHTML(name)}</strong></td><td>${g.count}</td><td>${m.dueItems}</td><td>${m.paidItems}</td><td>${money(m.remain)}</td><td><div class="class-progress"><span>${pct}%</span><span class="tiny-track"><i style="width:${pct}%"></i></span></div></td></tr>`;
  }).join('') : '<tr><td colspan="6" class="empty-cell">Chưa có dữ liệu. Nhập danh sách học sinh để bắt đầu.</td></tr>';
}
function renderClassDetail(className, students, transactions) {
  const classStudents=students.filter(s=>(s.className||'Chưa xếp lớp')===className);
  const codes=new Set(classStudents.map(s=>s.code));
  const classTransactions=transactions.filter(t=>codes.has(t.studentCode));
  const valid=classTransactions.filter(t=>t.paymentStatus==='valid');
  const paidItems=new Set(valid.flatMap(transactionPaidKeys));
  const feeGroups=new Map();
  classStudents.forEach(student=>studentDueItems(student).forEach(item=>{
    const key=item.name||getFeeLabel(item.category);
    const g=feeGroups.get(key)||{name:key,dueItems:0,paidItems:0,due:0,paid:0};
    g.dueItems++;g.due+=item.amount;
    if(paidItems.has(`${student.code}|${item.id}`)){g.paidItems++;g.paid+=item.amount;}
    feeGroups.set(key,g);
  }));
  const m=totals(classStudents,classTransactions);
  $('#classDetailTitle').textContent=`Lớp ${className}`;
  $('#classDetailSubtitle').textContent=`${classStudents.length} học sinh · ${m.paidItems}/${m.dueItems} món đã thu đủ · còn ${money(m.remain)}`;
  $('#classFeeSummary').innerHTML=[...feeGroups.values()].map(g=>{
    const pct=g.dueItems?Math.round(g.paidItems/g.dueItems*100):0;
    return `<div class="class-fee-chip"><strong>${escapeHTML(g.name)}</strong><span>${g.paidItems}/${g.dueItems} đã thu</span><b>${pct}%</b></div>`;
  }).join('')||'<div class="empty-inline">Lớp chưa có khoản phải thu.</div>';
  $('#classStudentTable').innerHTML=classStudents.map(student=>{
    const items=studentDueItems(student);
    const itemHtml=items.map(item=>{
      const paid=paidItems.has(`${student.code}|${item.id}`);
      return `<div class="student-fee-line ${paid?'paid':'unpaid'}"><span>${escapeHTML(item.name)}</span><strong>${money(item.amount)}</strong><em>${paid?'Đã thu':'Chưa thu'}</em></div>`;
    }).join('');
    return `<tr><td><strong>${escapeHTML(student.code)}</strong></td><td><strong>${escapeHTML(student.name)}</strong></td><td><div class="student-fees-mobile">${itemHtml}</div></td></tr>`;
  }).join('')||'<tr><td colspan="3" class="empty-cell">Không có học sinh trong lớp này.</td></tr>';
  $('#classDetailBackdrop').classList.add('open');
}
function closeClassDetail(){ $('#classDetailBackdrop').classList.remove('open'); }
async function openClassDetail(className){
  const [students,stored]=await Promise.all([all('students'),all('transactions')]);
  renderClassDetail(className,students,reconcileTransactions(students,stored));
}
function renderTransactions(transactions, limit) {
  const sorted=[...transactions].sort((a,b)=>(b.date||b.importedAt||'').localeCompare(a.date||a.importedAt||''));
  return (limit?sorted.slice(0,limit):sorted).map(t=>{
    const key=transactionCategory(t);const badgeClass=key==='service'?'service':key==='other'?'unknown':'';
    const statuses={valid:'Khớp món thu',unmatched:'Không tìm thấy mã khoản',missing_code:'Thiếu mã HS',amount_mismatch:'Sai số tiền món',missing_category:'Sai loại khoản',no_due:'Không có món phải thu',duplicate:'Trùng món thu',ambiguous:'Món tiền chưa phân biệt được',bank_not_successful:'Giao dịch không thành công'};
    return `<tr><td>${escapeHTML(t.date||'—')}</td><td>${escapeHTML(t.ref||t.id.slice(0,18))}</td><td><span class="category-badge ${badgeClass}">${getFeeLabel(key)}</span></td><td title="${escapeHTML(t.content)}">${escapeHTML((t.content||t.feeDetail||'—').slice(0,60))}</td><td>${escapeHTML(t.studentName||t.reportedStudentCode||'—')}</td><td><strong>${money(t.amount)}</strong></td><td><span class="status-badge ${t.paymentStatus==='valid'?'':'unmatched'}">${statuses[t.paymentStatus]||'Chưa đối soát'}</span></td></tr>`;
  }).join('');
}
function renderFeeProgress(summaries) {
  const visible=FEES.map(f=>summaries[f.key]).filter(s=>s.dueItems||s.paid);
  $('#feeProgressList').innerHTML=visible.length?visible.map((s,i)=>`<div class="fee-progress-row"><span class="fee-index">${String(i+1).padStart(2,'0')}</span><div class="fee-main"><strong>${s.label}</strong><small>${s.paidItems} / ${s.dueItems} món thu đủ</small></div><div class="fee-track"><i style="width:${s.pct}%"></i></div><div class="fee-progress-meta"><strong>${s.pct}%</strong><small>${money(s.paid)}</small></div></div>`).join(''):'<div class="empty-inline">Nhập danh sách học sinh và số phải thu theo từng khoản để xem tiến độ.</div>';
}
function renderChart(transactions) {
  const daily=new Map();
  transactions.filter(t=>t.paymentStatus==='valid').forEach(t=>{const date=t.date||String(t.importedAt||'').slice(0,10)||'Chưa rõ ngày';daily.set(date,(daily.get(date)||0)+num(t.amount));});
  const points=[...daily.entries()].sort((a,b)=>a[0].localeCompare(b[0]));
  $('#chartTotal').textContent=money(transactions.filter(t=>t.paymentStatus==='valid').reduce((sum,t)=>sum+num(t.amount),0)).replace(' ₫','');
  if(!points.length){$('#collectionChart').innerHTML='<div class="chart-empty">Chưa có giao dịch để lập biểu đồ.</div>';return;}
  let running=0;const values=points.map(([,amount])=>(running+=amount));const max=Math.max(...values,1);const width=480,height=165,pad={l:34,r:9,t:14,b:27};
  const coords=values.map((v,i)=>({x:pad.l+(points.length===1?(width-pad.l-pad.r)/2:i*(width-pad.l-pad.r)/(points.length-1)),y:pad.t+(1-v/max)*(height-pad.t-pad.b)}));
  const path=coords.map((p,i)=>`${i?'L':'M'}${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');const area=`${path} L${coords.at(-1).x},${height-pad.b} L${coords[0].x},${height-pad.b} Z`;
  const grid=[0,1,2,3].map(i=>{const y=pad.t+i*(height-pad.t-pad.b)/3;const label=money(max*(1-i/3)).replace(' ₫','');return `<line x1="${pad.l}" y1="${y}" x2="${width-pad.r}" y2="${y}" stroke="#e7eeeb" stroke-dasharray="3 4"/><text x="0" y="${y+3}" fill="#899994" font-size="9">${escapeHTML(label)}</text>`}).join('');
  const dates=points.length<5?points.map((p,i)=>i):[0,Math.floor((points.length-1)/3),Math.floor(2*(points.length-1)/3),points.length-1];
  const labels=[...new Set(dates)].map(i=>`<text x="${coords[i].x}" y="${height-5}" text-anchor="middle" fill="#899994" font-size="9">${escapeHTML(points[i][0].slice(5)||points[i][0])}</text>`).join('');
  $('#collectionChart').innerHTML=`<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Biểu đồ lũy kế số tiền đã thu"><defs><linearGradient id="chartFill" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#159987" stop-opacity=".2"/><stop offset="1" stop-color="#159987" stop-opacity="0"/></linearGradient></defs>${grid}<path d="${area}" fill="url(#chartFill)"/><path d="${path}" fill="none" stroke="#0c8b7a" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>${coords.map(p=>`<circle cx="${p.x}" cy="${p.y}" r="2.7" fill="#fff" stroke="#0c8b7a" stroke-width="2"/>`).join('')}${labels}</svg>`;
}
function renderFeeDetails(students, transactions, summaries) {
  const visible=['insurance','mandatory','service'].map(key=>summaries[key]).filter(s=>s.due||s.paid);
  $('#feeCategoryCards').innerHTML=visible.length?visible.map(s=>{
    const feeTx=transactions.filter(t=>transactionCategory(t)===s.key);const detailNames=[...new Set(feeTx.filter(t=>s.key==='service').map(t=>t.feeDetail||serviceDetail('',t.content)))].filter(Boolean).slice(0,6);
    const cls=s.key==='service'?'service-card':'';
    const description=s.key==='insurance'?'Bảo hiểm y tế theo mã khoản YT':s.key==='mandatory'?'Bảo hiểm thân thể tự nguyện theo mã khoản TT':'Gửi xe, nước uống và các dịch vụ khác';
    return `<article class="panel fee-detail-card ${cls}"><div class="fee-detail-title"><div><h2>${s.label}</h2><p>${description}</p></div><span class="category-badge ${s.key==='service'?'service':''}">${s.dueItems} món</span></div><div class="fee-total">${money(s.due)}</div><div class="fee-card-progress"><i style="width:${s.pct}%"></i></div><div class="fee-card-foot">Đã ghi nhận ${money(s.paid)} · Còn ${money(s.remain)} · ${s.pct}% giá trị</div><div class="fee-breakdown"><div><span>Phải thu</span><strong>${s.dueItems} món</strong></div><div><span>Đã thu đủ</span><strong>${s.paidItems} món</strong></div><div><span>Giao dịch</span><strong>${feeTx.length}</strong></div></div>${detailNames.length?`<div class="service-breakdown">${detailNames.map(n=>`<span class="service-chip">${escapeHTML(n)}</span>`).join('')}</div>`:''}</article>`;
  }).join(''):'<div class="panel empty-state">Chưa có số liệu theo từng khoản thu. Nhập danh sách học sinh và báo cáo thu để bắt đầu.</div>';
  const serviceTx=transactions.filter(t=>transactionCategory(t)==='service');const breakdown=new Map();
  serviceTx.forEach(t=>{const key=t.feeDetail||serviceDetail('',t.content);const b=breakdown.get(key)||{count:0,codes:new Set(),amount:0};b.count++;if(t.studentCode||t.reportedStudentCode)b.codes.add(t.studentCode||t.reportedStudentCode);if(t.paymentStatus==='valid')b.amount+=num(t.amount);breakdown.set(key,b);});
  $('#serviceTxnCount').textContent=`${serviceTx.length} giao dịch`;
  $('#serviceBreakdownTable').innerHTML=breakdown.size?[...breakdown.entries()].sort((a,b)=>b[1].amount-a[1].amount).map(([name,b])=>`<tr><td><strong>${escapeHTML(name)}</strong></td><td>${b.count}</td><td>${b.codes.size}</td><td><strong>${money(b.amount)}</strong></td></tr>`).join(''):'<tr><td colspan="4" class="empty-cell">Chưa có giao dịch dịch vụ.</td></tr>';
}
function qrText(value,maxLength=25){return String(value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toUpperCase().replace(/[^A-Z0-9 _.-]/g,' ').replace(/\s+/g,' ').trim().slice(0,maxLength);}
function emvTag(id,value){const text=String(value);if(text.length>99)throw new Error(`Trường QR ${id} vượt độ dài cho phép.`);return id+String(text.length).padStart(2,'0')+text;}
function crc16ccitt(value){let crc=0xFFFF;for(let i=0;i<value.length;i++){crc^=value.charCodeAt(i)<<8;for(let bit=0;bit<8;bit++)crc=crc&0x8000?(crc<<1)^0x1021:crc<<1;crc&=0xFFFF;}return crc.toString(16).toUpperCase().padStart(4,'0');}
function buildVietQrPayload(config,amount,remark){
  const accountInfo=emvTag('00','A000000727')+emvTag('01',emvTag('00',config.bin)+emvTag('01',config.accountNumber))+emvTag('02','QRIBFTTA');
  const reference=emvTag('08',qrText(remark,25));
  let payload='000201010212'+emvTag('38',accountInfo)+'52040000'+'5303704'+emvTag('54',String(num(amount)))+'5802VN'+emvTag('59',qrText(config.accountName,25))+'6007DONGHA'+emvTag('62',reference)+'6304';
  payload+=crc16ccitt(payload);return payload;
}
function qrCandidates(students,transactions){
  const paid=new Set(transactions.filter(t=>t.paymentStatus==='valid').flatMap(transactionPaidKeys));
  return students.flatMap(student=>studentDueItems(student).filter(item=>item.category!=='other'&&!item.legacy&&item.amount>0).map(item=>({student,item,paid:paid.has(`${student.code}|${item.id}`)})));
}
function renderQrPage(students,transactions,config){
  if(config){$('#qrBankBin').value=config.bin||'';$('#qrAccountNumber').value=config.accountNumber||'';$('#qrAccountName').value=config.accountName||'';$('#qrConfigStatus').textContent='Đã lưu trên thiết bị này';}
  $('#downloadQrs').hidden=true;$('#qrSelectionCount').textContent='Chưa có QR được tạo';$('#qrPreviewGrid').innerHTML='<div class="panel qr-empty-state">Chọn điều kiện rồi bấm “Tạo QR”.</div>';
  const currentClass=$('#qrClassFilter').value||'all';const classes=[...new Set(students.map(s=>s.className).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi'));
  $('#qrClassFilter').innerHTML='<option value="all">Tất cả lớp</option>'+classes.map(name=>`<option value="${escapeHTML(name)}">${escapeHTML(name)}</option>`).join('');
  if(classes.includes(currentClass))$('#qrClassFilter').value=currentClass;
  const candidates=qrCandidates(students,transactions);$('#qrDueCount').textContent=`${candidates.filter(x=>!x.paid).length} món còn phải thu`;
  if(students.some(s=>!s.hasFeeBreakdown))$('#qrPreviewSummary').textContent='Có học sinh chưa có dữ liệu chi tiết từng khoản. Hãy nhập lại danh sách phải thu để tạo QR chính xác.';
}
async function saveQrConfig(){
  const config={key:'qrAccount',bin:$('#qrBankBin').value.trim(),accountNumber:$('#qrAccountNumber').value.trim(),accountName:qrText($('#qrAccountName').value,25)};
  if(!/^\d{6}$/.test(config.bin))return toast('Mã BIN cần có đúng 6 chữ số.',true);
  if(!/^\d{4,30}$/.test(config.accountNumber))return toast('Hãy nhập số tài khoản nhận gồm 4–30 chữ số.',true);
  if(!config.accountName)return toast('Hãy nhập tên chủ tài khoản.',true);
  await request('meta','put',config);$('#qrAccountName').value=config.accountName;$('#qrConfigStatus').textContent='Đã lưu trên thiết bị này';toast('Đã lưu tài khoản nhận cục bộ.');
}
function qrFilename(entry){return `${slug(entry.student.className||'lop')}_${slug(entry.student.name)}_${slug(entry.student.code)}_${slug(entry.item.name)}.png`;}
function renderQrCards(entries){
  if(!entries.length){$('#qrPreviewGrid').innerHTML='<div class="panel qr-empty-state">Không có món phù hợp với điều kiện đã chọn.</div>';return;}
  $('#qrPreviewGrid').innerHTML=entries.map((entry,index)=>`<article class="panel qr-result-card"><div class="qr-result-top"><span class="category-badge ${entry.item.category==='service'?'service':''}">${escapeHTML(entry.item.name)}</span><span class="qr-class-tag">${escapeHTML(entry.student.className||'Chưa xếp lớp')}</span></div><div class="qr-person"><strong>${escapeHTML(entry.student.name)}</strong><small>${escapeHTML(entry.student.code)}</small></div><img src="${entry.png}" alt="Mã QR ${escapeHTML(entry.student.code)} ${escapeHTML(entry.item.name)}"><div class="qr-amount">${money(entry.item.amount)}</div><div class="qr-result-footer"><span>${escapeHTML(entry.remark)}</span><a class="button button-outline button-small" href="${entry.png}" download="${escapeHTML(entry.filename)}">Lưu PNG</a></div></article>`).join('');
}
async function generateQrs(){
  const config=await request('meta','get','qrAccount');if(!config?.bin||!config?.accountNumber||!config?.accountName)return toast('Hãy lưu tài khoản nhận tiền của trường trước.',true);
  if(typeof QRCode==='undefined'||typeof JSZip==='undefined')return toast('Thiếu bộ tạo ảnh cục bộ. Tải lại trang sau khi kết nối mạng.',true);
  const [students,stored]=await Promise.all([all('students'),all('transactions')]);const transactions=reconcileTransactions(students,stored);
  const className=$('#qrClassFilter').value,fee=$('#qrFeeFilter').value,status=$('#qrStatusFilter').value;
  const entries=qrCandidates(students,transactions).filter(x=>(className==='all'||x.student.className===className)&&(fee==='all'||x.item.category===fee)&&(status==='all'||!x.paid));
  if(!entries.length)return toast('Không có món phải thu phù hợp để tạo QR.',true);
  const output=[];const unique=new Set();
  for(const entry of entries){
    const key=`${entry.student.code}|${entry.item.id}`;if(unique.has(key))continue;unique.add(key);
    const code=entry.item.category==='insurance'?'BHYT':entry.item.category==='mandatory'?'BHTT':qrText(entry.item.name,12).replace(/\s+/g,'');
    const remark=qrText(entry.item.paymentCode||`${entry.student.code} ${code}`,25);const payload=buildVietQrPayload(config,entry.item.amount,remark);
    const holder=document.createElement('div');new QRCode(holder,{text:payload,width:240,height:240,correctLevel:QRCode.CorrectLevel.M});
    const canvas=holder.querySelector('canvas');if(!canvas)throw new Error('Không tạo được ảnh QR trên trình duyệt này.');
    output.push({...entry,remark,filename:qrFilename(entry),png:canvas.toDataURL('image/png')});
  }
  renderQrCards(output);$('#qrSelectionCount').textContent=`${output.length} ảnh QR đã tạo`;
  $('#qrPreviewSummary').textContent=`${output.length} mã, mỗi mã gắn với đúng một học sinh và một món thu.`;
  const zip=new JSZip();for(const item of output)zip.folder(qrText(item.student.className||'Chua_xep_lop',30)||'Chua_xep_lop').file(item.filename,item.png.split(',')[1],{base64:true});
  $('#downloadQrs').onclick=async()=>{const blob=await zip.generateAsync({type:'blob'});download(`SchoolCollect_QR_${new Date().toISOString().slice(0,10)}.zip`,blob,'application/zip');};
  $('#downloadQrs').hidden=false;
}
function feeCatalogRecord(items=[]){return {key:'feeCatalog',items};}
async function getFeeCatalog(){const rec=await request('meta','get','feeCatalog');return Array.isArray(rec?.items)?rec.items:[];}
function feeSafeCode(value,max=24){return String(value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toUpperCase().replace(/[^A-Z0-9]/g,'').slice(0,max);}
function buildPaymentCode(prefix,studentCode,feeCode){
  const a=feeSafeCode(prefix,6),b=feeSafeCode(studentCode,14),d=feeSafeCode(feeCode,8);
  return `${b.startsWith(a)?b:a+b}${d}`.slice(0,25);
}
function selectedValues(select){return [...select.selectedOptions].map(o=>o.value);}
function gradeOf(className){const m=String(className||'').trim().match(/^(\d{1,2})/);return m?m[1]:'';}
function feeTargetStudents(students,scope,targets){
  if(scope==='all')return students;
  const set=new Set(targets);
  if(scope==='grade')return students.filter(s=>set.has(gradeOf(s.className)));
  if(scope==='class')return students.filter(s=>set.has(s.className));
  if(scope==='student')return students.filter(s=>set.has(s.code));
  return [];
}
function populateFeeTargets(students){
  const scope=$('#feeScope')?.value||'all',select=$('#feeTargets');if(!select)return;
  if(scope==='all'){select.innerHTML='<option value="all" selected>Toàn bộ học sinh hiện có</option>';select.disabled=true;$('#feeTargetSummary').textContent=`${students.length} học sinh`;return;}
  select.disabled=false;
  let options=[];
  if(scope==='grade')options=[...new Set(students.map(s=>gradeOf(s.className)).filter(Boolean))].sort((a,b)=>Number(a)-Number(b)).map(x=>[x,`Khối ${x}`]);
  if(scope==='class')options=[...new Set(students.map(s=>s.className).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true})).map(x=>[x,x]);
  if(scope==='student')options=[...students].sort((a,b)=>a.className.localeCompare(b.className,'vi',{numeric:true})||a.name.localeCompare(b.name,'vi')).map(s=>[s.code,`${s.className} · ${s.name} · ${s.code}`]);
  select.innerHTML=options.map(([v,l])=>`<option value="${escapeHTML(v)}">${escapeHTML(l)}</option>`).join('');
  $('#feeTargetSummary').textContent=options.length?'Chưa chọn đối tượng':'Chưa có dữ liệu phù hợp';
}
function defaultFeeShortCode(category,code=''){
  if(category==='insurance')return 'YT';
  if(category==='mandatory')return 'TT';
  if(category==='service')return 'DV';
  const raw=feeSafeCode(code,2);return raw||'KH';
}
function updateFeePreview(students){
  const prefix=$('#feePrefix')?.value||'HG2',code=$('#feeCode')?.value||'KHOAN',category=$('#feeCategory')?.value||'service',shortCode=feeSafeCode($('#feeShortCode')?.value||defaultFeeShortCode(category,code),2),sample=students[0];
  $('#feeCodePreview').textContent=`${feeSafeCode(prefix)||'HG2'} + Mã học sinh + ${shortCode||'XX'}`;
  $('#feeCodePreviewExample').textContent=sample?`Ví dụ: ${buildPaymentCode(prefix,sample.classStudentCode||sample.code,shortCode)} · ${sample.name}`:'Ví dụ sẽ hiển thị sau khi có danh sách học sinh.';
}
function resetFeeForm(students=[]){
  $('#feeEditingId').value='';$('#feeBuilderTitle').textContent='Tạo khoản thu';$('#feeName').value='';$('#feeCode').value='';$('#feeShortCode').value='DV';$('#feeAmount').value='';$('#feeCategory').value='service';$('#feePrefix').value='HG2';$('#feeScope').value='all';$('#cancelFeeEdit').hidden=true;$('#saveFeeAssignment').textContent='Tạo & phân giao';populateFeeTargets(students);updateFeePreview(students);
}
function feeItemAssignmentCount(students,id){return students.reduce((n,s)=>n+studentDueItems(s).filter(x=>x.catalogId===id).length,0);}
function renderFeeCatalog(catalog,students){
  $('#feeCatalogCount').textContent=`${catalog.length} khoản`;
  $('#feeCatalogList').innerHTML=catalog.length?catalog.map(f=>{
    const count=feeItemAssignmentCount(students,f.id);
    const scope=f.lastScope==='all'?'Toàn trường':f.lastScope==='grade'?'Theo khối':f.lastScope==='class'?'Theo lớp':'Theo học sinh';
    return `<article class="fee-catalog-item"><div><span class="category-badge ${f.category==='service'?'service':''}">${escapeHTML(f.code)} / ${escapeHTML(f.shortCode||defaultFeeShortCode(f.category,f.code))}</span><h3>${escapeHTML(f.name)}</h3><p>${money(f.amount)} · ${scope} · ${count} học sinh</p></div><div class="fee-catalog-actions"><button class="button button-outline button-small" data-fee-edit="${f.id}">Sửa</button><button class="button button-danger button-small" data-fee-delete="${f.id}">Xóa</button></div></article>`;
  }).join(''):'<div class="empty-inline">Chưa có khoản thu. Tạo khoản đầu tiên ở biểu mẫu bên trái.</div>';
}
function bidvStudentParts(student){
  const classToken=feeSafeCode(student.className,8);
  const classCode=String(student.classStudentCode||'').toUpperCase();
  const m=classCode.match(/^HG2-([A-Z0-9]+)-(\d{2})$/);
  const seq=m?.[2]||'00';
  return {classToken:m?.[1]||classToken,seq};
}
function bidvCustomerCodeFromTemplate(template,student,feeCode,prefix,period){
  const {classToken,seq}=bidvStudentParts(student);
  const yearMatch=String(period||'').match(/(20\d{2})/);
  const yyyy=yearMatch?.[1]||String(new Date().getFullYear());
  const yy=yyyy.slice(-2);
  const vars={
    PREFIX:feeSafeCode(prefix,8),YY:feeSafeCode(yy,2),YYYY:feeSafeCode(yyyy,4),
    CLASS:classToken,SEQ:seq,HS:feeSafeCode(student.classStudentCode||student.code,16),
    BASE:feeSafeCode(student.code,16),CD:feeSafeCode(student.externalCode,20),FEE:feeSafeCode(feeCode,10)
  };
  let out=String(template||'{PREFIX}{YY}{CLASS}{SEQ}{FEE}').replace(/\{(PREFIX|YY|YYYY|CLASS|SEQ|HS|BASE|CD|FEE)\}/g,(_,k)=>vars[k]||'');
  out=feeSafeCode(out,40);
  return out;
}
function bidvName(value,max=120){return String(value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').replace(/[^A-Za-z0-9 /\\.,_-]/g,' ').replace(/\s+/g,' ').trim().slice(0,max);}
function bidvExportScopeStudents(students){
  const scope=$('#bidvExportScope')?.value||'all',target=$('#bidvExportTarget')?.value||'all';
  if(scope==='all'||target==='all')return students;
  if(scope==='grade')return students.filter(s=>gradeOf(s.className)===target);
  if(scope==='class')return students.filter(s=>s.className===target);
  return students;
}
function bidvFeeOptions(students,catalog){
  const out=catalog.map(f=>({...f,source:'catalog',optionId:`catalog:${f.id}`}));
  const seen=new Set(out.map(f=>slug(f.code)));
  const sourceDefs=[
    {category:'insurance',code:'BHYT',shortCode:'YT',name:'Bảo hiểm y tế (BHYT)'},
    {category:'mandatory',code:'BHTT',shortCode:'TT',name:'Bảo hiểm thân thể (BHTT)'},
    {category:'service',code:'DV',shortCode:'DV',name:'Dịch vụ khác'}
  ];
  for(const def of sourceDefs){
    const exists=students.some(s=>studentDueItems(s).some(x=>!x.catalogId&&x.category===def.category));
    if(exists&&!seen.has(slug(def.code)))out.push({...def,id:def.category,source:'source',optionId:`source:${def.category}`,amount:0});
  }
  return out;
}
function populateBidvExportControls(students,catalog){
  const fee=$('#bidvExportFee');if(!fee)return;
  const currentFee=fee.value,feeOptions=bidvFeeOptions(students,catalog);
  fee.innerHTML='<option value="">— Chọn khoản thu —</option>'+feeOptions.sort((a,b)=>a.name.localeCompare(b.name,'vi')).map(f=>`<option value="${escapeHTML(f.optionId)}">${escapeHTML(f.name)}${f.amount?` · ${money(f.amount)}`:''} · ${escapeHTML(f.code)}</option>`).join('');
  if([...fee.options].some(o=>o.value===currentFee))fee.value=currentFee;
  const scope=$('#bidvExportScope')?.value||'all',target=$('#bidvExportTarget'),currentTarget=target?.value||'all';
  if(!target)return;
  let options=[['all','Tất cả']];
  if(scope==='grade')options=[...new Set(students.map(s=>gradeOf(s.className)).filter(Boolean))].sort((a,b)=>Number(a)-Number(b)).map(x=>[x,`Khối ${x}`]);
  if(scope==='class')options=[...new Set(students.map(s=>s.className).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true})).map(x=>[x,x]);
  target.innerHTML=options.map(([v,l])=>`<option value="${escapeHTML(v)}">${escapeHTML(l)}</option>`).join('');
  if([...target.options].some(o=>o.value===currentTarget))target.value=currentTarget;
  updateBidvCustomerPreview(students,catalog);
}
function updateBidvCustomerPreview(students,catalog){
  const sample=bidvExportScopeStudents(students)[0]||students[0];
  const options=bidvFeeOptions(students,catalog),fee=options.find(f=>f.optionId===$('#bidvExportFee')?.value)||options[0];
  const prefix=$('#bidvCustomerPrefix')?.value||'HG2',period=$('#bidvBillPeriod')?.value||String(new Date().getFullYear()),template=$('#bidvCustomerTemplate')?.value||'{PREFIX}{YY}{CLASS}{SEQ}{FEE}';
  const feeToken=fee?.shortCode||defaultFeeShortCode(fee?.category,fee?.code);
  const code=sample&&fee?bidvCustomerCodeFromTemplate(template,sample,feeToken,prefix,period):'—';
  if($('#bidvCustomerPreview'))$('#bidvCustomerPreview').textContent=code;
  if($('#bidvCustomerPreviewText'))$('#bidvCustomerPreviewText').textContent=sample&&fee?`${sample.name} · ${sample.className} · ${fee.name}`:'Chọn khoản thu để xem mã mẫu.';
}
async function bidvExportRows(validateOnly=false){
  const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);
  const feeOptions=bidvFeeOptions(students,catalog),fee=feeOptions.find(f=>f.optionId===$('#bidvExportFee')?.value);
  if(!fee)throw new Error('Hãy chọn khoản thu cần xuất.');
  const period=String($('#bidvBillPeriod')?.value||'').trim();
  if(!period)throw new Error('Hãy nhập Kỳ hóa đơn.');
  const template=String($('#bidvCustomerTemplate')?.value||'').trim();
  if(!template)throw new Error('Hãy nhập cấu trúc Mã khách hàng.');
  const prefix=String($('#bidvCustomerPrefix')?.value||'').trim();
  const onlyUnpaid=$('#bidvOnlyUnpaid')?.checked!==false;
  const scoped=bidvExportScopeStudents(students).slice().sort((a,b)=>String(a.className||'').localeCompare(String(b.className||''),'vi',{numeric:true,sensitivity:'base'})||String(a.name||'').localeCompare(String(b.name||''),'vi',{sensitivity:'base'}));
  const transactions=reconcileTransactions(students,stored);
  const paidKeys=new Set(transactions.filter(t=>t.paymentStatus==='valid').flatMap(transactionPaidKeys));
  const rows=[],seen=new Set(),changes=new Map();
  for(const student of scoped){
    const item=studentDueItems(student).find(x=>fee.source==='catalog'?x.catalogId===fee.id:(!x.catalogId&&x.category===fee.category));
    if(!item)continue;
    const paymentKey=`${student.code}|${item.id}`;
    if(onlyUnpaid&&paidKeys.has(paymentKey))continue;
    const feeToken=fee.shortCode||defaultFeeShortCode(fee.category,fee.code);
    const customerId=bidvCustomerCodeFromTemplate(template,student,feeToken,prefix,period);
    if(!customerId)throw new Error(`Không tạo được Mã khách hàng cho ${student.name}.`);
    if(!/^[A-Z0-9]+$/.test(customerId))throw new Error(`Mã khách hàng ${customerId} có ký tự không hợp lệ.`);
    if(seen.has(customerId))throw new Error(`Trùng Mã khách hàng: ${customerId}. Hãy đổi cấu trúc mã.`);
    seen.add(customerId);
    if(item.paymentCode&&slug(item.paymentCode)!==slug(customerId)&&paidKeys.has(paymentKey))throw new Error(`Khoản ${fee.name} của ${student.name} đã thu với mã ${item.paymentCode}; không thể đổi mã.`);
    const vaName=bidvName(`${student.name} ${student.className}`,120);
    rows.push([
      rows.length+1,customerId,bidvName(student.name,120),vaName,period,num(item.amount),'VND','',
      '',bidvName(student.className,120),'',
      bidvName(student.externalCode||'',120),bidvName(student.code||'',120),
      bidvName(student.classStudentCode||'',120),bidvName(feeToken||'',120),bidvName(student.note||student.studentType||'',120)
    ]);
    if(slug(item.paymentCode)!==slug(customerId))changes.set(student.code,customerId);
  }
  if(!rows.length)throw new Error('Không có món thu phù hợp để xuất.');
  if(validateOnly)return {rows,fee,changes};
  if(changes.size){
    const updated=students.map(s=>{
      if(!changes.has(s.code))return s;
      const dueItems=studentDueItems(s).map(item=>(fee.source==='catalog'?item.catalogId===fee.id:(!item.catalogId&&item.category===fee.category))?{...item,paymentCode:changes.get(s.code),feeCode:item.feeCode||fee.code,shortCode:item.shortCode||fee.shortCode||defaultFeeShortCode(fee.category,fee.code)}:item);
      return {...s,dueItems,updatedAt:new Date().toISOString()};
    });
    await putMany('students',updated);
  }
  return {rows,fee,changes};
}
async function validateBidvExport(){
  const out=await bidvExportRows(true);
  $('#bidvExportStatus').textContent=`Hợp lệ: ${out.rows.length} dòng · ${new Set(out.rows.map(r=>r[1])).size} Mã khách hàng duy nhất · không trùng mã.`;
  toast(`Đã kiểm tra ${out.rows.length} dòng bảng kê BIDV.`);
}
async function exportBidvXlsx(){
  if(typeof JSZip==='undefined')throw new Error('Thiếu thư viện xuất Excel. Hãy tải lại trang.');
  const out=await bidvExportRows(false),rows=out.rows;
  const headers=['STT\nNo.','Mã khách hàng * \nCustomer ID *','Tên khách hàng *\nCustomer name*','Tên tài khoản định danh *\nVirtual account name*','Kỳ hóa đơn *\nBill Period*','Số tiền hóa đơn *\nBill amount','Loại tiền *\nCurrency','Mã hóa đơn \nBill ID ','Số điện thoại\nPhone number','Địa chỉ\nAddress','Email\nEmail','Thông tin bổ sung 1\nAdditional Information 1','Thông tin bổ sung 2\nAdditional Information 2','Thông tin bổ sung 3\nAdditional Information 3','Thông tin bổ sung 4\nAdditional Information 4','Thông tin bổ sung 5\nAdditional Information 5'];
  const zip=new JSZip();
  zip.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.folder('_rels').file('.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.folder('xl').file('workbook.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="MAU THEM MOI LOAI 2" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.folder('xl').folder('_rels').file('workbook.xml.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.folder('xl').file('styles.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1"/></cellXfs></styleSheet>');
  const xmlRows=[excelRow(1,headers.map(v=>({v,s:1})),44),...rows.map((r,i)=>excelRow(i+2,r.map(v=>({v,t:typeof v==='number'?'n':'s',s:2})),22))];
  const sheet=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="8" customWidth="1"/><col min="2" max="2" width="24" customWidth="1"/><col min="3" max="4" width="28" customWidth="1"/><col min="5" max="7" width="16" customWidth="1"/><col min="8" max="16" width="22" customWidth="1"/></cols><sheetData>${xmlRows.join('')}</sheetData><autoFilter ref="A1:P${rows.length+1}"/></worksheet>`;
  zip.folder('xl').folder('worksheets').file('sheet1.xml',sheet);
  const blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const safe=feeSafeCode(out.fee.code||out.fee.name,20),url=URL.createObjectURL(blob),link=document.createElement('a');
  link.href=url;link.download=`BIDV_Bang_ke_thu_ho_${safe}_${new Date().toISOString().slice(0,10)}.xlsx`;link.click();setTimeout(()=>URL.revokeObjectURL(url),1200);
  $('#bidvExportStatus').textContent=`Đã xuất ${rows.length} dòng · mã KH đã được lưu vào từng món thu để đối soát về sau.`;
  await refresh();toast(`Đã xuất ${rows.length} dòng bảng kê BIDV.`);
}
function qrFileText(value,max=120){
  return String(value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toUpperCase().replace(/[^A-Z0-9 ]/g,' ').replace(/\s+/g,' ').trim().slice(0,max);
}
function qrFileDate(value){
  const s=String(value||'').trim();if(!s)return '';
  const m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);if(m)return `${m[3]}/${m[2]}/${m[1]}`;
  return s;
}
function qrFileStudentCode(student,mode){
  if(mode==='class')return feeSafeCode(student.classStudentCode||student.code,24);
  return feeSafeCode(student.code,24);
}
function qrFileAllEntries(students,catalog,transactions){
  const paid=new Set(transactions.filter(t=>t.paymentStatus==='valid').flatMap(transactionPaidKeys));
  return students.flatMap(student=>studentDueItems(student).filter(item=>num(item.amount)>0).map(item=>({
    student,item,
    paid:paid.has(`${student.code}|${item.id}`),
    feeKey:item.catalogId?`catalog:${item.catalogId}`:`source:${item.category}`,
    feeName:item.name||getFeeLabel(item.category),
    feeShort:item.shortCode||catalog.find(f=>f.id===item.catalogId)?.shortCode||defaultFeeShortCode(item.category,item.feeCode||item.category)
  })));
}
function populateQrFileControls(students,catalog,transactions=[]){
  const scopeEl=$('#qrFileScope'),target=$('#qrFileTarget');if(!scopeEl||!target)return;
  const scope=scopeEl.value||'all',current=target.value||'all';
  const entries=qrFileAllEntries(students,catalog,transactions);
  let options=[['all','Tất cả']];
  if(scope==='class'){
    options=[...new Set(students.map(s=>s.className).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true})).map(x=>[x,x]);
  }else if(scope==='student'){
    options=[...students].sort((a,b)=>a.className.localeCompare(b.className,'vi',{numeric:true})||a.name.localeCompare(b.name,'vi')).map(s=>[s.code,`${s.className} · ${s.name} · ${s.classStudentCode||s.code}`]);
  }
  target.innerHTML=options.map(([v,l])=>`<option value="${escapeHTML(v)}">${escapeHTML(l)}</option>`).join('');
  if([...target.options].some(o=>o.value===current))target.value=current;
  const feeEl=$('#qrFileFee'),currentFee=feeEl?.value||'all';
  if(feeEl){
    const feeMap=new Map([['all','Tất cả món thu']]);
    for(const e of entries)if(!feeMap.has(e.feeKey))feeMap.set(e.feeKey,`${e.feeName} · ${e.feeShort}`);
    feeEl.innerHTML=[...feeMap.entries()].map(([v,l])=>`<option value="${escapeHTML(v)}">${escapeHTML(l)}</option>`).join('');
    if([...feeEl.options].some(o=>o.value===currentFee))feeEl.value=currentFee;
  }
  updateQrFilePreview(students,catalog,transactions);
}
function qrFileFilteredEntries(students,catalog,transactions){
  const scope=$('#qrFileScope')?.value||'all',target=$('#qrFileTarget')?.value||'all',feeTarget=$('#qrFileFee')?.value||'all',onlyUnpaid=$('#qrFileOnlyUnpaid')?.checked!==false;
  let entries=qrFileAllEntries(students,catalog,transactions);
  if(onlyUnpaid)entries=entries.filter(e=>!e.paid);
  if(feeTarget!=='all')entries=entries.filter(e=>e.feeKey===feeTarget);
  if(scope==='class'&&target!=='all')entries=entries.filter(e=>e.student.className===target);
  if(scope==='student'&&target!=='all')entries=entries.filter(e=>e.student.code===target);
  return entries.sort((a,b)=>String(a.student.className||'').localeCompare(String(b.student.className||''),'vi',{numeric:true,sensitivity:'base'})||String(a.student.name||'').localeCompare(String(b.student.name||''),'vi',{sensitivity:'base'})||String(a.feeName||'').localeCompare(String(b.feeName||''),'vi'));
}
function qrFileAccountNumber(prefix,student,item,mode,feeShort){
  const p=feeSafeCode(prefix,30),studentToken=qrFileStudentCode(student,mode),fee=feeSafeCode(feeShort,2);
  return feeSafeCode(`${p}${studentToken}${fee}`,50);
}
async function updateQrFilePreview(students,catalog,transactions=[]){
  const entries=qrFileFilteredEntries(students,catalog,transactions),sample=entries[0];
  if(!sample){$('#qrFileAccountPreview').textContent='—';$('#qrFileRemarkPreview').textContent='Không có món thu phù hợp.';return;}
  const prefix=$('#qrFilePrefix')?.value||'',period=$('#qrFilePeriod')?.value||'',mode=$('#qrFileStudentCodeMode')?.value||'base';
  const account=qrFileAccountNumber(prefix,sample.student,sample.item,mode,sample.feeShort);
  const accountName=qrFileText(`${sample.student.name} ${sample.student.className}`,80);
  const remark=qrFileText(`${accountName} NOP ${sample.feeName} ${period}`,140);
  $('#qrFileAccountPreview').textContent=account;
  $('#qrFileRemarkPreview').textContent=`${accountName} · ${remark}`;
}
async function qrFileRows(validateOnly=false){
  const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);
  const transactions=reconcileTransactions(students,stored),entries=qrFileFilteredEntries(students,catalog,transactions);
  if(!entries.length)throw new Error('Không có món thu phù hợp để xuất QR.');
  const prefix=String($('#qrFilePrefix')?.value||'').trim(),period=String($('#qrFilePeriod')?.value||'').trim(),bank='BIDV',bankBin='970418',mode=$('#qrFileStudentCodeMode')?.value||'base';
  if(!prefix)throw new Error('Hãy nhập Mã đầu định danh.');
  if(!period)throw new Error('Hãy nhập Kỳ thu.');
  const seen=new Set(),rows=[],changes=new Map();
  for(const e of entries){
    const accountNumber=qrFileAccountNumber(prefix,e.student,e.item,mode,e.feeShort);
    if(!accountNumber)throw new Error(`Không tạo được AccountNumber cho ${e.student.name}.`);
    if(!/^[A-Z0-9]+$/.test(accountNumber))throw new Error(`AccountNumber không hợp lệ: ${accountNumber}`);
    if(seen.has(accountNumber))throw new Error(`Trùng AccountNumber: ${accountNumber}. Hãy đổi tiền tố hoặc cấu trúc mã.`);
    seen.add(accountNumber);
    const accountName=qrFileText(`${e.student.name} ${e.student.className}`,80);
    const remark=qrFileText(`${accountName} NOP ${e.feeName} ${period}`,140);
    if(!accountName||!remark)throw new Error(`Thiếu AccountName/Remark của ${e.student.name}.`);
    rows.push([
      accountNumber,bank,bankBin,num(e.item.amount),accountName,remark,
      e.student.className||'',e.student.externalCode||e.student.code||'',qrFileDate(e.student.birthDate),
      e.feeName||'',e.student.name||''
    ]);
    if(slug(e.item.qrAccountNumber)!==slug(accountNumber)){
      if(!changes.has(e.student.code))changes.set(e.student.code,new Map());
      changes.get(e.student.code).set(e.item.id,accountNumber);
    }
  }
  if(validateOnly)return {rows,changes};
  if(changes.size){
    const updated=students.map(s=>{
      const m=changes.get(s.code);if(!m)return s;
      const dueItems=studentDueItems(s).map(item=>m.has(item.id)?{...item,qrAccountNumber:m.get(item.id)}:item);
      return {...s,dueItems,updatedAt:new Date().toISOString()};
    });
    await putMany('students',updated);
  }
  return {rows,changes};
}
async function validateQrFileExport(){
  const out=await qrFileRows(true);
  $('#qrFileStatus').textContent=`Hợp lệ: ${out.rows.length} dòng · ${new Set(out.rows.map(r=>r[0])).size} AccountNumber duy nhất · đúng 11 cột mẫu.`;
  toast(`Đã kiểm tra ${out.rows.length} dòng file QR.`);
}
async function exportQrFileXlsx(){
  if(typeof JSZip==='undefined')throw new Error('Thiếu thư viện xuất Excel. Hãy tải lại trang.');
  const out=await qrFileRows(false),rows=out.rows;
  const headers=['AccountNumber','Bank','BankBin','Amount','AccountName','Remark','Class','StudentID','Ngày sinh','Khoản Thu','Họ tên'];
  const zip=new JSZip();
  zip.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.folder('_rels').file('.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.folder('xl').file('workbook.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/><sheet name="Sheet3" sheetId="3" r:id="rId3"/></sheets></workbook>');
  zip.folder('xl').folder('_rels').file('workbook.xml.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.folder('xl').file('styles.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="10"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1"/></cellXfs></styleSheet>');
  const xmlRows=[excelRow(1,headers.map(v=>({v,s:1})),22),...rows.map((r,i)=>excelRow(i+2,r.map((v,j)=>({v,t:(j===2||j===3)?'n':'s',s:2})),20))];
  const sheet1=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="3" width="14" customWidth="1"/><col min="4" max="4" width="14" customWidth="1"/><col min="5" max="6" width="40" customWidth="1"/><col min="7" max="7" width="12" customWidth="1"/><col min="8" max="8" width="22" customWidth="1"/><col min="9" max="9" width="14" customWidth="1"/><col min="10" max="11" width="28" customWidth="1"/></cols><sheetData>${xmlRows.join('')}</sheetData><autoFilter ref="A1:K${rows.length+1}"/></worksheet>`;
  const blankSheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t></t></is></c></row></sheetData></worksheet>';
  const sheets=zip.folder('xl').folder('worksheets');sheets.file('sheet1.xml',sheet1);sheets.file('sheet2.xml',blankSheet);sheets.file('sheet3.xml',blankSheet);
  const blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=`File_tao_QR_dinh_danh_${new Date().toISOString().slice(0,10)}.xlsx`;link.click();setTimeout(()=>URL.revokeObjectURL(url),1200);
  $('#qrFileStatus').textContent=`Đã xuất ${rows.length} dòng theo đúng 11 cột mẫu Test 6E.xlsx.`;
  await refresh();toast(`Đã xuất ${rows.length} dòng tạo QR.`);
}
async function renderFeeSetup(students){
  const [catalog,stored]=await Promise.all([getFeeCatalog(),all('transactions')]);const transactions=reconcileTransactions(students,stored);renderFeeCatalog(catalog,students);populateFeeTargets(students);updateFeePreview(students);populateBidvExportControls(students,catalog);populateQrFileControls(students,catalog,transactions);
}
async function saveFeeAssignment(){
  const students=await all('students');if(!students.length)return toast('Hãy nhập danh sách học sinh trước khi tạo khoản thu.',true);
  const id=$('#feeEditingId').value||crypto.randomUUID(),name=$('#feeName').value.trim(),code=feeSafeCode($('#feeCode').value,8),amount=parseAmount($('#feeAmount').value),category=$('#feeCategory').value,shortCode=feeSafeCode($('#feeShortCode').value||defaultFeeShortCode(category,code),2),prefix=feeSafeCode($('#feePrefix').value||'HG',6),scope=$('#feeScope').value,targets=selectedValues($('#feeTargets'));
  if(!name||!code||shortCode.length!==2||amount<=0)return toast('Hãy nhập tên khoản, mã khoản, ký hiệu 2 ký tự và số tiền hợp lệ.',true);
  if(scope!=='all'&&!targets.length)return toast('Hãy chọn ít nhất một đối tượng áp dụng.',true);
  const catalog=await getFeeCatalog();const duplicate=catalog.find(f=>f.id!==id&&feeSafeCode(f.code)===code);if(duplicate)return toast('Mã khoản đã tồn tại. Hãy dùng mã khác.',true);
  const existingFee=catalog.find(f=>f.id===id);
  if(existingFee){const stored=await all('transactions'),tx=reconcileTransactions(students,stored),assigned=new Set(students.flatMap(s=>studentDueItems(s).filter(x=>x.catalogId===id).map(x=>`${s.code}|${x.id}`)));if(tx.some(t=>t.paymentStatus==='valid'&&transactionPaidKeys(t).some(key=>assigned.has(key))))return toast('Khoản này đã có giao dịch thu. Để bảo toàn đối soát, không thể sửa hoặc phân giao lại; hãy tạo khoản mới.',true);}
  const chosen=feeTargetStudents(students,scope,targets);if(!chosen.length)return toast('Không có học sinh nào trong phạm vi đã chọn.',true);
  const codes=new Set();for(const s of students)for(const item of studentDueItems(s))if(item.paymentCode)codes.add(slug(item.paymentCode));
  const chosenSet=new Set(chosen.map(s=>s.code));const now=new Date().toISOString();
  const updated=students.map(s=>{
    const items=studentDueItems(s).filter(item=>item.catalogId!==id);
    if(chosenSet.has(s.code)){
      const paymentCode=buildPaymentCode(prefix,s.classStudentCode||s.code,shortCode),paymentKey=slug(paymentCode);
      if(!paymentCode||[...codes].some(existing=>existing===paymentKey&&!studentDueItems(s).some(x=>x.catalogId===id&&slug(x.paymentCode)===paymentKey)))throw new Error(`Mã khách hàng bị trùng: ${paymentCode}. Hãy đổi tiền tố hoặc mã khoản.`);
      items.push({id:`catalog:${id}:${slug(s.code)}`,catalogId:id,paymentCode,category,name,amount,feeCode:code,shortCode,createdAt:now});
    }
    const due=items.reduce((sum,x)=>sum+num(x.amount),0),dueByCategory=items.reduce((o,x)=>(o[x.category]=(o[x.category]||0)+num(x.amount),o),{insurance:0,mandatory:0,service:0,other:0});
    return {...s,due,dueItems:items,dueByCategory,hasFeeBreakdown:true,updatedAt:now};
  });
  await putMany('students',updated);
  const entry={id,name,code,shortCode,amount,category,prefix,lastScope:scope,lastTargets:scope==='all'?[]:targets,updatedAt:now,createdAt:catalog.find(f=>f.id===id)?.createdAt||now};
  const next=[...catalog.filter(f=>f.id!==id),entry];await request('meta','put',feeCatalogRecord(next));
  await refresh();resetFeeForm(await all('students'));toast(`Đã phân giao “${name}” cho ${chosen.length} học sinh.`);
}
async function editFee(id){
  const [catalog,students]=await Promise.all([getFeeCatalog(),all('students')]);const f=catalog.find(x=>x.id===id);if(!f)return;
  $('#feeEditingId').value=f.id;$('#feeBuilderTitle').textContent='Cập nhật khoản thu';$('#feeName').value=f.name;$('#feeCode').value=f.code;$('#feeShortCode').value=f.shortCode||defaultFeeShortCode(f.category,f.code);$('#feeAmount').value=f.amount;$('#feeCategory').value=f.category;$('#feePrefix').value=f.prefix||'HG';$('#feeScope').value=f.lastScope||'all';populateFeeTargets(students);
  const targetSet=new Set(f.lastTargets||[]);[...$('#feeTargets').options].forEach(o=>o.selected=targetSet.has(o.value));$('#cancelFeeEdit').hidden=false;$('#saveFeeAssignment').textContent='Lưu & phân giao lại';updateFeePreview(students);setPage('fee-setup');
}
async function deleteFee(id){
  const [catalog,students,stored]=await Promise.all([getFeeCatalog(),all('students'),all('transactions')]);const f=catalog.find(x=>x.id===id);if(!f)return;
  const transactions=reconcileTransactions(students,stored);const assignedIds=new Set(students.flatMap(s=>studentDueItems(s).filter(x=>x.catalogId===id).map(x=>`${s.code}|${x.id}`)));
  if(transactions.some(t=>t.paymentStatus==='valid'&&transactionPaidKeys(t).some(key=>assignedIds.has(key))))return toast('Khoản này đã có giao dịch thu hợp lệ nên không thể xóa.',true);
  if(!confirm(`Xóa khoản “${f.name}” và toàn bộ phân giao chưa thu?`))return;
  const updated=students.map(s=>{const items=studentDueItems(s).filter(x=>x.catalogId!==id);return {...s,dueItems:items,due:items.reduce((a,x)=>a+x.amount,0),dueByCategory:items.reduce((o,x)=>(o[x.category]=(o[x.category]||0)+x.amount,o),{insurance:0,mandatory:0,service:0,other:0})};});
  await putMany('students',updated);await request('meta','put',feeCatalogRecord(catalog.filter(x=>x.id!==id)));await refresh();toast('Đã xóa khoản thu chưa phát sinh thanh toán.');
}
async function openStudentProfile(code){
  const [students,stored]=await Promise.all([all('students'),all('transactions')]);const s=students.find(x=>x.code===code);if(!s)return;const tx=reconcileTransactions(students,stored);const paid=new Set(tx.filter(t=>t.paymentStatus==='valid'&&t.studentCode===code).flatMap(transactionMatchedItemIds));
  $('#studentProfileTitle').textContent=`${s.name} · ${s.className}`;$('#studentProfileSubtitle').textContent=`Mã theo lớp: ${s.classStudentCode||'—'} · Mã gốc: ${s.code}`;
  const info=[['Mã số CĐ / mã ngoài',s.externalCode],['Giới tính',s.gender],['Ngày sinh',s.birthDate],['Phân loại HS',s.studentType],['Đăng ký khoản khác',s.registrationInfo],['Ghi chú',s.note]].filter(x=>x[1]);
  $('#studentProfileInfo').innerHTML=info.map(([k,v])=>`<div><span>${k}</span><strong>${escapeHTML(v)}</strong></div>`).join('')||'<div class="empty-inline">Chưa có thông tin hồ sơ bổ sung.</div>';
  const items=studentDueItems(s);$('#studentProfileFees').innerHTML=items.length?items.map(item=>`<div class="student-profile-fee ${paid.has(item.id)?'paid':''}"><div><strong>${escapeHTML(item.name)}</strong><small>${escapeHTML(item.paymentCode||'Chưa có mã thanh toán')}</small></div><b>${money(item.amount)}</b><span>${paid.has(item.id)?'Đã thu':'Chưa thu'}</span></div>`).join(''):'<div class="empty-inline">Học sinh chưa được phân giao khoản thu.</div>';
  $('#studentProfileBackdrop').classList.add('open');
}
function closeStudentProfile(){$('#studentProfileBackdrop').classList.remove('open');}

function unpaidCashCandidates(students,transactions){
  const paid=new Set(transactions.filter(t=>t.paymentStatus==='valid').flatMap(transactionPaidKeys));
  return students.flatMap(student=>studentDueItems(student).filter(item=>item.amount>0&&!paid.has(`${student.code}|${item.id}`)).map(item=>({student,item})));
}
function renderCashEntry(students,transactions){
  const candidates=unpaidCashCandidates(students,transactions),classEl=$('#cashClass'),studentEl=$('#cashStudent'),feeEl=$('#cashFee');if(!classEl)return;
  const currentClass=classEl.value,currentStudent=studentEl.value,currentFee=feeEl.value;
  const classes=[...new Set(candidates.map(x=>x.student.className).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true}));
  classEl.innerHTML='<option value="">— Chọn lớp —</option>'+classes.map(v=>`<option value="${escapeHTML(v)}">${escapeHTML(v)}</option>`).join('');
  if(classes.includes(currentClass))classEl.value=currentClass;
  const studentsInClass=[...new Map(candidates.filter(x=>!classEl.value||x.student.className===classEl.value).map(x=>[x.student.code,x.student])).values()].sort((a,b)=>a.name.localeCompare(b.name,'vi'));
  studentEl.innerHTML='<option value="">— Chọn học sinh —</option>'+studentsInClass.map(s=>`<option value="${escapeHTML(s.code)}">${escapeHTML(s.name)} · ${escapeHTML(s.code)}</option>`).join('');
  if(studentsInClass.some(s=>s.code===currentStudent))studentEl.value=currentStudent;
  const feeCandidates=candidates.filter(x=>(!classEl.value||x.student.className===classEl.value)&&(!studentEl.value||x.student.code===studentEl.value));
  feeEl.innerHTML='<option value="">— Chọn khoản thu —</option>'+feeCandidates.map(x=>`<option value="${escapeHTML(x.item.id)}" data-student-code="${escapeHTML(x.student.code)}">${escapeHTML(x.item.name)} · ${money(x.item.amount)}</option>`).join('');
  if([...feeEl.options].some(o=>o.value===currentFee))feeEl.value=currentFee;
  if(!$('#cashDate').value)$('#cashDate').value=new Date().toISOString().slice(0,10);
  updateCashPreview(students,transactions);
}
function updateCashPreview(students,transactions){
  const code=$('#cashStudent')?.value||'',itemId=$('#cashFee')?.value||'',student=students.find(s=>s.code===code),item=student?studentDueItems(student).find(x=>x.id===itemId):null;
  if(!student||!item){$('#cashAmount').value='';$('#cashEntryPreview').textContent='Chọn học sinh và khoản thu để ghi nhận tiền mặt.';return;}
  $('#cashAmount').value=money(item.amount);
  $('#cashEntryPreview').innerHTML=`<strong>${escapeHTML(student.name)} · ${escapeHTML(student.className||'')}</strong><span>${escapeHTML(item.name)} · ${money(item.amount)}</span><small>Mã thanh toán: ${escapeHTML(item.paymentCode||'—')}</small>`;
}
async function refreshCashEntry(){
  const [students,stored]=await Promise.all([all('students'),all('transactions')]);renderCashEntry(students,reconcileTransactions(students,stored));
}
async function recordCashPayment(printAfter=false){
  const [students,stored]=await Promise.all([all('students'),all('transactions')]);const txs=reconcileTransactions(students,stored);
  const student=students.find(s=>s.code===$('#cashStudent').value),item=studentDueItems(student||{}).find(x=>x.id===$('#cashFee').value);
  if(!student||!item)return toast('Hãy chọn đúng học sinh và khoản chưa thu.',true);
  if(txs.some(t=>t.paymentStatus==='valid'&&t.studentCode===student.code&&transactionMatchesItem(t,item.id)))return toast('Khoản này đã được ghi nhận đã thu.',true);
  const date=$('#cashDate').value||new Date().toISOString().slice(0,10),payer=$('#cashPayer').value.trim(),note=$('#cashNote').value.trim();
  if(!payer)return toast('Hãy nhập người nộp tiền.',true);
  const id=`cash:${crypto.randomUUID()}`,cashTx={id,ref:`TM-${date.replace(/-/g,'')}-${String(Date.now()).slice(-6)}`,date,content:note||`Thu tiền mặt ${item.name}`,amount:item.amount,feeCategory:item.category,feeDetail:item.name,reportedStudentCode:student.code,reportedPaymentCode:item.paymentCode||student.code,bankStatus:'thanh cong',studentCode:student.code,studentName:student.name,sourceFile:'Ghi nhận thủ công',importedAt:new Date().toISOString(),matched:true,sourceType:'cash',paymentChannel:'cash',manualDueItemId:item.id,payerName:payer};
  await request('transactions','put',cashTx);
  await request('history','put',{id:crypto.randomUUID(),kind:'Thu tiền mặt',fileName:cashTx.ref,rows:1,imported:1,detail:`${student.name} · ${item.name} · ${money(item.amount)}`,at:cashTx.importedAt});
  const reconciled=reconcileTransactions(students,[...stored,cashTx]),valid=reconciled.find(t=>t.id===id&&t.paymentStatus==='valid');
  if(!valid)throw new Error('Không thể xác nhận giao dịch tiền mặt với món thu đã chọn.');
  $('#cashPayer').value='';$('#cashNote').value='';
  if(printAfter)await issueReceiptCandidates([{paymentKey:receiptPaymentKey(student,item,valid),student,item,txn:valid}],{print:true});else{await refresh();toast('Đã ghi nhận khoản thu tiền mặt.');}
}

function noticeHash(value){
  let h=2166136261;for(const ch of String(value)){h^=ch.charCodeAt(0);h=Math.imul(h,16777619);}return (h>>>0).toString(36).toUpperCase().slice(0,5);
}
function noticeRemark(student,items){
  const code=feeSafeCode(student.code,10)||'HS';
  const seed=items.map(x=>x.id+'='+num(x.amount)).sort().join('|');
  return (`HG${code}TB${noticeHash(student.code+'|'+seed)}`).slice(0,25);
}
async function getNoticeBundles(){return (await request('meta','get','noticeBundles'))||{key:'noticeBundles',items:[]};}
async function saveNoticeBundles(entries){
  const old=await getNoticeBundles(),map=new Map((old.items||[]).map(x=>[slug(x.remark),x]));
  entries.forEach(x=>map.set(slug(x.remark),x));
  const items=[...map.values()].sort((a,b)=>(b.createdAt||'').localeCompare(a.createdAt||''));
  const record={key:'noticeBundles',items};await request('meta','put',record);setNoticeBundleCache(record);
}
function noticeSelectedFeeKeys(){
  const values=selectedValues($('#noticeFees'));return !values.length||values.includes('all')?null:new Set(values);
}
function noticeItemKey(item){return item.catalogId?`catalog:${item.catalogId}`:`name:${slug(item.name)}`;}
function noticeCandidates(students,transactions){
  const paid=new Set(transactions.filter(t=>t.paymentStatus==='valid').flatMap(transactionPaidKeys));
  const cls=$('#noticeClass')?.value||'all',studentCode=$('#noticeStudent')?.value||'all',status=$('#noticeDueStatus')?.value||'unpaid',feeKeys=noticeSelectedFeeKeys();
  return students.filter(s=>(cls==='all'||s.className===cls)&&(studentCode==='all'||s.code===studentCode)).map(student=>{
    const items=studentDueItems(student).filter(item=>!item.legacy&&item.amount>0).filter(item=>status==='all'||!paid.has(`${student.code}|${item.id}`)).filter(item=>!feeKeys||feeKeys.has(noticeItemKey(item)));
    return {student,items,total:items.reduce((sum,x)=>sum+num(x.amount),0)};
  }).filter(x=>x.items.length&&x.total>0);
}
function populateNoticeControls(students,catalog){
  const classEl=$('#noticeClass'),studentEl=$('#noticeStudent'),feeEl=$('#noticeFees');if(!classEl)return;
  const currentClass=classEl.value||'all',currentStudent=studentEl.value||'all',selected=new Set(selectedValues(feeEl));
  const classes=[...new Set(students.map(s=>s.className).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true}));
  classEl.innerHTML='<option value="all">Tất cả lớp</option>'+classes.map(v=>`<option value="${escapeHTML(v)}">${escapeHTML(v)}</option>`).join('');
  classEl.value=classes.includes(currentClass)?currentClass:'all';
  const scoped=students.filter(s=>classEl.value==='all'||s.className===classEl.value).sort((a,b)=>a.className.localeCompare(b.className,'vi',{numeric:true})||a.name.localeCompare(b.name,'vi'));
  studentEl.innerHTML='<option value="all">Tất cả học sinh trong phạm vi</option>'+scoped.map(s=>`<option value="${escapeHTML(s.code)}">${escapeHTML(s.className||'')} · ${escapeHTML(s.name)} · ${escapeHTML(s.code)}</option>`).join('');
  studentEl.value=scoped.some(s=>s.code===currentStudent)?currentStudent:'all';
  const unique=new Map();
  students.forEach(s=>studentDueItems(s).forEach(item=>{if(!item.legacy&&item.amount>0){const k=noticeItemKey(item);if(!unique.has(k))unique.set(k,item);}}));
  feeEl.innerHTML='<option value="all">Tất cả khoản phù hợp</option>'+[...unique.entries()].sort((a,b)=>a[1].name.localeCompare(b[1].name,'vi')).map(([k,item])=>`<option value="${escapeHTML(k)}">${escapeHTML(item.name)} · ${money(item.amount)}</option>`).join('');
  const kept=[...feeEl.options].filter(o=>selected.has(o.value));if(kept.length)kept.forEach(o=>o.selected=true);else feeEl.options[0].selected=true;
  if(!$('#noticeDeadline').value){const d=new Date();d.setDate(d.getDate()+10);$('#noticeDeadline').value=d.toISOString().slice(0,10);}
}
function updateNoticeSummary(students,transactions){
  const groups=noticeCandidates(students,transactions),items=groups.reduce((n,x)=>n+x.items.length,0),amount=groups.reduce((n,x)=>n+x.total,0);
  $('#noticeStudentCount').textContent=groups.length.toLocaleString('vi-VN');$('#noticeItemCount').textContent=items.toLocaleString('vi-VN');$('#noticeAmountTotal').textContent=money(amount);
  return groups;
}
function noticeDateVi(value){if(!value)return '—';const [y,m,d]=value.split('-');return `${d}/${m}/${y}`;}
function qrDataUrl(payload,size=360){
  if(typeof QRCode==='undefined')throw new Error('Thiếu thư viện tạo QR.');
  const holder=document.createElement('div');new QRCode(holder,{text:payload,width:size,height:size,correctLevel:QRCode.CorrectLevel.H});
  const canvas=holder.querySelector('canvas');if(!canvas)throw new Error('Không tạo được QR.');return canvas.toDataURL('image/png');
}
function noticeBrandHtml(qr){
  return `<div class="notice-branded-qr"><div class="bidv-flower">✿</div><div class="notice-qr-image"><img src="${qr}" alt="QR thanh toán"><span class="notice-qr-v">V</span></div><div class="notice-qr-brands"><b>napas<span>247</span></b><i></i><strong>BIDV</strong><em>✿</em></div><small>Quét mã để thanh toán</small></div>`;
}
async function buildNoticeEntries(){
  const [students,stored,config,receiptConfig]=await Promise.all([all('students'),all('transactions'),request('meta','get','qrAccount'),getReceiptConfig()]);
  if(!config?.bin||!config?.accountNumber||!config?.accountName)throw new Error('Hãy cấu hình tài khoản nhận tiền ở mục Tạo mã QR trước.');
  const transactions=reconcileTransactions(students,stored),groups=noticeCandidates(students,transactions);if(!groups.length)throw new Error('Không có học sinh/khoản thu phù hợp để tạo thông báo.');
  const deadline=$('#noticeDeadline').value||'',message=$('#noticeMessage').value.trim(),now=new Date().toISOString(),bundles=[],entries=[];
  for(const group of groups){
    const remark=noticeRemark(group.student,group.items),payload=buildVietQrPayload(config,group.total,remark),qr=qrDataUrl(payload,420);
    const bundle={remark,studentCode:group.student.code,itemIds:group.items.map(x=>x.id),amount:group.total,deadline,createdAt:now};
    bundles.push(bundle);entries.push({...group,remark,qr,deadline,message,school:{parentUnit:receiptConfig.parentUnit||'UBND XÃ HIẾU GIANG',schoolName:receiptConfig.schoolName||'TRƯỜNG THCS SỐ 2 HIẾU GIANG',schoolAddress:receiptConfig.schoolAddress||'',head:receiptConfig.head||''},qrConfig:config});
  }
  await saveNoticeBundles(bundles);return entries;
}
function noticeA4Html(entry){
  const rows=entry.items.map((item,i)=>`<tr><td>${i+1}</td><td>${escapeHTML(item.name)}</td><td>${money(item.amount)}</td></tr>`).join('');
  return `<article class="notice-a4-sheet">
    <div class="notice-a4-letterhead"><div><strong>${escapeHTML(entry.school.parentUnit)}</strong><b>${escapeHTML(entry.school.schoolName)}</b></div><div><strong>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</strong><b>Độc lập - Tự do - Hạnh phúc</b></div></div>
    <h1>THÔNG BÁO KHOẢN THU HỌC SINH</h1>
    <div class="notice-a4-student"><div><span>Học sinh</span><strong>${escapeHTML(entry.student.name)}</strong></div><div><span>Ngày sinh</span><strong>${escapeHTML(entry.student.birthDate||'—')}</strong></div><div><span>Lớp</span><strong>${escapeHTML(entry.student.className||'—')}</strong></div><div><span>Mã học sinh</span><strong>${escapeHTML(entry.student.classStudentCode||entry.student.code)}</strong></div></div>
    <h2>Chi tiết các khoản thu</h2><table><thead><tr><th>STT</th><th>Nội dung khoản thu</th><th>Số tiền</th></tr></thead><tbody>${rows}<tr class="notice-total-row"><td colspan="2">Tổng cộng</td><td>${money(entry.total)}</td></tr></tbody></table>
    <div class="notice-a4-words"><b>Bằng chữ:</b> ${escapeHTML(receiptAmountWords(entry.total))}.</div>
    <div class="notice-a4-payment"><div>${noticeBrandHtml(entry.qr)}</div><div class="notice-a4-guide"><h2>Hướng dẫn thanh toán</h2><p><b>1.</b> Quý phụ huynh quét QR để thanh toán.</p><p>QR đã có sẵn tổng tiền và mã tham chiếu <strong>${escapeHTML(entry.remark)}</strong>, không cần nhập lại nội dung.</p><p><b>2. Hạn nộp:</b> <strong class="notice-deadline">Trước ngày ${escapeHTML(noticeDateVi(entry.deadline))}</strong></p><p><b>3.</b> Sau khi thanh toán thành công, nhà trường cập nhật theo báo cáo thu BIDV.</p><p class="notice-parent-message">${escapeHTML(entry.message)}</p></div></div>
    <div class="notice-a4-sign"><div></div><div><i>Hiếu Giang, ngày ${new Intl.DateTimeFormat('vi-VN').format(new Date())}</i><strong>HIỆU TRƯỞNG</strong><span>${escapeHTML(entry.school.head||'')}</span></div></div>
  </article>`;
}
function renderNoticePreview(entries){
  const template=document.querySelector('input[name="noticeTemplate"]:checked')?.value||'a4';
  $('#noticePreviewSummary').textContent=`${entries.length} thông báo · ${entries.reduce((n,x)=>n+x.items.length,0)} món thu`;$('#noticeQrStatus').textContent=`${entries.length} QR đã tạo`;
  $('#noticePreviewList').innerHTML=entries.slice(0,6).map(e=>template==='a4'?noticeA4Html(e):`<article class="notice-mobile-preview"><div class="notice-mobile-head"><strong>${escapeHTML(e.school.schoolName)}</strong><span>Thông báo khoản thu học sinh</span></div><div class="notice-mobile-person"><b>${escapeHTML(e.student.name)}</b><span>${escapeHTML(e.student.className||'')} · ${escapeHTML(e.student.birthDate||'')}</span><small>Mã HS: ${escapeHTML(e.student.classStudentCode||e.student.code)}</small></div><div class="notice-mobile-total"><span>TỔNG PHẢI NỘP</span><strong>${money(e.total)}</strong><b>Hạn nộp: ${escapeHTML(noticeDateVi(e.deadline))}</b></div><div class="notice-mobile-items">${e.items.map((x,i)=>`<div><span>${i+1}. ${escapeHTML(x.name)}</span><strong>${money(x.amount)}</strong></div>`).join('')}</div>${noticeBrandHtml(e.qr)}<p>${escapeHTML(e.message)}</p></article>`).join('')+(entries.length>6?'<div class="notice-preview-more">… và '+(entries.length-6)+' thông báo khác sẽ được xuất.</div>':'');
}
function noticeCanvasRound(ctx,x,y,w,h,r,fill,stroke){
  const rr=Math.min(r,w/2,h/2);ctx.beginPath();ctx.roundRect(x,y,w,h,rr);if(fill){ctx.fillStyle=fill;ctx.fill();}if(stroke){ctx.strokeStyle=stroke;ctx.stroke();}
}
function canvasWrap(ctx,text,x,y,maxWidth,lineHeight,maxLines=10){
  const words=String(text).split(/\s+/);let line='',lines=[];
  for(const word of words){const test=line?line+' '+word:word;if(ctx.measureText(test).width>maxWidth&&line){lines.push(line);line=word;}else line=test;}
  if(line)lines.push(line);lines=lines.slice(0,maxLines);lines.forEach((l,i)=>ctx.fillText(l,x,y+i*lineHeight));return y+lines.length*lineHeight;
}
function drawFlower(ctx,x,y,s){
  ctx.save();ctx.fillStyle='#F7B928';for(let i=0;i<5;i++){const a=-Math.PI/2+i*Math.PI*2/5;ctx.beginPath();ctx.arc(x+Math.cos(a)*s*.34,y+Math.sin(a)*s*.34,s*.27,0,Math.PI*2);ctx.fill();}ctx.fillStyle='#fff';ctx.beginPath();ctx.arc(x,y,s*.22,0,Math.PI*2);ctx.fill();ctx.restore();
}
async function noticeMobilePng(entry){
  const W=1080,H=1920,canvas=document.createElement('canvas');canvas.width=W;canvas.height=H;const ctx=canvas.getContext('2d');
  ctx.fillStyle='#f8fcff';ctx.fillRect(0,0,W,H);ctx.textBaseline='top';
  const dark='#10335f',teal='#087f78',muted='#61738b',line='#d7eaf5';
  ctx.fillStyle='#e9f7ff';ctx.fillRect(0,0,W,190);ctx.fillStyle=dark;ctx.font='700 44px Arial';ctx.fillText(entry.school.schoolName,70,55);ctx.font='30px Arial';ctx.fillStyle=muted;ctx.fillText('Thông báo khoản thu học sinh',70,113);
  noticeCanvasRound(ctx,45,215,990,205,30,'#fff','#d9e9f2');ctx.fillStyle=dark;ctx.font='700 40px Arial';ctx.fillText(entry.student.name,105,260);ctx.font='28px Arial';ctx.fillStyle=muted;ctx.fillText(`Lớp ${entry.student.className||'—'}  •  ${entry.student.birthDate||'—'}`,105,320);ctx.fillText(`Mã HS: ${entry.student.classStudentCode||entry.student.code}`,105,360);
  noticeCanvasRound(ctx,45,450,990,230,34,'#e5f8fb');ctx.fillStyle=dark;ctx.font='700 34px Arial';ctx.fillText('TỔNG PHẢI NỘP',105,492);ctx.fillStyle=teal;ctx.font='700 72px Arial';ctx.fillText(money(entry.total),105,545);noticeCanvasRound(ctx,740,500,245,120,24,'#fff0ef');ctx.fillStyle='#cf4944';ctx.font='26px Arial';ctx.fillText('Hạn nộp:',775,525);ctx.font='700 30px Arial';ctx.fillText(noticeDateVi(entry.deadline),775,566);
  const itemY=715,itemH=Math.min(410,115+entry.items.length*65);noticeCanvasRound(ctx,45,itemY,990,itemH,30,'#fff','#d9e9f2');ctx.fillStyle=dark;ctx.font='700 34px Arial';ctx.fillText('Chi tiết khoản thu',85,itemY+30);let y=itemY+90;ctx.font='27px Arial';
  entry.items.slice(0,6).forEach((item,i)=>{ctx.fillStyle='#eaf6fc';ctx.beginPath();ctx.arc(92,y+17,22,0,Math.PI*2);ctx.fill();ctx.fillStyle=dark;ctx.font='700 22px Arial';ctx.textAlign='center';ctx.fillText(String(i+1),92,y+5);ctx.textAlign='left';ctx.font='27px Arial';ctx.fillText(item.name,135,y);ctx.font='700 27px Arial';ctx.textAlign='right';ctx.fillText(money(item.amount),970,y);ctx.textAlign='left';ctx.strokeStyle=line;ctx.beginPath();ctx.moveTo(80,y+52);ctx.lineTo(985,y+52);ctx.stroke();y+=65;});
  const qrY=itemY+itemH+30;noticeCanvasRound(ctx,45,qrY,990,580,30,'#eefcfb','#d9ecea');ctx.fillStyle=dark;ctx.font='700 34px Arial';ctx.textAlign='center';ctx.fillText('Quét mã để thanh toán',W/2,qrY+26);ctx.font='24px Arial';ctx.fillStyle=muted;ctx.fillText('Không cần nhập lại số tiền hoặc nội dung chuyển khoản',W/2,qrY+70);drawFlower(ctx,W/2,qrY+122,40);
  const qrImg=new Image();qrImg.src=entry.qr;await new Promise((res,rej)=>{qrImg.onload=res;qrImg.onerror=rej;});ctx.drawImage(qrImg,360,qrY+145,360,360);
  ctx.fillStyle='#fff';ctx.fillRect(518,qrY+303,44,44);ctx.fillStyle='#e72b3b';ctx.font='900 34px Arial';ctx.textAlign='center';ctx.fillText('V',540,qrY+308);
  ctx.font='italic 700 28px Arial';ctx.fillStyle='#245487';ctx.fillText('napas',425,qrY+520);ctx.fillStyle='#28a8df';ctx.fillText('247',510,qrY+520);ctx.fillStyle='#778698';ctx.fillRect(580,qrY+516,2,34);ctx.font='700 35px Arial';ctx.fillStyle=teal;ctx.fillText('BIDV',660,qrY+514);drawFlower(ctx,748,qrY+532,22);
  ctx.textAlign='left';noticeCanvasRound(ctx,45,H-150,990,95,28,'#e8f5ff');ctx.fillStyle='#245487';ctx.font='25px Arial';canvasWrap(ctx,entry.message,90,H-120,900,31,2);
  return canvas.toDataURL('image/png');
}
async function previewNotices(){
  const entries=await buildNoticeEntries();renderNoticePreview(entries);window.__noticeEntries=entries;return entries;
}
async function printNoticeA4(){
  const entries=await buildNoticeEntries(),root=$('#noticePrintRoot');root.innerHTML=entries.map(noticeA4Html).join('');document.body.classList.add('printing-notices');setTimeout(()=>{window.print();setTimeout(()=>document.body.classList.remove('printing-notices'),400);},80);
}
async function downloadNoticeImages(){
  if(typeof JSZip==='undefined')throw new Error('Thiếu thư viện ZIP.');
  const entries=await buildNoticeEntries(),zip=new JSZip();let done=0;$('#noticeQrStatus').textContent='Đang tạo ảnh…';
  for(const entry of entries){const png=await noticeMobilePng(entry),folder=zip.folder(qrText(entry.student.className||'Chua_xep_lop',30)||'Chua_xep_lop');folder.file(`${feeSafeCode(entry.student.className||'LOP',10)}_${feeSafeCode(entry.student.code,14)}_${qrText(entry.student.name,30).replace(/\s+/g,'_')}.png`,png.split(',')[1],{base64:true});done++;$('#noticeQrStatus').textContent=`${done}/${entries.length} ảnh`;}
  const blob=await zip.generateAsync({type:'blob'});download(`Thong_bao_khoan_thu_${new Date().toISOString().slice(0,10)}.zip`,blob,'application/zip');$('#noticeQrStatus').textContent=`${entries.length} ảnh đã xuất`;toast(`Đã xuất ${entries.length} ảnh thông báo.`);
}
async function renderNoticeTool(students,transactions,catalog){
  populateNoticeControls(students,catalog);updateNoticeSummary(students,transactions);
}

function receiptConfigDefaults(){
  return {key:'receiptConfig',parentUnit:'UBND XÃ HIẾU GIANG',schoolName:'TRƯỜNG THCS SỐ 2 HIẾU GIANG',schoolAddress:'',unitCode:'',transferPrefix:'XN',cashPrefix:'PT',preparer:'',cashier:'',accountant:'',head:''};
}
async function getReceiptConfig(){
  const stored=(await request('meta','get','receiptConfig'))||{},base=receiptConfigDefaults();
  const out={...base,...stored};
  if(!stored.transferPrefix&&stored.prefix)out.transferPrefix=stored.prefix;
  return out;
}
function receiptPaymentKey(student,item,txn){return `${student.code}|${item.id}|${txn.id}`;}
function receiptCandidates(students,transactions){
  const byStudent=new Map(students.map(s=>[s.code,s]));
  return transactions.filter(t=>t.paymentStatus==='valid'&&t.studentCode).flatMap(t=>{
    const student=byStudent.get(t.studentCode);if(!student)return [];
    let ids=transactionMatchedItemIds(t);
    if(!ids.length){
      const category=transactionCategory(t),amount=num(t.amount);
      const possible=studentDueItems(student).filter(item=>item.amount===amount&&(category==='other'||item.category===category));
      if(possible.length===1)ids=[possible[0].id];
    }
    return ids.map(id=>studentDueItems(student).find(x=>x.id===id)).filter(Boolean).map(item=>({paymentKey:receiptPaymentKey(student,item,t),student,item,txn:t}));
  });
}
function receiptStateFor(candidate,receipts){
  const related=receipts.filter(r=>r.paymentKey===candidate.paymentKey).sort((a,b)=>(b.issuedAt||'').localeCompare(a.issuedAt||''));
  const active=related.find(r=>r.status==='issued');
  if(active)return {status:'issued',receipt:active};
  if(related.length)return {status:'cancelled',receipt:related[0]};
  return {status:'ready',receipt:null};
}
function receiptStatusLabel(status){return status==='issued'?'Đã phát hành':status==='cancelled'?'Đã hủy':'Chưa phát hành';}
function receiptAmountWords(n){
  n=num(n);if(n===0)return 'Không đồng';
  const digit=['không','một','hai','ba','bốn','năm','sáu','bảy','tám','chín'];
  const read3=(v,full=false)=>{
    const h=Math.floor(v/100),t=Math.floor(v%100/10),u=v%10;const out=[];
    if(h||full){out.push(digit[h],'trăm');}
    if(t>1){out.push(digit[t],'mươi');if(u===1)out.push('mốt');else if(u===5)out.push('lăm');else if(u)out.push(digit[u]);}
    else if(t===1){out.push('mười');if(u===5)out.push('lăm');else if(u)out.push(digit[u]);}
    else if(u){if(h||full)out.push('lẻ');out.push(digit[u]);}
    return out.join(' ');
  };
  const units=['','nghìn','triệu','tỷ','nghìn tỷ','triệu tỷ'];const chunks=[];let x=n;
  while(x>0){chunks.push(x%1000);x=Math.floor(x/1000);}
  const parts=[];for(let i=chunks.length-1;i>=0;i--){if(!chunks[i])continue;parts.push(read3(chunks[i],i<chunks.length-1&&chunks[i]<100),units[i]);}
  const s=parts.join(' ').replace(/\s+/g,' ').trim();return s.charAt(0).toLocaleUpperCase('vi-VN')+s.slice(1)+' đồng';
}
async function saveReceiptConfig(){
  const config={key:'receiptConfig',parentUnit:$('#receiptParentUnit').value.trim(),schoolName:$('#receiptSchoolName').value.trim(),schoolAddress:$('#receiptSchoolAddress').value.trim(),unitCode:$('#receiptUnitCode').value.trim(),transferPrefix:feeSafeCode($('#receiptTransferPrefix').value||'XN',6)||'XN',cashPrefix:feeSafeCode($('#receiptCashPrefix').value||'PT',6)||'PT',preparer:$('#receiptPreparer').value.trim(),cashier:$('#receiptCashier').value.trim(),accountant:$('#receiptAccountant').value.trim(),head:$('#receiptHead').value.trim()};
  if(!config.schoolName)return toast('Hãy nhập tên đơn vị trên chứng từ.',true);
  await request('meta','put',config);$('#receiptTransferPrefix').value=config.transferPrefix;$('#receiptCashPrefix').value=config.cashPrefix;toast('Đã lưu thông tin chứng từ trên thiết bị này.');
}
function fillReceiptConfig(config){
  const map={receiptParentUnit:'parentUnit',receiptSchoolName:'schoolName',receiptSchoolAddress:'schoolAddress',receiptUnitCode:'unitCode',receiptTransferPrefix:'transferPrefix',receiptCashPrefix:'cashPrefix',receiptPreparer:'preparer',receiptCashier:'cashier',receiptAccountant:'accountant',receiptHead:'head'};
  Object.entries(map).forEach(([id,key])=>{const el=$('#'+id);if(el)el.value=config[key]||'';});
}
async function nextReceiptNumber(method,config){
  const year=new Date().getFullYear();
  const prefix=method==='cash'?(config.cashPrefix||'PT'):(config.transferPrefix||config.prefix||'XN');
  const key=`receiptSeq:${year}:${prefix}`,seq=await request('meta','get',key);const current=num(seq?.value)+1;
  await request('meta','put',{key,value:current});
  return {year,seq:current,number:`${prefix}-${year}-${String(current).padStart(6,'0')}`};
}
function receiptVerificationCode(number,id){return `RCT-${String(number).replace(/[^A-Z0-9]/gi,'').slice(-10).toUpperCase()}-${String(id).slice(0,6).toUpperCase()}`;}
async function issueReceiptCandidates(candidates,{print=true}={}){
  if(!candidates.length)return toast('Không có món phù hợp để phát hành.',true);
  const [receipts,config]=await Promise.all([all('receipts'),getReceiptConfig()]);
  const issuable=candidates.filter(x=>!receipts.some(r=>r.paymentKey===x.paymentKey&&r.status==='issued'));
  if(!issuable.length){if(print)printReceiptRecords(receipts.filter(r=>r.status==='issued'&&candidates.some(x=>x.paymentKey===r.paymentKey)),config);return toast('Các món đã chọn đều đã có chứng từ.');}
  const now=new Date().toISOString(),records=[];
  for(const x of issuable){
    const method=x.txn.paymentChannel==='cash'?'cash':'transfer',n=await nextReceiptNumber(method,config),id=crypto.randomUUID();
    records.push({id,paymentKey:x.paymentKey,number:n.number,sequence:n.seq,year:n.year,status:'issued',issuedAt:now,receiptType:method==='cash'?'cash_receipt':'payment_confirmation',paymentMethod:method==='cash'?'Tiền mặt':'Chuyển khoản/QR',verificationCode:receiptVerificationCode(n.number,id),studentCode:x.student.code,studentName:x.student.name,className:x.student.className||'',feeName:x.item.name,feeCode:x.item.feeCode||x.item.category||'',paymentCode:x.item.paymentCode||'',amount:x.item.amount,paymentDate:x.txn.date||'',transactionRef:x.txn.ref||x.txn.id,transactionId:x.txn.id,content:x.txn.content||'',payerName:x.txn.payerName||'',snapshot:{parentUnit:config.parentUnit,schoolName:config.schoolName,schoolAddress:config.schoolAddress,unitCode:config.unitCode,preparer:config.preparer,cashier:config.cashier,accountant:config.accountant,head:config.head}});
  }
  await putMany('receipts',records);await request('history','put',{id:crypto.randomUUID(),kind:'Phát hành chứng từ',fileName:records.length===1?records[0].number:`${records.length} chứng từ`,rows:records.length,imported:records.length,detail:`Đã cấp ${records.length} chứng từ thanh toán`,at:now});
  await refresh();toast(`Đã phát hành ${records.length} chứng từ.`);
  if(print)printReceiptRecords(records,config);
}
function receiptRecordHtml(r,config){
  const s={...config,...(r.snapshot||{})},cancelled=r.status==='cancelled';
  return `<article class="receipt-print-sheet ${cancelled?'receipt-print-cancelled':''}">
    <div class="receipt-print-top"><div><strong>${escapeHTML(s.parentUnit||'')}</strong><b>${escapeHTML(s.schoolName||'')}</b>${s.schoolAddress?`<span>${escapeHTML(s.schoolAddress)}</span>`:''}${s.unitCode?`<span>Mã đơn vị/MST: ${escapeHTML(s.unitCode)}</span>`:''}</div><div class="receipt-number"><span>Số chứng từ</span><strong>${escapeHTML(r.number)}</strong></div></div>
    <h1>${r.receiptType==='cash_receipt'?'PHIẾU THU':'PHIẾU XÁC NHẬN THANH TOÁN'}</h1>
    <div class="receipt-print-meta"><span>Ngày lập: ${escapeHTML(new Intl.DateTimeFormat('vi-VN').format(new Date(r.issuedAt)))}</span><span>Hình thức: ${escapeHTML(r.paymentMethod||'Chuyển khoản/QR')}</span></div>
    <div class="receipt-print-person"><p><span>Học sinh</span><strong>${escapeHTML(r.studentName)}</strong></p><p><span>Mã học sinh</span><strong>${escapeHTML(r.studentCode)}</strong></p><p><span>Lớp</span><strong>${escapeHTML(r.className||'—')}</strong></p></div>
    ${r.receiptType==='cash_receipt'?`<div class="receipt-payer-line"><span>Người nộp tiền:</span><strong>${escapeHTML(r.payerName||'—')}</strong><span>Lý do:</span><strong>${escapeHTML(r.content||r.feeName)}</strong></div>`:''}
    <table class="receipt-print-table"><thead><tr><th>Khoản thu</th><th>Mã thanh toán</th><th>Ngày thanh toán</th><th>Số tiền</th></tr></thead><tbody><tr><td>${escapeHTML(r.feeName)}</td><td>${escapeHTML(r.paymentCode||'—')}</td><td>${escapeHTML(r.paymentDate||'—')}</td><td>${money(r.amount)}</td></tr></tbody></table>
    <div class="receipt-print-amount"><span>Số tiền bằng chữ:</span><strong>${escapeHTML(receiptAmountWords(r.amount))}</strong></div>
    <div class="receipt-print-bank">${r.receiptType==='cash_receipt'?`<span>Số tham chiếu tiền mặt: <b>${escapeHTML(r.transactionRef||'—')}</b></span>`:`<span>Mã giao dịch ngân hàng: <b>${escapeHTML(r.transactionRef||'—')}</b></span>`}<span>Mã kiểm tra: <b>${escapeHTML(r.verificationCode||'—')}</b></span></div>
    ${cancelled?`<div class="receipt-cancel-stamp">ĐÃ HỦY · ${escapeHTML(r.cancelReason||'')}</div>`:''}
    ${r.receiptType==='cash_receipt'?`<div class="receipt-signatures cash-signatures"><div><strong>Người nộp tiền</strong><span>${escapeHTML(r.payerName||'')}</span></div><div><strong>Người lập</strong><span>${escapeHTML(s.preparer||'')}</span></div><div><strong>Thủ quỹ</strong><span>${escapeHTML(s.cashier||'')}</span></div><div><strong>Kế toán / Phụ trách kế toán</strong><span>${escapeHTML(s.accountant||'')}</span></div><div><strong>Thủ trưởng đơn vị</strong><span>${escapeHTML(s.head||'')}</span></div></div>`:`<div class="receipt-signatures"><div><strong>Người lập</strong><span>${escapeHTML(s.preparer||'')}</span></div><div><strong>Kế toán / Phụ trách kế toán</strong><span>${escapeHTML(s.accountant||'')}</span></div><div><strong>Thủ trưởng đơn vị</strong><span>${escapeHTML(s.head||'')}</span></div></div>`}
    <p class="receipt-print-note">Chứng từ được lập từ dữ liệu giao dịch đã đối soát trên hệ thống quản lý thu của nhà trường. Đây không phải hóa đơn điện tử hoặc biên lai điện tử theo pháp luật về hóa đơn, phí và lệ phí.</p>
  </article>`;
}
function printReceiptRecords(records,config){
  if(!records.length)return toast('Không có chứng từ để in.',true);
  const root=$('#receiptPrintRoot');root.innerHTML=records.sort((a,b)=>a.number.localeCompare(b.number)).map(r=>receiptRecordHtml(r,config)).join('');
  document.body.classList.add('printing-receipts');setTimeout(()=>{window.print();setTimeout(()=>document.body.classList.remove('printing-receipts'),400);},80);
}
async function cancelReceipt(id){
  const r=await request('receipts','get',id);if(!r||r.status!=='issued')return;
  const reason=prompt('Nhập lý do hủy chứng từ (bắt buộc):');if(reason===null)return;if(!reason.trim())return toast('Cần nhập lý do hủy để lưu dấu vết.',true);
  const updated={...r,status:'cancelled',cancelledAt:new Date().toISOString(),cancelReason:reason.trim()};await request('receipts','put',updated);
  await request('history','put',{id:crypto.randomUUID(),kind:'Hủy chứng từ',fileName:r.number,rows:1,imported:0,detail:`Đã hủy ${r.number}: ${reason.trim()}`,at:updated.cancelledAt});
  await refresh();toast(`Đã hủy ${r.number}. Số chứng từ được giữ trong lịch sử.`);
}
function receiptFilterCandidates(candidates,receipts){
  const fee=$('#receiptFeeFilter')?.value||'all',cls=$('#receiptClassFilter')?.value||'all',stu=$('#receiptStudentFilter')?.value||'all',status=$('#receiptStatusFilter')?.value||'all',method=$('#receiptMethodFilter')?.value||'all',from=$('#receiptDateFrom')?.value||'',to=$('#receiptDateTo')?.value||'';
  return candidates.filter(x=>{
    const state=receiptStateFor(x,receipts);
    const feeOk=fee==='all'||(fee.startsWith('catalog:')?x.item.catalogId===fee.slice(8):slug(x.item.name)===fee);
    const channel=x.txn.paymentChannel==='cash'?'cash':'transfer';
    return feeOk&&(cls==='all'||x.student.className===cls)&&(stu==='all'||x.student.code===stu)&&(status==='all'||state.status===status)&&(method==='all'||channel===method)&&(!from||(x.txn.date||'')>=from)&&(!to||(x.txn.date||'')<=to);
  });
}
function populateReceiptFilters(candidates,catalog){
  const feeEl=$('#receiptFeeFilter'),classEl=$('#receiptClassFilter'),stuEl=$('#receiptStudentFilter');if(!feeEl)return;
  const current={fee:feeEl.value,cls:classEl.value,stu:stuEl.value};
  const exact=[...new Map(candidates.map(x=>[x.item.catalogId?`catalog:${x.item.catalogId}`:slug(x.item.name),x.item])).entries()];
  feeEl.innerHTML='<option value="all">Tất cả khoản thu</option>'+exact.sort((a,b)=>a[1].name.localeCompare(b[1].name,'vi')).map(([v,item])=>`<option value="${escapeHTML(v)}">${escapeHTML(item.name)} · ${money(item.amount)}</option>`).join('');
  const classes=[...new Set(candidates.map(x=>x.student.className).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true}));
  classEl.innerHTML='<option value="all">Tất cả lớp</option>'+classes.map(v=>`<option value="${escapeHTML(v)}">${escapeHTML(v)}</option>`).join('');
  const students=[...new Map(candidates.map(x=>[x.student.code,x.student])).values()].sort((a,b)=>a.className.localeCompare(b.className,'vi',{numeric:true})||a.name.localeCompare(b.name,'vi'));
  stuEl.innerHTML='<option value="all">Tất cả học sinh</option>'+students.map(s=>`<option value="${escapeHTML(s.code)}">${escapeHTML(s.className)} · ${escapeHTML(s.name)}</option>`).join('');
  if([...feeEl.options].some(o=>o.value===current.fee))feeEl.value=current.fee;if([...classEl.options].some(o=>o.value===current.cls))classEl.value=current.cls;if([...stuEl.options].some(o=>o.value===current.stu))stuEl.value=current.stu;
}
async function renderReceiptPage(students,transactions){
  const [receipts,catalog,config]=await Promise.all([all('receipts'),getFeeCatalog(),getReceiptConfig()]);fillReceiptConfig(config);
  const candidates=receiptCandidates(students,transactions);populateReceiptFilters(candidates,catalog);renderCashEntry(students,transactions);
  const states=candidates.map(x=>receiptStateFor(x,receipts));$('#receiptReadyCount').textContent=states.filter(x=>x.status==='ready').length;$('#receiptIssuedCount').textContent=receipts.filter(r=>r.status==='issued').length;$('#receiptCancelledCount').textContent=receipts.filter(r=>r.status==='cancelled').length;
  const filtered=receiptFilterCandidates(candidates,receipts);
  const validTxnCount=transactions.filter(t=>t.paymentStatus==='valid').length;
  const unmatchedTxnCount=transactions.filter(t=>t.paymentStatus!=='valid').length;
  $('#receiptFilterSummary').textContent=filtered.length
    ?`${filtered.length} món đã thu phù hợp bộ lọc`
    :(validTxnCount? `${validTxnCount} giao dịch đã khớp nhưng không có món phù hợp bộ lọc.` : (transactions.length? `Đã có ${transactions.length} giao dịch nhưng ${unmatchedTxnCount} giao dịch chưa khớp học sinh/khoản thu.` : 'Chưa có giao dịch thu.'));
  $('#receiptTable').innerHTML=filtered.length?filtered.map(x=>{
    const state=receiptStateFor(x,receipts),r=state.receipt;const statusClass=state.status==='issued'?'paid':state.status==='cancelled'?'cancelled':'ready';
    const actions=state.status==='issued'?`<button class="text-button" data-receipt-print="${r.id}">In</button><button class="text-button danger-link" data-receipt-cancel="${r.id}">Hủy</button>`:`<button class="text-button" data-receipt-issue="${escapeHTML(x.paymentKey)}">${state.status==='cancelled'?'Phát hành lại':'Phát hành'}</button>`;
    return `<tr><td>${escapeHTML(x.txn.date||'—')}</td><td><strong>${escapeHTML(r?.number||'—')}</strong></td><td>${escapeHTML(x.student.name)}</td><td>${escapeHTML(x.student.className||'—')}</td><td>${escapeHTML(x.item.name)}</td><td><strong>${money(x.item.amount)}</strong></td><td>${escapeHTML(x.txn.ref||x.txn.id.slice(0,14))}</td><td><span class="receipt-status ${statusClass}">${receiptStatusLabel(state.status)}</span></td><td><div class="receipt-row-actions">${actions}</div></td></tr>`;
  }).join(''):'<tr><td colspan="9" class="empty-cell">Không có món đã thu phù hợp bộ lọc.</td></tr>';
}
async function exportReceiptRegister(){
  const receipts=(await all('receipts')).sort((a,b)=>(a.number||'').localeCompare(b.number||''));
  if(!receipts.length)return toast('Chưa có chứng từ để xuất sổ.',true);
  const headers=['Số chứng từ','Trạng thái','Ngày lập','Ngày thanh toán','Mã học sinh','Họ tên','Lớp','Khoản thu','Mã thanh toán','Số tiền','Mã giao dịch','Hình thức','Ngày hủy','Lý do hủy','Mã kiểm tra'];
  const lines=receipts.map(r=>[r.number,receiptStatusLabel(r.status),r.issuedAt,r.paymentDate,r.studentCode,r.studentName,r.className,r.feeName,r.paymentCode,r.amount,r.transactionRef,r.paymentMethod,r.cancelledAt||'',r.cancelReason||'',r.verificationCode].map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(','));
  download(`so-chung-tu-thu-${new Date().toISOString().slice(0,10)}.csv`,'\uFEFF'+headers.join(',')+'\r\n'+lines.join('\r\n'),'text/csv;charset=utf-8');
}

async function currentReceiptContext(){
  const [students,stored,receipts,config]=await Promise.all([all('students'),all('transactions'),all('receipts'),getReceiptConfig()]);
  const transactions=reconcileTransactions(students,stored),candidates=receiptCandidates(students,transactions);
  return {students,transactions,receipts,config,candidates,filtered:receiptFilterCandidates(candidates,receipts)};
}
async function issueFilteredReceipts(){
  const ctx=await currentReceiptContext(),issuable=ctx.filtered.filter(x=>!ctx.receipts.some(r=>r.paymentKey===x.paymentKey&&r.status==='issued'));
  if(!issuable.length)return toast('Không có món chưa phát hành trong bộ lọc hiện tại.',true);
  if(issuable.length>1&&!confirm(`Phát hành ${issuable.length} chứng từ và mở cửa sổ In/Lưu PDF?`))return;
  await issueReceiptCandidates(issuable,{print:true});
}
async function printIssuedFilteredReceipts(){
  const ctx=await currentReceiptContext(),keys=new Set(ctx.filtered.map(x=>x.paymentKey)),records=ctx.receipts.filter(r=>r.status==='issued'&&keys.has(r.paymentKey));
  if(!records.length)return toast('Không có chứng từ đã phát hành trong bộ lọc hiện tại.',true);
  printReceiptRecords(records,ctx.config);
}
async function handleReceiptTableClick(e){
  const issue=e.target.closest('[data-receipt-issue]'),print=e.target.closest('[data-receipt-print]'),cancel=e.target.closest('[data-receipt-cancel]');
  if(issue){const ctx=await currentReceiptContext(),candidate=ctx.candidates.find(x=>x.paymentKey===issue.dataset.receiptIssue);if(candidate)await issueReceiptCandidates([candidate],{print:true});return;}
  if(print){const [r,config]=await Promise.all([request('receipts','get',print.dataset.receiptPrint),getReceiptConfig()]);if(r)printReceiptRecords([r],config);return;}
  if(cancel)await cancelReceipt(cancel.dataset.receiptCancel);
}

async function refresh() {
  const [students,storedTransactions,history,qrConfig,noticeBundles]=await Promise.all([all('students'),all('transactions'),all('history'),request('meta','get','qrAccount'),getNoticeBundles()]);
  setNoticeBundleCache(noticeBundles);
  const transactions=reconcileTransactions(students,storedTransactions);
  const priorById=new Map(storedTransactions.map(t=>[t.id,t]));
  if(transactions.some(t=>{const old=priorById.get(t.id);return !old||['paymentStatus','matched','studentCode','studentName','matchedDueItem','matchedDueItemId','matchedDueItemIds','matchedBy'].some(key=>t[key]!==old[key]);})) await putMany('transactions',transactions);
  const t=totals(students,transactions);const classes=new Set(students.map(s=>s.className).filter(Boolean));
  $('#statStudents').textContent=students.length.toLocaleString('vi-VN');$('#statClasses').textContent=students.length?`${classes.size} lớp`:'Chưa có danh sách';
  $('#statFeeItems').textContent=t.dueItems.toLocaleString('vi-VN');$('#statDueAmount').textContent=`${money(t.due)} phải thu`;
  $('#statPaidItems').textContent=t.paidItems.toLocaleString('vi-VN');$('#statPaidAmount').textContent=`${money(t.paid)} đã ghi nhận`;
  $('#statUnpaidItems').textContent=t.unpaidItems.toLocaleString('vi-VN');$('#statRemainAmount').textContent=`${money(t.remain)} còn lại`;
  $('#completionPercent').textContent=`${t.pct}%`;
  $('#completionRing').style.background=`conic-gradient(var(--teal) ${t.pct*3.6}deg,#dce8e5 0deg)`;
  const legacy=students.some(s=>!s.hasFeeBreakdown);
  const notice=students.length?`${students.length} học sinh · ${transactions.filter(x=>x.paymentStatus==='valid').length} món thu khớp chính xác / ${transactions.length} giao dịch. ${legacy?'Danh sách cũ chỉ có tổng phải thu; hãy nhập lại file chi tiết từng khoản.':'Dữ liệu đang lưu riêng trên máy này.'}`:'Chọn danh sách học sinh và báo cáo thu để bắt đầu theo dõi.';
  $('#dataNoticeText').textContent=notice;
  renderFeeProgress(t.summaries);renderChart(transactions);renderClasses(students,transactions);renderStudents(students,transactions);renderFeeDetails(students,transactions,t.summaries);
  const feeCatalog=await getFeeCatalog();await populateFeeReportFilters(feeCatalog);await renderFeeSetup(students);renderQrPage(students,transactions,qrConfig);await renderNoticeTool(students,transactions,feeCatalog);await renderReceiptPage(students,transactions);
  const transactionHtml=transactions.length?renderTransactions(transactions):'<tr><td colspan="7" class="empty-cell">Chưa có báo cáo thu.</td></tr>';
  $('#transactionsTable').innerHTML=transactionHtml;$('#importsTransactionsTable').innerHTML=transactionHtml;
  ['allTxnCount','importsAllTxnCount'].forEach(id=>{const el=$(`#${id}`);if(el)el.textContent=transactions.length;});
  ['matchedTxnCount','importsMatchedTxnCount'].forEach(id=>{const el=$(`#${id}`);if(el)el.textContent=transactions.filter(x=>x.paymentStatus==='valid').length;});
  ['unmatchedTxnCount','importsUnmatchedTxnCount'].forEach(id=>{const el=$(`#${id}`);if(el)el.textContent=transactions.filter(x=>x.paymentStatus!=='valid').length;});
  $('#recentTransactions').innerHTML=transactions.length?[...transactions].sort((a,b)=>(b.importedAt||'').localeCompare(a.importedAt||'')).slice(0,4).map(x=>`<div class="recent-row"><strong><span class="category-badge ${transactionCategory(x)==='service'?'service':''}">${getFeeLabel(transactionCategory(x))}</span> ${escapeHTML(x.studentName||x.content||'Giao dịch thu')}</strong><span>${escapeHTML(x.date||'—')}</span><b>${money(x.amount)}</b></div>`).join(''):'<div class="empty-inline">Chưa có giao dịch được nhập.</div>';
  $('#historyTable').innerHTML=history.length?history.sort((a,b)=>b.at.localeCompare(a.at)).map(h=>`<tr><td>${dateTime(h.at)}</td><td>${escapeHTML(h.kind)}</td><td>${escapeHTML(h.fileName)}</td><td>${h.rows}</td><td>${escapeHTML(h.detail)}</td></tr>`).join(''):'<tr><td colspan="5" class="empty-cell">Chưa có lịch sử nhập file.</td></tr>';
  $('#storageStatus').textContent='Kho trình duyệt đã sẵn sàng';
}
function download(name, content, type='application/json') {
  const url=URL.createObjectURL(new Blob([content],{type}));const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
async function backup() {
  const password=prompt('Đặt mật khẩu cho tệp sao lưu (ít nhất 8 ký tự):');if(password===null)return;
  if(password.length<8)return toast('Mật khẩu cần có ít nhất 8 ký tự.',true);
  try {
    const data=JSON.stringify({version:2,createdAt:new Date().toISOString(),students:await all('students'),transactions:await all('transactions'),history:await all('history'),meta:await all('meta'),receipts:await all('receipts')});
    const salt=crypto.getRandomValues(new Uint8Array(16)),iv=crypto.getRandomValues(new Uint8Array(12));
    const base=await crypto.subtle.importKey('raw',new TextEncoder().encode(password),'PBKDF2',false,['deriveKey']);
    const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:250000,hash:'SHA-256'},base,{name:'AES-GCM',length:256},false,['encrypt']);
    const cipher=await crypto.subtle.encrypt({name:'AES-GCM',iv},key,new TextEncoder().encode(data));
    download(`so-thu-sao-luu-${new Date().toISOString().slice(0,10)}.sctbackup`,JSON.stringify({format:'SCTBACKUP1',salt:[...salt],iv:[...iv],cipher:[...new Uint8Array(cipher)]}));
    toast('Đã tạo bản sao lưu mã hóa. Hãy cất mật khẩu riêng.');
  } catch(e) { console.error(e);toast('Không tạo được tệp sao lưu.',true); }
}
async function restore(file) {
  const password=prompt('Nhập mật khẩu của bản sao lưu:');if(password===null)return;
  try {
    const payload=JSON.parse(await file.text());if(payload.format!=='SCTBACKUP1')throw new Error('Tệp sao lưu không đúng định dạng.');
    const salt=new Uint8Array(payload.salt),iv=new Uint8Array(payload.iv),cipher=new Uint8Array(payload.cipher);
    const base=await crypto.subtle.importKey('raw',new TextEncoder().encode(password),'PBKDF2',false,['deriveKey']);
    const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:250000,hash:'SHA-256'},base,{name:'AES-GCM',length:256},false,['decrypt']);
    const raw=await crypto.subtle.decrypt({name:'AES-GCM',iv},key,cipher);const data=JSON.parse(new TextDecoder().decode(raw));
    if(!confirm('Khôi phục sẽ thay thế toàn bộ dữ liệu hiện có trên máy này. Tiếp tục?'))return;
    await clearAll();await Promise.all([putMany('students',data.students||[]),putMany('transactions',data.transactions||[]),putMany('history',data.history||[]),putMany('meta',data.meta||[]),putMany('receipts',data.receipts||[])]);
    await refresh();toast('Đã khôi phục dữ liệu từ bản sao lưu.');
  } catch(e) { console.error(e);toast('Không mở được bản sao lưu. Kiểm tra đúng tệp và mật khẩu.',true); }
}

function excelCol(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
function excelXml(value,type='inlineStr',style=0){
  if(type==='n')return `<c s="${style}" t="n"><v>${Number(value)||0}</v></c>`;
  const v=escapeHTML(String(value??'')).replace(/&#39;/g,'&apos;');
  return `<c s="${style}" t="inlineStr"><is><t xml:space="preserve">${v}</t></is></c>`;
}
function excelRow(row,cells,height){
  return `<row r="${row}"${height?` ht="${height}" customHeight="1"`:''}>${cells.map((c,i)=>`<c r="${excelCol(i+1)}${row}" s="${c.s||0}" t="${c.t==='n'?'n':'inlineStr'}">${c.t==='n'?`<v>${Number(c.v)||0}</v>`:`<is><t xml:space="preserve">${escapeHTML(String(c.v??'')).replace(/&#39;/g,'&apos;')}</t></is>`}</c>`).join('')}</row>`;
}
async function exportClassFeeReport(){
  if(typeof JSZip==='undefined')return toast('Thiếu thư viện xuất Excel. Hãy tải lại trang.',true);
  const filter=$('#classFeeFilter')?.value||'all';
  if(filter==='all')return toast('Hãy chọn một món thu cụ thể trước khi xuất báo cáo.',true);
  const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);
  const transactions=reconcileTransactions(students,stored);
  const exactFee=filter.startsWith('catalog:')?catalog.find(f=>f.id===filter.slice(8)):null;
  const label=exactFee?.name||(filter==='mandatory'?'Nhóm BH thân thể (BHTT)':filter==='insurance'?'Nhóm Bảo hiểm y tế (BHYT)':filter==='service'?'Nhóm Dịch vụ khác':'Nhóm khoản khác');
  const groups=[...new Set(students.map(s=>s.className||'Chưa xếp lớp'))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true,sensitivity:'base'}));
  const summary=[];const details=[];
  for(const className of groups){
    const classStudents=students.filter(s=>(s.className||'Chưa xếp lớp')===className);
    let dueItems=0,paidItems=0,due=0,paid=0;
    for(const student of classStudents){
      const items=studentDueItems(student).filter(item=>itemMatchesFeeFilter(item,filter));
      const itemIds=new Set(items.map(item=>item.id));
      const valid=transactions.filter(t=>t.studentCode===student.code&&t.paymentStatus==='valid'&&transactionMatchedItemIds(t).some(id=>itemIds.has(id)));
      const paidKeys=new Set(valid.flatMap(transactionMatchedItemIds));
      for(const item of items){
        const isPaid=paidKeys.has(item.id);
        const txn=valid.find(t=>transactionMatchesItem(t,item.id));
        dueItems++;due+=item.amount;if(isPaid){paidItems++;paid+=item.amount;}
        details.push([className,student.classStudentCode||'',student.code,student.name,item.name,item.feeCode||item.category||'',item.paymentCode||'',item.amount,isPaid?'Đã thu':'Chưa thu',isPaid?item.amount:0,isPaid?0:item.amount,txn?.date||'',txn?.ref||'']);
      }
    }
    summary.push([className,classStudents.length,dueItems,paidItems,dueItems-paidItems,due,paid,Math.max(0,due-paid),due?paid/due:0]);
  }
  const zip=new JSZip();
  zip.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.folder('_rels').file('.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.folder('xl').file('workbook.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Tổng hợp theo lớp" sheetId="1" r:id="rId1"/><sheet name="Chi tiết học sinh" sheetId="2" r:id="rId2"/></sheets></workbook>');
  zip.folder('xl').folder('_rels').file('workbook.xml.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.folder('xl').file('styles.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="4"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FF063F3B"/><name val="Arial"/></font></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF063F3B"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF078778"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8F4F0"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFE1E9E6"/></left><right style="thin"><color rgb="FFE1E9E6"/></right><top style="thin"><color rgb="FFE1E9E6"/></top><bottom style="thin"><color rgb="FFE1E9E6"/></bottom></border></borders><cellXfs count="8"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="3" fontId="0" fillId="0" borderId="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1"/><xf numFmtId="10" fontId="3" fillId="4" borderId="1"/><xf numFmtId="0" fontId="3" fillId="4" borderId="1"/><xf numFmtId="3" fontId="3" fillId="4" borderId="1"/></cellXfs></styleSheet>');
  const summaryRows=[];
  summaryRows.push(excelRow(1,[{v:'BÁO CÁO TIẾN ĐỘ THU THEO LỚP',s:1}],28));
  summaryRows.push(excelRow(2,[{v:`Khoản thu: ${label}`,s:6}],22));
  summaryRows.push(excelRow(3,[{v:`Ngày xuất: ${new Intl.DateTimeFormat('vi-VN',{dateStyle:'short',timeStyle:'short'}).format(new Date())}`,s:4}],20));
  summaryRows.push(excelRow(5,['LỚP','SĨ SỐ','MÓN PHẢI THU','ĐÃ THU','CHƯA THU','PHẢI THU (Đ)','ĐÃ THU (Đ)','CÒN LẠI (Đ)','TIẾN ĐỘ'].map(v=>({v,s:2})),28));
  summary.forEach((r,idx)=>summaryRows.push(excelRow(idx+6,r.map((v,j)=>({v,t:j===0?'s':'n',s:j===8?5:(j>=5?3:4)})),21)));
  const total=summary.reduce((o,r)=>r.map((v,i)=>i===0?'TỔNG':i===8?0:(o[i]||0)+(Number(v)||0)),[]);
  total[8]=total[5]?total[6]/total[5]:0;
  summaryRows.push(excelRow(summary.length+6,total.map((v,j)=>({v,t:j===0?'s':'n',s:j===8?5:(j>=5?7:6)})),24));
  const sheet1=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="5" width="14" customWidth="1"/><col min="6" max="8" width="18" customWidth="1"/><col min="9" max="9" width="13" customWidth="1"/></cols><sheetData>${summaryRows.join('')}</sheetData><autoFilter ref="A5:I${summary.length+5}"/><mergeCells count="3"><mergeCell ref="A1:I1"/><mergeCell ref="A2:I2"/><mergeCell ref="A3:I3"/></mergeCells></worksheet>`;
  zip.folder('xl').folder('worksheets').file('sheet1.xml',sheet1);
  const detailRows=[];
  detailRows.push(excelRow(1,[{v:`CHI TIẾT HỌC SINH – ${label}`,s:1}],28));
  detailRows.push(excelRow(3,['LỚP','MÃ HS THEO LỚP','MÃ HS GỐC','HỌ VÀ TÊN','KHOẢN THU','MÃ KHOẢN','MÃ THANH TOÁN','PHẢI THU (Đ)','TRẠNG THÁI','ĐÃ THU (Đ)','CÒN LẠI (Đ)','NGÀY THU','MÃ GIAO DỊCH'].map(v=>({v,s:2})),28));
  details.forEach((r,idx)=>detailRows.push(excelRow(idx+4,r.map((v,j)=>({v,t:[4,6,7].includes(j)?'n':'s',s:[4,6,7].includes(j)?3:4})),21)));
  const sheet2=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="11" customWidth="1"/><col min="2" max="2" width="17" customWidth="1"/><col min="3" max="3" width="28" customWidth="1"/><col min="4" max="4" width="25" customWidth="1"/><col min="5" max="8" width="16" customWidth="1"/><col min="9" max="9" width="13" customWidth="1"/><col min="10" max="10" width="22" customWidth="1"/></cols><sheetData>${detailRows.join('')}</sheetData><autoFilter ref="A3:J${details.length+3}"/><mergeCells count="1"><mergeCell ref="A1:J1"/></mergeCells></worksheet>`;
  zip.folder('xl').folder('worksheets').file('sheet2.xml',sheet2);
  const blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const safe=exactFee?feeSafeCode(exactFee.code||exactFee.name,20):(filter==='mandatory'?'BHTT':filter==='insurance'?'BHYT':filter==='service'?'DICH-VU':'KHAC');
  const url=URL.createObjectURL(blob);const link=document.createElement('a');link.href=url;link.download=`Bao-cao-thu-${safe}-theo-lop.xlsx`;link.click();setTimeout(()=>URL.revokeObjectURL(url),1200);
  toast(`Đã xuất báo cáo Excel ${label}.`);
}

function exportStudents() {
  Promise.all([all('students'),all('transactions')]).then(([items,stored])=>{
    const tx=reconcileTransactions(items,stored),paid=new Set(tx.filter(t=>t.paymentStatus==='valid').flatMap(transactionPaidKeys));
    const headers=['Mã học sinh gốc','Mã HS theo lớp','Mã số CĐ / mã ngoài','Họ và tên','Lớp','Giới tính','Ngày sinh','Phân loại HS','Đăng ký khoản khác','Ghi chú','Điện thoại','Số món phải thu','Tổng phải thu','Đã thu','Còn lại'];
    const lines=items.map(s=>{const fees=studentDueItems(s),due=fees.reduce((a,x)=>a+x.amount,0),paidAmount=fees.filter(x=>paid.has(`${s.code}|${x.id}`)).reduce((a,x)=>a+x.amount,0);return [s.code,s.classStudentCode||'',s.externalCode||'',s.name,s.className,s.gender||'',s.birthDate||'',s.studentType||'',s.registrationInfo||'',s.note||'',s.phone||'',fees.length,due,paidAmount,Math.max(0,due-paidAmount)].map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',');});
    download('danh-sach-hoc-sinh-thcs-so-2-hieu-giang.csv','\uFEFF'+headers.join(',')+'\r\n'+lines.join('\r\n'),'text/csv;charset=utf-8');
  });
}
async function migrateSchool2StudentData(){
  const students=await all('students');
  if(!students.length)return;
  const stable=students.filter(s=>!String(s.code||'').startsWith('TMP-'));
  const codeAllocator=createStudentCodeAllocator(stable);
  const classAllocator=createClassStudentCodeAllocator(students);
  const remap=new Map(),updated=[],deleteCodes=[];
  for(const s of students){
    let code=s.code;
    if(String(code||'').startsWith('TMP-')){const old=code;code=codeAllocator.next();remap.set(old,code);deleteCodes.push(old);}
    let classStudentCode=s.classStudentCode;
    if(!classAllocator.validForClass(classStudentCode,s.className))classStudentCode=classAllocator.next(s.className||'Chưa xếp lớp');
    const clean={
      ...s,code,classStudentCode,
      personalId:'',ethnicity:'',fatherName:'',motherName:'',phone:'',
      dataWarnings:studentWarnings({...s,code,classStudentCode,dataWarnings:[]})
    };
    updated.push(clean);
  }
  for(const old of deleteCodes)await request('students','delete',old);
  await putMany('students',updated);
  if(remap.size){
    const txs=await all('transactions');
    await putMany('transactions',txs.map(t=>remap.has(t.studentCode)?{...t,studentCode:remap.get(t.studentCode)}:t));
    const receipts=await all('receipts');
    await putMany('receipts',receipts.map(r=>remap.has(r.studentCode)?{...r,studentCode:remap.get(r.studentCode)}:r));
  }
}

function wire() {
  $$('.nav-item[data-page]').forEach(btn=>btn.addEventListener('click',()=>setPage(btn.dataset.page)));
  $$('[data-go]').forEach(btn=>btn.addEventListener('click',()=>setPage(btn.dataset.go)));
  $('#studentImportButton').onclick=$('#studentImportButton2').onclick=()=>$('#studentFileInput').click();
  $('#bankImportButton').onclick=()=>$('#bankFileInput').click();
  $('#studentFileInput').onchange=e=>{chooseFile('students',e.target.files[0]);e.target.value='';};
  $('#bankFileInput').onchange=e=>{chooseFile('bank',e.target.files[0]);e.target.value='';};
  $('#restoreFileInput').onchange=e=>{restore(e.target.files[0]);e.target.value='';};
  $('#modalClose').onclick=$('#modalCancel').onclick=closeModal;$('#modalConfirm').onclick=confirmImport;
  $('#modalBackdrop').addEventListener('click',e=>{if(e.target.id==='modalBackdrop')closeModal();});
  $('#studentSearch').addEventListener('input',async()=>{const [students,stored]=await Promise.all([all('students'),all('transactions')]);renderStudents(students,reconcileTransactions(students,stored));});
  $('#showIncompleteStudents').onclick=async()=>{showIncompleteOnly=!showIncompleteOnly;const [students,stored]=await Promise.all([all('students'),all('transactions')]);renderStudents(students,reconcileTransactions(students,stored));};
  $('#studentsTable').addEventListener('click',e=>{const btn=e.target.closest('[data-student-code]');if(btn)openStudentProfile(btn.dataset.studentCode);});
  $('#studentProfileClose').onclick=closeStudentProfile;$('#studentProfileBackdrop').addEventListener('click',e=>{if(e.target.id==='studentProfileBackdrop')closeStudentProfile();});
  $('#feeScope').addEventListener('change',async()=>populateFeeTargets(await all('students')));
  $('#feeTargets').addEventListener('change',async()=>{const students=await all('students');const chosen=feeTargetStudents(students,$('#feeScope').value,selectedValues($('#feeTargets')));$('#feeTargetSummary').textContent=`${chosen.length} học sinh được chọn`;});
  ['feePrefix','feeCode','feeShortCode'].forEach(id=>$(`#${id}`).addEventListener('input',async()=>updateFeePreview(await all('students'))));
  $('#feeCategory').addEventListener('change',async()=>{const code=defaultFeeShortCode($('#feeCategory').value,$('#feeCode').value);$('#feeShortCode').value=code;updateFeePreview(await all('students'));});
  $('#bidvExportScope').addEventListener('change',async()=>{const [students,catalog]=await Promise.all([all('students'),getFeeCatalog()]);populateBidvExportControls(students,catalog);});
  ['bidvExportFee','bidvExportTarget'].forEach(id=>$(`#${id}`).addEventListener('change',async()=>{const [students,catalog]=await Promise.all([all('students'),getFeeCatalog()]);updateBidvCustomerPreview(students,catalog);}));
  ['bidvBillPeriod','bidvCustomerPrefix','bidvCustomerTemplate'].forEach(id=>$(`#${id}`).addEventListener('input',async()=>{const [students,catalog]=await Promise.all([all('students'),getFeeCatalog()]);updateBidvCustomerPreview(students,catalog);}));
  $('#bidvValidateExport').onclick=()=>validateBidvExport().catch(e=>{console.error(e);$('#bidvExportStatus').textContent=e.message||'Dữ liệu chưa hợp lệ.';toast(e.message||'Dữ liệu chưa hợp lệ.',true);});
  $('#bidvExportXlsx').onclick=()=>exportBidvXlsx().catch(e=>{console.error(e);$('#bidvExportStatus').textContent=e.message||'Không xuất được bảng kê.';toast(e.message||'Không xuất được bảng kê.',true);});
  $('#qrFileScope').addEventListener('change',async()=>{const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);populateQrFileControls(students,catalog,reconcileTransactions(students,stored));});
  $('#qrFileTarget').addEventListener('change',async()=>{const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);updateQrFilePreview(students,catalog,reconcileTransactions(students,stored));});
  ['qrFilePrefix','qrFilePeriod'].forEach(id=>$(`#${id}`).addEventListener('input',async()=>{const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);updateQrFilePreview(students,catalog,reconcileTransactions(students,stored));}));
  $('#qrFileFee').addEventListener('change',async()=>{const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);updateQrFilePreview(students,catalog,reconcileTransactions(students,stored));});
  $('#qrFileStudentCodeMode').addEventListener('change',async()=>{const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);updateQrFilePreview(students,catalog,reconcileTransactions(students,stored));});
  $('#qrFileOnlyUnpaid').addEventListener('change',async()=>{const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);updateQrFilePreview(students,catalog,reconcileTransactions(students,stored));});
  $('#qrFileValidate').onclick=()=>validateQrFileExport().catch(e=>{console.error(e);$('#qrFileStatus').textContent=e.message||'File QR chưa hợp lệ.';toast(e.message||'File QR chưa hợp lệ.',true);});
  $('#qrFileExport').onclick=()=>exportQrFileXlsx().catch(e=>{console.error(e);$('#qrFileStatus').textContent=e.message||'Không xuất được file QR.';toast(e.message||'Không xuất được file QR.',true);});
  $('#saveFeeAssignment').onclick=()=>saveFeeAssignment().catch(e=>{console.error(e);toast(e.message||'Không tạo được khoản thu.',true);});
  $('#newFeeButton').onclick=async()=>resetFeeForm(await all('students'));$('#cancelFeeEdit').onclick=async()=>resetFeeForm(await all('students'));
  $('#feeCatalogList').addEventListener('click',e=>{const edit=e.target.closest('[data-fee-edit]'),del=e.target.closest('[data-fee-delete]');if(edit)editFee(edit.dataset.feeEdit);if(del)deleteFee(del.dataset.feeDelete);});
  const refreshNoticeUi=async()=>{const [students,stored,catalog]=await Promise.all([all('students'),all('transactions'),getFeeCatalog()]);const tx=reconcileTransactions(students,stored);populateNoticeControls(students,catalog);updateNoticeSummary(students,tx);};
  $('#noticeClass').addEventListener('change',refreshNoticeUi);$('#noticeStudent').addEventListener('change',refreshNoticeUi);$('#noticeDueStatus').addEventListener('change',refreshNoticeUi);$('#noticeFees').addEventListener('change',e=>{const selected=[...e.currentTarget.selectedOptions];if(selected.length>1&&selected.some(o=>o.value==='all'))e.currentTarget.querySelector('option[value="all"]').selected=false;refreshNoticeUi();});$('#noticeDeadline').addEventListener('change',refreshNoticeUi);
  $$('input[name="noticeTemplate"]').forEach(r=>r.addEventListener('change',()=>{$$('.notice-template-option').forEach(x=>x.classList.toggle('selected',x.querySelector('input').checked));if(window.__noticeEntries)renderNoticePreview(window.__noticeEntries);}));
  $('#previewNotices').onclick=()=>previewNotices().catch(e=>{console.error(e);toast(e.message||'Không tạo được xem trước.',true);});$('#printNoticeA4').onclick=()=>printNoticeA4().catch(e=>{console.error(e);toast(e.message||'Không tạo được bản A4.',true);});$('#downloadNoticeImages').onclick=()=>downloadNoticeImages().catch(e=>{console.error(e);toast(e.message||'Không xuất được ảnh thông báo.',true);});
  $('#saveReceiptConfig').onclick=saveReceiptConfig;$('#receiptConfigToggle').onclick=()=>$('#receiptConfigCard').classList.toggle('collapsed');
  ['receiptFeeFilter','receiptClassFilter','receiptStudentFilter','receiptStatusFilter','receiptMethodFilter','receiptDateFrom','receiptDateTo'].forEach(id=>$(`#${id}`).addEventListener('change',async()=>{const [students,stored]=await Promise.all([all('students'),all('transactions')]);await renderReceiptPage(students,reconcileTransactions(students,stored));}));
  $('#cashClass').addEventListener('change',refreshCashEntry);$('#cashStudent').addEventListener('change',refreshCashEntry);$('#cashFee').addEventListener('change',refreshCashEntry);
  $('#recordCashPayment').onclick=()=>recordCashPayment(false).catch(e=>{console.error(e);toast(e.message||'Không ghi nhận được tiền mặt.',true);});$('#recordAndPrintCash').onclick=()=>recordCashPayment(true).catch(e=>{console.error(e);toast(e.message||'Không ghi nhận/in được phiếu thu.',true);});
  $('#issueFilteredReceipts').onclick=()=>issueFilteredReceipts().catch(e=>{console.error(e);toast(e.message||'Không phát hành được chứng từ.',true);});$('#printIssuedReceipts').onclick=()=>printIssuedFilteredReceipts().catch(e=>{console.error(e);toast('Không in được chứng từ.',true);});$('#exportReceiptRegister').onclick=exportReceiptRegister;$('#receiptTable').addEventListener('click',e=>handleReceiptTableClick(e).catch(err=>{console.error(err);toast(err.message||'Không thực hiện được thao tác chứng từ.',true);}));
  $('#classFeeFilter').addEventListener('change',async()=>{const [students,stored]=await Promise.all([all('students'),all('transactions')]);renderClasses(students,reconcileTransactions(students,stored));});
  $('#exportClassFeeReport').onclick=()=>exportClassFeeReport().catch(e=>{console.error(e);toast('Không xuất được báo cáo Excel.',true);});
  $('#classTable').addEventListener('click',e=>{const row=e.target.closest('.class-summary-row');if(row)openClassDetail(row.dataset.className);});
  $('#classTable').addEventListener('keydown',e=>{const row=e.target.closest('.class-summary-row');if(row&&(e.key==='Enter'||e.key===' ')){e.preventDefault();openClassDetail(row.dataset.className);}});
  $('#classDetailClose').onclick=closeClassDetail;
  $('#classDetailBackdrop').addEventListener('click',e=>{if(e.target.id==='classDetailBackdrop')closeClassDetail();});
  $('#saveQrConfig').onclick=saveQrConfig;$('#generateQrs').onclick=()=>generateQrs().catch(e=>{console.error(e);toast(e.message||'Không tạo được mã QR.',true);});
  ['qrFeeFilter','qrClassFilter','qrStatusFilter'].forEach(id=>$(`#${id}`).addEventListener('change',()=>$('#downloadQrs').hidden=true));
  $('#exportStudents').onclick=exportStudents;$('#backupButton').onclick=$('#backupButtonTop').onclick=backup;$('#restoreButton').onclick=()=>$('#restoreFileInput').click();
  $('#clearDataButton').onclick=async()=>{if(confirm('Xóa toàn bộ dữ liệu học sinh, giao dịch và lịch sử trên trình duyệt này?')){await clearAll();await refresh();toast('Đã xóa dữ liệu trên máy này.');}};
  $('#menuToggle').onclick=()=>$('#sidebar').classList.toggle('open');
  if('serviceWorker' in navigator&&location.protocol.startsWith('http'))navigator.serviceWorker.register('sw.js').catch(err=>console.warn('Offline cache:',err));
}
document.addEventListener('DOMContentLoaded',async()=>{
  if(!$('#storageStatus'))return; // Standalone visual tests do not open the production database.
  try { db=await openDatabase();await migrateSchool2StudentData();wire();await refresh(); }
  catch(error){console.error(error);$('#storageStatus').textContent='Không mở được kho dữ liệu';toast('Trình duyệt không cho phép lưu dữ liệu cục bộ.',true);}
});
