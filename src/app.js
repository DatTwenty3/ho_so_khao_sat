/* Giao diện. */
'use strict';

const KEY = 'hoso-khaosat-v1';
const $ = id => document.getElementById(id);
const $$ = sel => [...document.querySelectorAll(sel)];
const b64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0)).buffer;
const BUILTIN = { nk: b64(ASSETS.nk), bb: b64(ASSETS.bb) };
const templates = { ...BUILTIN };
const tplNames = { nk: null, bb: null };

const blank = () => ({ schedule: [], edits: {}, info: {}, convert: true, noHighlight: true, dirty: false, savedAt: null });
let state = blank();
let previous = null;                 // phiên đã tự lưu, chỉ nạp khi người dùng bấm "Tiếp tục"
let tab = 'lich', filter = 'all', selected = null;
let defaults = {};                   // thông tin chung đọc từ mẫu đang dùng

const el = (tag, props = {}, ...children) => {
  const node = Object.assign(document.createElement(tag), props);
  node.append(...children);
  return node;
};

function icon(name) {
  const NS = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(NS, 'svg'), use = document.createElementNS(NS, 'use');
  svg.setAttribute('class', 'ic');
  svg.setAttribute('aria-hidden', 'true');
  use.setAttribute('href', '#i-' + name);
  svg.append(use);
  return svg;
}

/* Chiều cao phần đang neo, để thanh công cụ và tiêu đề bảng xếp nối tiếp nhau. */
function syncOffsets() {
  const header = document.querySelector('body > header');
  const head = getComputedStyle(header).position === 'sticky' ? header.offsetHeight : 0;
  const bar = document.querySelector('section.page:not([hidden]) .toolbar')?.offsetHeight || 0;
  document.documentElement.style.setProperty('--head', head + 'px');
  document.documentElement.style.setProperty('--bar', bar + 'px');
}

/* ---------- Thông báo ---------- */
let toastTimer = 0;
function toast(text, { error = false, action = null, ms = 3500 } = {}) {
  const box = $('toast');
  box.className = error ? 'error' : '';
  box.replaceChildren(icon(error ? 'alert' : 'check'), el('span', { textContent: text }));
  if (action) box.append(el('button', { textContent: action.label, onclick: () => { box.hidden = true; action.run(); } }));
  box.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { box.hidden = true; }, error ? 8000 : ms);
}

/* ---------- Lưu trữ ---------- */
function persist() {
  try {
    if (state.schedule.length) {
      state.savedAt = new Date().toISOString();
      localStorage.setItem(KEY, JSON.stringify(state));
      previous = null;
    }
  } catch { toast('Trình duyệt không cho tự lưu — hãy bấm "Lưu".', { error: true }); }
}
/* Người dùng sửa dữ liệu (khác với việc chỉ nạp vào). */
function changed() { state.dirty = true; persist(); updateCounts(); }

function download(blob, name) {
  const a = el('a', { href: URL.createObjectURL(blob), download: name });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}
async function run(fn) {
  try { toast('Đang tạo file…', { ms: 60000 }); await fn(); toast('Đã tạo xong, kiểm tra thư mục Tải xuống.'); }
  catch (e) { console.error(e); toast('Lỗi: ' + e.message, { error: true }); }
}

/* ---------- Dữ liệu dẫn xuất ---------- */
/* Biên bản được sửa tay lưu theo ngày + nội dung của dòng, để chèn/xoá dòng khác không lệch. */
const editKey = row => row.date + '|' + row.content;
/* Số biên bản luôn theo STT, trừ khi người dùng tự gõ số khác. Bản sửa lưu từ trước
   không có cờ numberEdited: '02' khi đó là số mặc định nên cũng thay theo STT. */
function record(i) {
  const edit = state.edits[editKey(state.schedule[i])];
  const auto = suggestRecord(state.schedule, i, state);
  if (!edit) return auto;
  const own = (edit.numberEdited ?? edit.number !== '02') && String(edit.number).trim();
  return own ? edit : { ...edit, number: auto.number };
}
/* Chèn hoặc xoá dòng thì STT các dòng phía sau dồn theo. */
const shiftNo = (from, step) => state.schedule.slice(from).forEach(r => { if (r.stt) r.stt += step; });
const diaryRows = () => state.schedule.filter(r => r.nk && r.date);
const recordIdx = () => state.schedule.map((r, i) => i).filter(i => state.schedule[i].nt && state.schedule[i].date);
const weekday = iso => iso ? ['Chủ nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'][new Date(iso + 'T00:00:00Z').getUTCDay()] : '';
const dayLabel = iso => iso ? `${weekday(iso)}, ${dmy(iso)}` : 'Chưa chọn ngày';

function fileNames(list, prefix) {
  const seen = {};
  return list.map(date => {
    seen[date] = (seen[date] || 0) + 1;
    return `${prefix} ${date}${seen[date] > 1 ? ` (${seen[date]})` : ''}.docx`;
  });
}

const makeDiary = row => buildDoc(templates.nk, fillDiaryXml, row, state);
const makeRecord = rec => buildDoc(templates.bb, fillRecordXml, rec, state);

async function buildAll() {
  const out = new Map();
  const bytes = async blob => new Uint8Array(await blob.arrayBuffer());
  const nk = diaryRows(), nkNames = fileNames(nk.map(r => r.date), 'NK');
  for (const [i, row] of nk.entries()) out.set('Nhật ký/' + nkNames[i], await bytes(await makeDiary(row)));
  const recs = recordIdx().map(record), bbNames = fileNames(recs.map(r => r.date), 'BBNT');
  for (const [i, rec] of recs.entries()) out.set('Biên bản nghiệm thu/' + bbNames[i], await bytes(await makeRecord(rec)));
  return out;
}

const exporters = {
  all: () => run(async () => {
    const files = await buildAll();
    if (!files.size) throw new Error('Chưa có ngày nào đánh dấu Nhật ký hoặc Nghiệm thu');
    download(zip(files), 'ho-so-khao-sat.zip');
  }),
  merged: () => run(async () => {
    const rows = diaryRows();
    if (!rows.length) throw new Error('Chưa có ngày nào đánh dấu Nhật ký');
    download(await buildMerged(templates.nk, fillDiaryXml, rows, state), 'Nhat ky khao sat.docx');
  }),
};

/* ---------- Khung chung ---------- */
function updateCounts() {
  const nk = diaryRows().length, bb = recordIdx().length;
  $('nLich').textContent = state.schedule.length;
  $('nBb').textContent = bb;
  $('nNk').textContent = nk;
  $('allInfo').textContent = `${nk} nhật ký và ${bb} biên bản nghiệm thu, chia sẵn thành hai thư mục.`;
  const dates = state.schedule.map(r => r.date).filter(Boolean).sort();
  $('summary').textContent = dates.length
    ? `${state.schedule.length} ngày công việc · ${dmy(dates[0])} → ${dmy(dates[dates.length - 1])}`
    : 'Nhật ký và biên bản nghiệm thu từ lịch công việc';
}

function show() {
  const has = state.schedule.length > 0;
  $('empty').hidden = has;
  $('tabs').hidden = !has;
  $('actions').hidden = !has;              // màn hình trống đã có sẵn các nút này
  $$('section.page').forEach(s => { s.hidden = !has || s.id !== tab; });
  $$('.tab').forEach(t => t.classList.toggle('on', t.dataset.tab === tab));
  $('noHl').checked = state.noHighlight;
  $('convert').checked = state.convert;
  updateCounts();
  syncOffsets();
  if (!has) {
    const when = previous?.savedAt ? `, lưu lúc ${new Date(previous.savedAt).toLocaleString('vi-VN')}` : '';
    $('restore').hidden = !previous;
    if (previous) $('restore').replaceChildren(icon('reset'), `Tiếp tục phiên trước (${previous.schedule.length} dòng${when})`);
    return;
  }
  if (tab === 'tt') renderInfo();
  if (tab === 'lich') renderSchedule();
  if (tab === 'bb') renderRecords();
  if (tab === 'nk') renderDiaries();
}

function go(name) { tab = name; show(); window.scrollTo(0, 0); }

/* ---------- Thông tin chung ---------- */
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const infoValue = key => state.info[key] ?? defaults[key] ?? '';

/* Chỉ lưu những gì khác mẫu. */
function setInfo(key, value) {
  if (same(value, defaults[key] ?? '')) delete state.info[key];
  else state.info[key] = value;
  changed();
}

async function loadDefaults() {
  try { defaults = await readDefaults(templates); }
  catch (err) { console.error(err); defaults = {}; toast('Không đọc được thông tin trong mẫu: ' + err.message, { error: true }); }
  if (state.schedule.length && tab === 'tt') renderInfo();
}

function renderInfo() {
  const boxes = {};
  const mark = key => boxes[key]?.classList.toggle('edited', state.info[key] !== undefined);
  const field = (key, label, { long = false, list = '' } = {}) => {
    const node = el(long ? 'textarea' : 'input', long ? { rows: 1 } : { type: 'text' });
    node.value = infoValue(key);
    if (list) node.setAttribute('list', list);
    boxes[key] = node;
    mark(key);
    node.oninput = () => { setInfo(key, node.value); mark(key); };
    return el('label', {}, label, node);
  };

  const people = (key, signKey, listId) => {
    const box = el('div', { className: 'members' });
    const options = el('datalist', { id: listId });
    const copy = () => structuredClone(infoValue(key) || []);
    const paintOptions = () => options.replaceChildren(...copy().map(p => el('option', { value: p.name })));
    const commit = list => {
      setInfo(key, list);
      box.classList.toggle('edited-list', state.info[key] !== undefined);
      paintOptions();
    };
    const draw = () => {
      box.classList.toggle('edited-list', state.info[key] !== undefined);
      box.replaceChildren(...copy().map((person, i) => {
        const edit = (prop, value) => {
          const list = copy(), old = list[i].name;
          list[i][prop] = value;
          commit(list);
          if (prop === 'name' && old && infoValue(signKey) === old) {   // người ký đổi tên theo
            setInfo(signKey, value);
            boxes[signKey].value = value;
            mark(signKey);
          }
        };
        const title = el('select', { onchange: e => edit('title', e.target.value) },
          el('option', { textContent: 'Ông' }), el('option', { textContent: 'Bà' }));
        title.value = person.title;
        title.setAttribute('aria-label', 'Danh xưng');
        const text = (prop, hint) => {
          const node = el('input', { type: 'text', value: person[prop], placeholder: hint, oninput: e => edit(prop, e.target.value) });
          node.setAttribute('aria-label', hint);
          return node;
        };
        return el('div', { className: 'member' }, title, text('name', 'Họ và tên'), text('role', 'Chức vụ'),
          el('button', { className: 'btn ghost danger icon', title: 'Xoá người này', onclick: () => {
            const list = copy();
            list.splice(i, 1);
            commit(list);
            draw();
          } }, icon('close')));
      }), el('button', { className: 'btn', onclick: () => {
        commit([...copy(), { title: 'Ông', name: '', role: '' }]);
        draw();
      } }, icon('plus'), 'Thêm người'), options);
      paintOptions();
    };
    draw();
    return box;
  };

  const card = (title, ...children) => el('div', { className: 'card' }, el('h2', { textContent: title }), ...children);
  const cols = (...children) => el('div', { className: 'cols' }, ...children);
  const wide = { long: true };

  $('infoForm').replaceChildren(
    card('Công trình',
      field('projectName', 'Tên công trình / dự án', wide),
      field('location', 'Địa điểm xây dựng', wide),
      cols(field('phase', 'Bước thiết kế'), field('place', 'Nơi lập biên bản'))),
    card('Hợp đồng',
      cols(field('contractNo', 'Số hợp đồng'), field('contractDate', 'Ngày ký')),
      field('investor', 'Bên giao (chủ đầu tư / ban quản lý)', wide),
      field('contractor', 'Bên nhận (nhà thầu / liên danh)', wide),
      field('contractAbout', 'Về việc', wide)),
    card('Nhà thầu giám sát khảo sát',
      field('supervisorCompany', 'Tên công ty', wide),
      people('supervisors', 'signSupervisor', 'dsGiamSat'),
      field('signSupervisor', 'Người ký phía giám sát', { list: 'dsGiamSat' })),
    card('Nhà thầu khảo sát',
      field('surveyCompany', 'Tên công ty', wide),
      people('surveyors', 'signSurveyor', 'dsKhaoSat'),
      field('signSurveyor', 'Người ký phía khảo sát', { list: 'dsKhaoSat' })),
    card('Nội dung cố định trong nhật ký',
      field('manpower', 'Nhân lực', wide),
      field('equipment', 'Thiết bị', wide),
      cols(field('weather', 'Thời tiết'), field('remark', 'Nhận xét')),
      cols(field('issues', 'Vấn đề vướng mắc'), field('solution', 'Biện pháp xử lý'))));
}

/* ---------- Lịch công việc ---------- */
function renderSchedule() {
  const q = $('search').value.trim().toLowerCase();
  const rows = state.schedule.map((row, i) => [row, i]).filter(([row]) =>
    (filter === 'all' || row.nt) &&
    (!q || `${row.content}\n${row.qty}\n${row.date ? dmy(row.date) : ''}`.toLowerCase().includes(q)));

  $('rows').replaceChildren(...rows.map(([row, i]) => {
    const tr = el('tr', { className: row.nt ? 'nt' : '' });
    tr.dataset.i = i;
    const chip = (key, label) => {
      const b = el('button', { type: 'button', className: 'chip ' + key });
      const paint = () => { b.replaceChildren(icon(row[key] ? 'check' : 'plus'), label); b.setAttribute('aria-pressed', String(!!row[key])); };
      b.onclick = () => { row[key] = !row[key]; paint(); tr.className = row.nt ? 'nt' : ''; changed(); };
      paint();
      return b;
    };
    const day = el('div', { className: 'day', textContent: weekday(row.date) });
    const field = (key, node) => {
      node.value = row[key];
      node.setAttribute('aria-label', { date: 'Ngày', content: 'Nội dung công việc', qty: 'Khối lượng' }[key]);
      node.oninput = () => {
        const old = editKey(row);
        row[key] = node.value;
        if (state.edits[old] && editKey(row) !== old) { state.edits[editKey(row)] = state.edits[old]; delete state.edits[old]; }
        if (key === 'date') day.textContent = weekday(row.date);
        changed();
      };
      return node;
    };
    const area = key => field(key, el('textarea', { rows: Math.max(1, row[key].split('\n').length) }));
    tr.append(
      el('td', { className: 'stt', textContent: rowNo(state.schedule, i) }),
      el('td', {}, field('date', el('input', { type: 'date' })), day),
      el('td', {}, chip('nk', 'Nhật ký'), chip('nt', 'Nghiệm thu')),
      el('td', {}, area('content')),
      el('td', {}, area('qty')),
      el('td', { className: 'rowtools' },
        el('button', { className: 'btn ghost icon', title: 'Chèn ngày bên dưới', onclick: () => insertRow(i + 1) }, icon('plus')),
        el('button', { className: 'btn ghost danger icon', title: 'Xoá dòng này', onclick: () => removeRow(i) }, icon('close'))));
    return tr;
  }));
  if (!rows.length) $('rows').append(el('tr', {}, el('td', { colSpan: 6, className: 'hint', textContent: 'Không có dòng nào khớp.' })));
}

function insertRow(at) {
  const prev = state.schedule[at - 1];
  const d = prev && prev.date ? new Date(prev.date + 'T00:00:00Z') : new Date();
  if (prev) d.setUTCDate(d.getUTCDate() + 1);
  const stt = prev ? rowNo(state.schedule, at - 1) + 1 : 1;
  shiftNo(at, 1);
  state.schedule.splice(at, 0, { stt, nt: false, nk: true, date: d.toISOString().slice(0, 10), content: '', qty: '' });
  filter = 'all';
  $('search').value = '';
  $$('#filter button').forEach(b => b.classList.toggle('on', b.dataset.f === 'all'));
  state.dirty = true;
  persist();
  tab = 'lich';
  show();
  const tr = document.querySelector(`#rows tr[data-i="${at}"]`);
  tr.scrollIntoView({ block: 'center' });
  tr.querySelector('textarea').focus();
}

function removeRow(i) {
  const row = state.schedule[i], key = editKey(row), edit = state.edits[key], no = rowNo(state.schedule, i);
  state.schedule.splice(i, 1);
  shiftNo(i, -1);
  delete state.edits[key];
  state.dirty = true;
  persist();
  if (!state.schedule.length) {
    previous = null;
    try { localStorage.removeItem(KEY); } catch { /* bỏ qua */ }
  }
  show();
  toast(`Đã xoá dòng ${no}.`, { ms: 7000, action: { label: 'Hoàn tác', run: () => {
    shiftNo(i, 1);
    state.schedule.splice(i, 0, row);
    if (edit) state.edits[key] = edit;
    persist();
    show();
  } } });
}

/* ---------- Biên bản nghiệm thu ---------- */
function renderRecords() {
  const idx = recordIdx();
  $('bbSplit').hidden = !idx.length;
  $('bbNone').hidden = !!idx.length;
  if (!idx.length) return;
  if (!idx.some(i => editKey(state.schedule[i]) === selected)) selected = editKey(state.schedule[idx[0]]);

  const items = new Map();
  const paintItem = i => {
    const row = state.schedule[i], key = editKey(row), rec = record(i);
    const title = el('b', {}, `Số ${rec.number}/NT · ${dmy(row.date)}`);
    if (suggestRecord(state.schedule, i, state).warnings.length) title.append(el('span', { className: 'pill warn' }, icon('alert'), 'kiểm tra'));
    if (state.edits[key]) title.append(el('span', { className: 'pill edit', textContent: 'đã sửa' }));
    items.get(key).replaceChildren(title, el('span', { className: 'sub', textContent: lines(rec.subjects).join(' · ') || '(chưa có đối tượng)' }));
  };
  $('recordList').replaceChildren(...idx.map(i => {
    const key = editKey(state.schedule[i]);
    const button = el('button', { className: 'item' + (key === selected ? ' on' : ''), onclick: () => { selected = key; renderRecords(); } });
    items.set(key, button);
    paintItem(i);
    return button;
  }));
  const current = idx.find(i => editKey(state.schedule[i]) === selected);
  renderEditor(current, () => paintItem(current));
}

function renderEditor(i, repaint) {
  const row = state.schedule[i], key = editKey(row);
  const rec = structuredClone(record(i));
  const warnings = suggestRecord(state.schedule, i, state).warnings;
  const reset = el('button', { className: 'btn ghost', disabled: !state.edits[key], onclick: () => {
    const old = state.edits[key];
    delete state.edits[key];
    persist();
    renderRecords();
    toast('Đã lấy lại nội dung đề xuất.', { ms: 7000, action: { label: 'Hoàn tác', run: () => { state.edits[key] = old; persist(); renderRecords(); } } });
  } }, icon('reset'), 'Đề xuất lại từ lịch');
  const touch = () => { state.edits[key] = rec; reset.disabled = false; changed(); repaint(); };

  const input = (obj, prop, type = 'text', label = '') => {
    const node = el('input', { type, value: obj[prop], oninput: e => { obj[prop] = e.target.value; touch(); } });
    if (label) node.setAttribute('aria-label', label);
    return node;
  };
  const area = (obj, prop, label) => {
    const node = el('textarea', { value: obj[prop], rows: Math.max(1, String(obj[prop]).split('\n').length), oninput: e => { obj[prop] = e.target.value; touch(); } });
    node.setAttribute('aria-label', label);
    return node;
  };
  const field = (text, node) => el('label', {}, text, node);
  const heading = (text, note = '') => el('h3', {}, text + ' ', el('small', { textContent: note }));
  // để trống = quay về số theo STT
  const numberInput = el('input', { type: 'text', value: rec.number, placeholder: suggestRecord(state.schedule, i, state).number,
    oninput: e => { rec.number = e.target.value; rec.numberEdited = !!rec.number.trim(); touch(); } });

  const tbody = el('tbody');
  const drawItems = () => tbody.replaceChildren(...rec.items.map((it, j) => el('tr', {},
    el('td', { className: 'stt', textContent: j + 1 }),
    el('td', {}, area(it, 'name', 'Nội dung công việc')),
    el('td', {}, input(it, 'unit', 'text', 'Đơn vị')),
    el('td', {}, input(it, 'qty', 'text', 'Khối lượng')),
    el('td', {}, el('button', { className: 'btn ghost danger icon', title: 'Xoá dòng', onclick: () => { rec.items.splice(j, 1); touch(); drawItems(); } }, icon('close'))))));
  drawItems();

  $('editor').replaceChildren(
    el('div', { className: 'editor-head' },
      el('h2', { textContent: `Biên bản ngày ${dmy(row.date)}` }),
      el('div', {}, reset, ' ',
        el('button', { className: 'btn primary', onclick: () => run(async () => download(await makeRecord(record(i)), `BBNT ${rec.date}.docx`)) }, icon('download'), 'Tải biên bản này'))),
    ...warnings.map(w => el('div', { className: 'notice' }, icon('alert'), el('span', { textContent: w }))),
    heading('Thông tin chung'),
    el('div', { className: 'fields' },
      field('Số biên bản (…/NT)', numberInput),
      field('Ngày lập', input(rec, 'date', 'date')),
      field('Bắt đầu kiểm tra', input(rec, 'start', 'date')),
      field('Kết thúc kiểm tra', input(rec, 'end', 'date'))),
    heading('I. Đối tượng kiểm tra', '— mỗi dòng một mục'),
    area(rec, 'subjects', 'Đối tượng kiểm tra'),
    heading('IV.2 Khối lượng khảo sát'),
    el('div', { className: 'fields' }, field('Tên nhóm trong bảng', input(rec, 'group'))),
    el('div', { className: 'scroll' }, el('table', { className: 'items' },
      el('thead', {}, el('tr', {}, ...[['TT', 36], ['Nội dung công việc', 0], ['Đơn vị', 90], ['Khối lượng', 110], ['', 40]]
        .map(([t, w]) => el('th', { textContent: t, style: w ? `width:${w}px` : '' })))),
      tbody)),
    el('button', { className: 'btn', onclick: () => { rec.items.push({ name: '', unit: '', qty: '' }); touch(); drawItems(); } }, icon('plus'), 'Thêm dòng'),
    heading('V. Kết luận', '— mỗi dòng một mục'),
    area(rec, 'conclusions', 'Kết luận'));
}

/* ---------- Nhật ký ---------- */
function renderDiaries() {
  const rows = diaryRows();
  $('nkNone').hidden = !!rows.length;
  $('diaries').replaceChildren(...rows.map(row => {
    const items = dayItems(row);
    return el('article', { className: 'card diary' },
      el('header', {}, el('b', { textContent: dmy(row.date) }), el('span', { textContent: weekday(row.date) })),
      el('ol', {}, ...(items.length ? items.map(it => {
        const li = el('li', {}, it.name + ' ');
        if (it.qty) li.append(el('span', { className: 'qty', textContent: `${it.qty} ${it.unit}`.trim() }));
        return li;
      }) : [el('li', { textContent: '(chưa có nội dung)' })])),
      el('footer', {},
        el('button', { className: 'btn ghost', onclick: () => {
          const i = state.schedule.indexOf(row);
          filter = 'all';
          $('search').value = '';
          $$('#filter button').forEach(b => b.classList.toggle('on', b.dataset.f === 'all'));
          go('lich');
          const tr = document.querySelector(`#rows tr[data-i="${i}"]`);
          tr.scrollIntoView({ block: 'center' });
          tr.querySelector('textarea').focus({ preventScroll: true });
        } }, icon('pencil'), 'Sửa'),
        el('button', { className: 'btn', onclick: () => run(async () => download(await makeDiary(row), `NK ${row.date}.docx`)) }, icon('download'), 'Tải')));
  }));
}

/* ---------- Nạp dữ liệu ---------- */
/* Chỉ hỏi khi sắp thay mất phần người dùng đã tự nhập. */
function confirmReplace() {
  const old = state.schedule.length ? state : previous;
  return !old || old.dirty === false || confirm('Dữ liệu bạn đã nhập hoặc sửa trước đó sẽ bị thay thế. Tiếp tục?');
}

function load(data) {
  // Tuỳ chọn tô vàng không lấy theo dữ liệu đã lưu: mỗi lần mở trang đều mặc định bỏ.
  state = { ...blank(), ...data, noHighlight: state.noHighlight };
  selected = null;
  tab = 'lich';
  filter = 'all';
  $('search').value = '';
  $$('#filter button').forEach(b => b.classList.toggle('on', b.dataset.f === 'all'));
  persist();
  show();
  window.scrollTo(0, 0);
}

/* Thông tin chung đã sửa đi theo công trình, không mất khi nhập lịch Excel khác. */
const keptInfo = () => structuredClone((state.schedule.length ? state.info : previous?.info) || {});

async function loadSchedule(buf, what, info) {
  const schedule = scheduleFromSheet(await readXlsx(buf));
  load({ schedule, info });
  toast(`Đã ${what} ${schedule.length} dòng.`);
}

async function openFile(file) {
  try {
    const name = file.name.toLowerCase();
    if (name.endsWith('.xlsx')) {
      if (confirmReplace()) await loadSchedule(await file.arrayBuffer(), 'nhập', keptInfo());
    } else if (name.endsWith('.json')) {
      const data = JSON.parse(await file.text());
      if (!Array.isArray(data.schedule)) throw new Error('File không đúng định dạng dữ liệu');
      if (confirmReplace()) { load({ ...data, dirty: true }); toast('Đã mở dữ liệu.'); }
    } else throw new Error('Chỉ nhận file .xlsx hoặc .json');
  } catch (err) { console.error(err); toast('Không đọc được file: ' + err.message, { error: true }); }
}

async function loadSample() {
  if (!confirmReplace()) return;
  try { await loadSchedule(b64(ASSETS.sample), 'tải dữ liệu mẫu,', {}); }
  catch (err) { console.error(err); toast('Không tải được dữ liệu mẫu: ' + err.message, { error: true }); }
}

/* ---------- File mẫu Word ---------- */
function paintTemplates() {
  for (const [kind, id] of [['nk', 'tplNkName'], ['bb', 'tplBbName']]) {
    $(id).textContent = tplNames[kind] || 'Mẫu gốc';
    document.querySelector(`[data-reset="${kind}"]`).hidden = !tplNames[kind];
  }
}

async function useTemplate(kind, file) {
  const fill = kind === 'nk' ? fillDiaryXml : fillRecordXml;
  const sample = kind === 'nk' ? diaryRows()[0] : recordIdx().map(record)[0];
  try {
    const buf = await file.arrayBuffer();
    if (sample) await buildDoc(buf, fill, sample, state);   // thử trước để báo lỗi mẫu ngay
    templates[kind] = buf;
    tplNames[kind] = file.name;
    paintTemplates();
    await loadDefaults();
    toast(`Đang dùng mẫu "${file.name}".`);
  } catch (err) { console.error(err); toast('Không dùng được mẫu: ' + err.message, { error: true }); }
}

/* ---------- Sự kiện ---------- */
const onFile = (id, fn) => { $(id).onchange = e => {
  const file = e.target.files[0];
  e.target.value = '';
  if (file) fn(file);
}; };

['inXlsx', 'inJson', 'emptyXlsx', 'emptyJson'].forEach(id => onFile(id, openFile));
onFile('inTplNk', f => useTemplate('nk', f));
onFile('inTplBb', f => useTemplate('bb', f));
$$('[data-reset]').forEach(b => { b.onclick = () => {
  const kind = b.dataset.reset;
  templates[kind] = BUILTIN[kind];
  tplNames[kind] = null;
  paintTemplates();
  loadDefaults();
}; });

$('loadSample').onclick = loadSample;
$('emptySample').onclick = loadSample;
$('emptyManual').onclick = () => {
  if (!confirmReplace()) return;
  state = { ...blank(), noHighlight: state.noHighlight, info: keptInfo() };
  previous = null;
  insertRow(0);
};
$('restore').onclick = () => { load(previous); toast('Đã mở lại phiên trước.'); };
$('saveJson').onclick = () => {
  download(new Blob([JSON.stringify(state, null, 1)], { type: 'application/json' }), 'du-lieu-khao-sat.json');
};
$('clearAll').onclick = () => {
  if (!confirm('Xoá toàn bộ dữ liệu đang nhập? Bấm "Lưu" trước nếu muốn giữ lại.')) return;
  state = { ...blank(), noHighlight: state.noHighlight };
  previous = null;
  try { localStorage.removeItem(KEY); } catch { /* bỏ qua */ }
  show();
};

$('resetInfo').onclick = () => {
  const old = state.info;
  if (!Object.keys(old).length) return toast('Thông tin đang đúng theo mẫu.');
  state.info = {};
  persist();
  renderInfo();
  toast('Đã khôi phục theo mẫu.', { ms: 7000, action: { label: 'Hoàn tác', run: () => { state.info = old; persist(); renderInfo(); } } });
};
$$('.tab').forEach(b => { b.onclick = () => go(b.dataset.tab); });
$$('[data-export]').forEach(b => { b.onclick = () => exporters[b.dataset.export](); });
$$('[data-add]').forEach(b => { b.onclick = () => insertRow(state.schedule.length); });
$$('#filter button').forEach(b => { b.onclick = () => {
  filter = b.dataset.f;
  $$('#filter button').forEach(x => x.classList.toggle('on', x === b));
  renderSchedule();
  window.scrollTo(0, 0);
}; });
$('search').oninput = () => { renderSchedule(); window.scrollTo(0, 0); };
$('noHl').onchange = e => { state.noHighlight = e.target.checked; persist(); };
$('convert').onchange = e => { state.convert = e.target.checked; persist(); renderRecords(); };

/* Kéo thả file vào bất kỳ chỗ nào trên trang. */
const hasFiles = e => [...(e.dataTransfer?.types || [])].includes('Files');
document.addEventListener('dragover', e => { if (hasFiles(e)) { e.preventDefault(); document.body.classList.add('dragging'); } });
document.addEventListener('dragleave', e => { if (!e.relatedTarget) document.body.classList.remove('dragging'); });
document.addEventListener('drop', e => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  document.body.classList.remove('dragging');
  const file = e.dataTransfer.files[0];
  if (file) openFile(file);
});

/* Mở lên luôn trống; phiên tự lưu chỉ được mời mở lại. */
try {
  const saved = JSON.parse(localStorage.getItem(KEY));
  if (saved?.schedule?.length) previous = saved;
} catch { /* không đọc được bộ nhớ trình duyệt */ }
paintTemplates();
const watcher = new ResizeObserver(syncOffsets);
$$('body > header, .toolbar').forEach(n => watcher.observe(n));
window.addEventListener('resize', syncOffsets);
show();
loadDefaults();
