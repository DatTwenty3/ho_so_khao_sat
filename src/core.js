/* Lõi xử lý: đọc/ghi ZIP, đọc xlsx, suy ra hồ sơ từ lịch, thế dữ liệu vào chỗ tô vàng của mẫu docx. */
'use strict';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const enc = new TextEncoder(), dec = new TextDecoder();

/* ---------- ZIP ---------- */
async function unzip(buf) {
  const u8 = new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let e = u8.length - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error('Không phải file .docx/.xlsx hợp lệ');
  const count = dv.getUint16(e + 10, true);
  let p = dv.getUint32(e + 16, true);
  const files = new Map();
  for (let i = 0; i < count; i++) {
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    const off = dv.getUint32(p + 42, true);
    const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
    const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
    let data = u8.subarray(start, start + csize);
    if (method === 8) {
      const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      data = new Uint8Array(await new Response(stream).arrayBuffer());
    } else if (method !== 0) throw new Error('Kiểu nén ZIP không hỗ trợ');
    files.set(name, data);
    p += 46 + nlen + xlen + clen;
  }
  return files;
}

const CRC = Uint32Array.from({ length: 256 }, (_, n) => {
  for (let k = 0; k < 8; k++) n = n & 1 ? 0xedb88320 ^ (n >>> 1) : n >>> 1;
  return n >>> 0;
});
function crc32(d) {
  let c = 0xffffffff;
  for (let i = 0; i < d.length; i++) c = CRC[(c ^ d[i]) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/* Ghi ZIP không nén: file nhỏ, Word đọc bình thường. */
function zip(files) {
  const local = [], central = [];
  let offset = 0, csize = 0;
  const now = new Date();
  const time = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  for (const [name, data] of files) {
    const n = enc.encode(name), crc = crc32(data);
    const head = (sig, len, extra) => {
      const b = new Uint8Array(len), v = new DataView(b.buffer);
      v.setUint32(0, sig, true);
      let o = 4;
      if (extra) { v.setUint16(o, 20, true); o += 2; }
      v.setUint16(o, 20, true); v.setUint16(o + 2, 0x800, true); v.setUint16(o + 4, 0, true);
      v.setUint16(o + 6, time, true); v.setUint16(o + 8, date, true);
      v.setUint32(o + 10, crc, true); v.setUint32(o + 14, data.length, true); v.setUint32(o + 18, data.length, true);
      v.setUint16(o + 22, n.length, true);
      return [b, v];
    };
    const [lh] = head(0x04034b50, 30, false);
    const [ch, cv] = head(0x02014b50, 46, true);
    cv.setUint32(42, offset, true);
    local.push(lh, n, data);
    central.push(ch, n);
    offset += 30 + n.length + data.length;
    csize += 46 + n.length;
  }
  const end = new Uint8Array(22), ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, files.size, true); ev.setUint16(10, files.size, true);
  ev.setUint32(12, csize, true); ev.setUint32(16, offset, true);
  return new Blob([...local, ...central, end]);
}

const parseXml = bytes => new DOMParser().parseFromString(dec.decode(bytes), 'application/xml');
function serializeXml(doc) {
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return enc.encode(s);
}

/* ---------- Đọc Excel ---------- */
async function readXlsx(buf) {
  const files = await unzip(buf);
  const all = (node, tag) => [...node.getElementsByTagNameNS('*', tag)];
  const shared = files.has('xl/sharedStrings.xml')
    ? all(parseXml(files.get('xl/sharedStrings.xml')), 'si').map(si =>
        all(si, 't').filter(t => t.parentNode.localName !== 'rPh').map(t => t.textContent).join(''))
    : [];
  const wb = parseXml(files.get('xl/workbook.xml'));
  const rid = all(wb, 'sheet')[0].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
  const rel = all(parseXml(files.get('xl/_rels/workbook.xml.rels')), 'Relationship').find(r => r.getAttribute('Id') === rid);
  const path = rel.getAttribute('Target').replace(/^\/?(xl\/)?/, 'xl/');
  const rows = [];
  for (const row of all(parseXml(files.get(path)), 'row')) {
    const cells = {};
    for (const c of all(row, 'c')) {
      const col = c.getAttribute('r').replace(/\d+/g, ''), t = c.getAttribute('t');
      const v = all(c, 'v')[0];
      let val = '';
      if (t === 'inlineStr') val = all(c, 't').map(x => x.textContent).join('');
      else if (v) val = t === 's' ? shared[+v.textContent] : (t === 'str' ? v.textContent : Number(v.textContent));
      cells[col] = val;
    }
    rows.push(cells);
  }
  return rows;
}

function scheduleFromSheet(rows) {
  const norm = s => String(s ?? '').normalize('NFC').toUpperCase().trim();
  const hi = rows.findIndex(r => Object.values(r).some(v => norm(v) === 'N.THU'));
  if (hi < 0) throw new Error('Không thấy dòng tiêu đề có ô "N.THU"');
  const col = test => Object.keys(rows[hi]).find(k => test(norm(rows[hi][k])));
  const c = {
    nt: col(h => h === 'N.THU'), nk: col(h => h === 'N.KÝ'), day: col(h => h === 'NGÀY'),
    month: col(h => h === 'THÁNG'), content: col(h => h.startsWith('ĐỐI TƯỢNG')), qty: col(h => h.startsWith('KL')),
  };
  for (const [k, v] of Object.entries(c)) if (!v) throw new Error('Thiếu cột trong Excel: ' + k);
  const sttCol = col(h => h === 'STT');        // không bắt buộc: thiếu thì đánh số theo thứ tự dòng
  const above = rows.slice(0, hi + 1).flatMap(r => Object.values(r));
  let year = above.find(v => typeof v === 'number' && Number.isInteger(v) && v >= 2000 && v <= 2100) || new Date().getFullYear();
  const out = [];
  let prevMonth = 0;
  for (const r of rows.slice(hi + 1)) {
    const d = r[c.day], m = r[c.month];
    if (!d || !m) continue;
    if (prevMonth && m < prevMonth) year++;
    prevMonth = m;
    const flag = v => String(v ?? '').trim() !== '';
    const stt = parseInt(r[sttCol], 10);
    out.push({
      ...(stt > 0 && { stt }),
      nt: flag(r[c.nt]), nk: flag(r[c.nk]),
      date: `${year}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      content: String(r[c.content] ?? '').replace(/\r/g, '').trim(),
      qty: String(r[c.qty] ?? '').replace(/\r/g, '').trim(),
    });
  }
  if (!out.length) throw new Error('Không đọc được dòng dữ liệu nào');
  return out;
}

/* STT của dòng: lấy theo cột STT trong Excel, không có thì theo vị trí. */
const rowNo = (schedule, i) => schedule[i].stt ?? i + 1;

/* ---------- Suy ra nội dung hồ sơ ---------- */
const lines = s => String(s || '').split('\n').map(x => x.replace(/^\s*[-+•]\s*/, '').trim()).filter(Boolean);
const isAccept = l => /^nghiệm thu\s+/i.test(l.normalize('NFC'));
const baseName = l => l.normalize('NFC').replace(/^nghiệm thu\s+/i, '').split(':')[0].trim().replace(/[.;,]+$/, '');
const keyOf = l => baseName(l).toLowerCase().replace(/thuỷ/g, 'thủy').replace(/\s+/g, ' ');
const cap = s => s.charAt(0).toUpperCase() + s.slice(1);

function parseNum(s) {
  s = s.trim();
  if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}
function splitQty(s) {
  const m = /^\s*([\d.,]+)\s*(.*)$/.exec(s || '');
  return m ? { qty: m[1], unit: m[2].trim() } : { qty: s || '', unit: '' };
}
/* Đơn vị định mức (100m, 100ha) giữ ít nhất 2 số lẻ như mẫu; còn lại bỏ số 0 thừa. */
const fmtNum = (n, unit) => {
  const s = (Math.round(n * 1e4) / 1e4).toFixed(4);
  return (unit.startsWith('100') ? s.replace(/0{1,2}$/, '') : s.replace(/\.?0+$/, '')).replace('.', ',');
};

/* Ghép từng dòng khối lượng với dòng công việc tương ứng. */
function dayItems(row) {
  const names = lines(row.content), qtys = lines(row.qty);
  const items = names.map(name => ({ name, unit: '', qty: '' }));
  let targets = items;
  if (qtys.length !== items.length) {
    const work = items.filter(i => !isAccept(i.name));
    if (work.length) targets = work;
  }
  qtys.forEach((q, i) => {
    const t = targets[Math.min(i, targets.length - 1)];
    if (!t) return;
    const p = splitQty(q);
    if (t.qty) { t.qty += '; ' + p.qty; t.unit += '; ' + p.unit; } else Object.assign(t, p);
  });
  return items;
}

/* Cộng dồn khối lượng của các công việc thoả điều kiện, tách theo đơn vị. */
function totals(rows, accept, convert) {
  const map = new Map();
  for (const row of rows) for (const it of dayItems(row)) {
    if (isAccept(it.name)) continue;
    const k = keyOf(it.name), n = parseNum(it.qty);
    if (!accept(k) || n === null || !it.qty) continue;
    const id = k + '|' + it.unit;
    if (!map.has(id)) map.set(id, { name: cap(baseName(it.name)), unit: it.unit, sum: 0 });
    map.get(id).sum += n;
  }
  return [...map.values()].map(t => {
    let { unit, sum } = t;
    if (convert && (unit === 'm' || unit === 'ha')) { sum /= 100; unit = '100' + unit; }
    return { name: t.name, unit, qty: fmtNum(sum, unit) };
  });
}

/* Nội dung đề xuất cho biên bản nghiệm thu của dòng thứ idx. */
function suggestRecord(schedule, idx, opt) {
  const rows = [...schedule];
  const row = rows[idx];
  const accepted = lines(row.content).filter(isAccept);
  const subjects = (accepted.length ? accepted : lines(row.content)).map(l => cap(baseName(l)));
  const warnings = [];
  let items = [];
  if (!accepted.length) {
    items = dayItems(row).map(i => ({ name: cap(baseName(i.name)), unit: i.unit, qty: i.qty }));
  } else {
    for (const line of accepted) {
      const k = keyOf(line);
      if (k.startsWith('hoàn thành')) { items.push(...totals(rows.slice(0, idx + 1), () => true, opt.convert)); continue; }
      let from = 0;
      for (let i = 0; i < idx; i++)
        if (lines(rows[i].content).some(l => isAccept(l) && keyOf(l) === k)) from = i + 1;
      if (from > 0) warnings.push(`"${cap(baseName(line))}" đã được nghiệm thu một lần trước đó — kiểm tra lại lịch.`);
      const found = totals(rows.slice(from, idx + 1), x => x === k, opt.convert);
      if (found.length) items.push(...found);
      else items.push({ name: cap(baseName(line)), unit: '', qty: '' });
      if (found.length > 1) warnings.push(`"${cap(baseName(line))}" có nhiều đơn vị khác nhau, không cộng gộp được.`);
    }
  }
  return {
    number: String(rowNo(rows, idx)).padStart(2, '0'), numberEdited: false,
    date: row.date, start: row.date, end: row.date,
    subjects: subjects.join('\n'), group: 'Khảo sát phần tuyến', items,
    conclusions: (items.length ? [...new Set(items.map(i => i.name))] : subjects).join('\n'),
    warnings,
  };
}

/* ---------- Thao tác trên mẫu docx ---------- */
const kids = (el, tag) => [...el.children].filter(c => c.namespaceURI === W && c.localName === tag);
const kid = (el, tag) => el && kids(el, tag)[0];
const textOf = el => [...el.getElementsByTagNameNS(W, 't')].map(t => t.textContent).join('').normalize('NFC');
const isHl = r => !!kid(kid(r, 'rPr'), 'highlight');
const hlRuns = p => [...p.getElementsByTagNameNS(W, 'r')].filter(isHl);

function setRunText(r, text) {
  const doc = r.ownerDocument;
  [...r.children].forEach(c => { if (c.localName !== 'rPr') r.removeChild(c); });
  String(text).split('\n').forEach((part, i) => {
    if (i) r.appendChild(doc.createElementNS(W, 'w:br'));
    const t = doc.createElementNS(W, 'w:t');
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = part;
    r.appendChild(t);
  });
}

/* Thế text vào các run tô vàng của đoạn; phần không tô vàng giữ nguyên. */
function fillPara(p, text) {
  let runs = hlRuns(p);
  if (!runs.length) runs = [...p.getElementsByTagNameNS(W, 'r')];
  if (!runs.length) {
    if (!text) return;
    const r = p.ownerDocument.createElementNS(W, 'w:r');
    const mark = kid(kid(p, 'pPr'), 'rPr');
    if (mark) r.appendChild(mark.cloneNode(true));
    p.appendChild(r);
    runs = [r];
  }
  setRunText(runs[0], text);
  runs.slice(1).forEach(r => r.parentNode.removeChild(r));
}

function fillCell(tc, text) {
  const ps = kids(tc, 'p');
  const main = ps.find(p => hlRuns(p).length) || ps[0];
  ps.forEach(p => { if (p !== main) tc.removeChild(p); });
  fillPara(main, text);
}

/* Nhân bản đoạn mẫu cho đủ số dòng. */
function fillBlock(samples, texts) {
  const proto = samples[0].cloneNode(true), parent = samples[0].parentNode, anchor = samples[0];
  (texts.length ? texts : ['']).forEach(t => {
    const p = proto.cloneNode(true);
    parent.insertBefore(p, anchor);
    fillPara(p, t);
  });
  samples.forEach(p => parent.removeChild(p));
}

function fillRows(table, protoRow, oldRows, data) {
  const proto = protoRow.cloneNode(true);
  oldRows.forEach(r => table.removeChild(r));
  data.forEach(values => {
    const tr = proto.cloneNode(true);
    kids(tr, 'tc').forEach((tc, i) => fillCell(tc, values[i] ?? ''));
    table.appendChild(tr);
  });
}

const need = (x, what) => {
  if (!x || (Array.isArray(x) && !x.length)) throw new Error('File mẫu thiếu chỗ tô vàng: ' + what);
  return x;
};

const dmy = iso => iso.split('-').reverse().join('/');
const longDate = iso => { const [, m, d] = iso.split('-'); return `ngày ${d} tháng ${m} năm`; };

function fillDiaryXml(doc, row) {
  const body = doc.getElementsByTagNameNS(W, 'body')[0];
  const time = need(kids(body, 'p').find(p => textOf(p).includes('Thời gian') && hlRuns(p).length), 'Thời gian');
  fillPara(time, dmy(row.date));
  const table = need(kids(body, 'tbl').find(t => textOf(t).includes('Nội dung công việc')), 'bảng khối lượng');
  const data = need(kids(table, 'tr').slice(1), 'dòng mẫu của bảng');
  const items = dayItems(row);
  fillRows(table, data[data.length - 1], data,
    (items.length ? items : [{ name: '', unit: '', qty: '' }]).map((it, i) => [String(i + 1), it.name, it.unit, it.qty]));
}

/* Thay nhóm 1 của regex bằng value, kể cả khi chuỗi bị Word cắt ra nhiều run. */
function replaceAcross(ts, regex, value) {
  const full = ts.map(t => t.textContent).join('');
  const m = new RegExp(regex.source, regex.flags + 'd').exec(full);
  if (!m || m[1] === value) return;
  const [from, to] = m.indices[1];
  let pos = 0, done = false;
  for (const t of ts) {
    const text = t.textContent, a = Math.max(from - pos, 0), b = Math.min(to - pos, text.length);
    pos += text.length;
    if (b > a) {
      t.textContent = text.slice(0, a) + (done ? '' : value) + text.slice(b);
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      done = true;
    }
  }
}

/* Năm nằm ngoài vùng tô vàng: chỉ sửa khi khác năm trong mẫu. */
function fixYear(p, iso) {
  const runs = [...p.getElementsByTagNameNS(W, 'r')];
  const last = runs.map(isHl).lastIndexOf(true);
  const ts = runs.slice(last + 1).flatMap(r => kids(r, 't'));
  replaceAcross(ts, /\b(20\d\d)\b/, iso.slice(0, 4));
}

const setNumber = (p, number) => replaceAcross([...p.getElementsByTagNameNS(W, 't')], /(\d+)\s*\/\s*NT/, number);

function fillRecordXml(doc, rec) {
  const body = doc.getElementsByTagNameNS(W, 'body')[0];
  const ps = kids(body, 'p');
  const table = need(kids(body, 'tbl').find(t => textOf(t).includes('Nội dung công việc')), 'bảng khối lượng');
  const order = [...body.children];
  const dateP = need(ps.find(p => hlRuns(p).length), 'ngày lập biên bản');
  const head = need(ps.find(p => textOf(p).includes('Đối tượng kiểm tra')), 'Đối tượng kiểm tra');
  const subjects = [];
  for (let n = head.nextElementSibling; n && n.localName === 'p' && hlRuns(n).length; n = n.nextElementSibling) subjects.push(n);
  const start = need(ps.find(p => textOf(p).includes('Bắt đầu') && hlRuns(p).length), 'Bắt đầu');
  const end = need(ps.find(p => textOf(p).includes('Kết thúc') && hlRuns(p).length), 'Kết thúc');
  const concl = ps.filter(p => order.indexOf(p) > order.indexOf(table) && hlRuns(p).length);
  const num = ps.find(p => textOf(p).includes('BIÊN BẢN SỐ'));

  for (const [p, iso] of [[dateP, rec.date], [start, rec.start], [end, rec.end]]) {
    fillPara(p, longDate(iso));
    fixYear(p, iso);
  }
  if (num && rec.number) setNumber(num, rec.number.trim());
  fillBlock(need(subjects, 'dòng đối tượng kiểm tra'), lines(rec.subjects).map(s => '- ' + s));
  fillBlock(need(concl, 'phần kết luận'), lines(rec.conclusions).map(s => '+ ' + s.replace(/[.;]+$/, '') + '.'));

  const trs = kids(table, 'tr');
  const group = need(trs[1], 'dòng nhóm'), proto = need(trs[2], 'dòng công việc').cloneNode(true);
  kids(group, 'tc').forEach((tc, i) => { if (i === 1) fillCell(tc, rec.group); });
  if (!rec.group.trim()) table.removeChild(group);
  fillRows(table, proto, trs.slice(2), rec.items.map((it, i) => [String(i + 1), it.name, it.unit, it.qty]));
}

function stripHighlight(doc) {
  [...doc.getElementsByTagNameNS(W, 'highlight')].forEach(h => h.parentNode.removeChild(h));
}

/* Thế nội dung hồ sơ, rồi thông tin chung người dùng đã sửa. */
function render(doc, fill, data, opt) {
  fill(doc, data);
  if (opt.info) applyInfo(doc, fill === fillDiaryXml ? 'nk' : 'bb', opt.info);
  if (opt.noHighlight) stripHighlight(doc);
}

async function buildDoc(template, fill, data, opt) {
  const files = await unzip(template);
  const doc = parseXml(files.get('word/document.xml'));
  render(doc, fill, data, opt);
  files.set('word/document.xml', serializeXml(doc));
  return zip(files);
}

/* Gộp nhiều ngày nhật ký vào một file, mỗi ngày bắt đầu trang mới. */
async function buildMerged(template, fill, list, opt) {
  const files = await unzip(template);
  let main = null, body = null, sect = null;
  for (const [i, data] of list.entries()) {
    const doc = parseXml(files.get('word/document.xml'));
    render(doc, fill, data, opt);
    const b = doc.getElementsByTagNameNS(W, 'body')[0];
    if (!i) { main = doc; body = b; sect = kid(b, 'sectPr'); continue; }
    [...b.children].filter(n => n.localName !== 'sectPr').forEach((n, j) => {
      const node = main.importNode(n, true);
      if (!j && node.localName === 'p') {
        let pPr = kid(node, 'pPr');
        if (!pPr) pPr = node.insertBefore(main.createElementNS(W, 'w:pPr'), node.firstChild);
        const br = main.createElementNS(W, 'w:pageBreakBefore');
        const style = kid(pPr, 'pStyle');
        pPr.insertBefore(br, style ? style.nextSibling : pPr.firstChild);
      }
      body.insertBefore(node, sect);
    });
  }
  // Mẫu ẩn chân trang ở trang đầu (titlePg): giữ như vậy cho mọi ngày trong file gộp.
  if (sect && kid(sect, 'titlePg')) kids(sect, 'footerReference')
    .filter(f => f.getAttributeNS(W, 'type') === 'default').forEach(f => sect.removeChild(f));
  files.set('word/document.xml', serializeXml(main));
  return zip(files);
}
