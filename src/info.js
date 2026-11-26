/* Thông tin chung (tên công trình, công ty, người ký, hợp đồng…): đọc từ mẫu và thế vào file xuất.
   Mỗi trường neo theo nhãn đứng trước, chỉ thay phần giá trị nên định dạng được giữ.
   Trường người dùng không sửa thì không đụng tới. */
'use strict';

/* dot: mẫu có sẵn dấu chấm cuối câu. Các trường cùng một đoạn xếp từ cuối đoạn lên đầu,
   để thay trường sau không làm lệch mốc của trường trước. */
const INFO_TEXT = {
  nk: [
    { key: 'phase', re: /Bước lập\s+(.*?)\s*\)/ },
    { key: 'projectName', re: /Công trình:\s*(.*?)\.?\s*$/, dot: true },
    { key: 'location', re: /Địa điểm xây dựng:\s*(.*?)\.?\s*$/, dot: true },
    { key: 'weather', re: /Thời tiết:\s*(.*?)\s*$/ },
    { key: 'manpower', re: /Nhân lực:\s*(.*?)\.?\s*$/, dot: true },
    { key: 'equipment', re: /Thiết bị:\s*(.*?)\.?\s*$/, dot: true },
    { key: 'remark', re: /NHẬN XÉT:\s*(.*?)\s*$/ },
  ],
  bb: [
    { key: 'place', re: /^\s*([^,]+?)\s*,\s*ngày\s/ },
    { key: 'supervisorCompany', re: /Nhà thầu giám sát khảo sát\s*:\s*(.*?)\.?\s*$/, dot: true },
    { key: 'surveyCompany', re: /Nhà thầu khảo sát\s*:\s*(.*?)\.?\s*$/, dot: true },
    { key: 'projectName', re: /thuộc dự án\s+(.*?)\.?\s*$/, dot: true },
    { key: 'contractAbout', re: /về việc\s+(.*?)\s+thuộc dự án\s/ },
    { key: 'contractor', re: /ký kết giữa\s+.*?\s+và\s+(.*?)\s+về việc\s/ },
    { key: 'investor', re: /ký kết giữa\s+(.*?)\s+và\s+/ },
    { key: 'contractDate', re: /Hợp đồng số\s+.*?\s+ngày\s+(\S+)/ },
    { key: 'contractNo', re: /Hợp đồng số\s+(.*?)\s+ngày\s/ },
    { key: 'projectName', re: /DỰ ÁN\s*:\s*(.*?)\s*$/, upper: true },
  ],
};
const INFO_MEMBERS = [
  { key: 'supervisors', label: /Nhà thầu giám sát khảo sát\s*:/ },
  { key: 'surveyors', label: /Nhà thầu khảo sát\s*:/ },
];

const tNodes = el => [...el.getElementsByTagNameNS(W, 't')];
const matchPara = (body, re) => kids(body, 'p').find(p => re.test(textOf(p)));

/* Bảng danh sách người nằm ngay sau đoạn có nhãn. */
function memberTable(body, label) {
  let n = matchPara(body, label)?.nextElementSibling;
  while (n && n.localName === 'p' && !textOf(n).trim()) n = n.nextElementSibling;
  return n && n.localName === 'tbl' ? n : null;
}

/* Bảng chữ ký là bảng cuối; tên người ký là đoạn có chữ cuối cùng trong mỗi ô. */
function signatureParas(body) {
  const table = kids(body, 'tbl').pop();
  const row = table && kids(table, 'tr')[0];
  const out = {};
  for (const tc of row ? kids(row, 'tc') : []) {
    const ps = kids(tc, 'p').filter(p => textOf(p).trim());
    if (ps.length < 2) continue;
    const role = ps.slice(0, -1).map(textOf).join(' ');
    out[/GSKS|giám sát/i.test(role) ? 'signSupervisor' : 'signSurveyor'] = ps[ps.length - 1];
  }
  return out;
}

const issueCells = body => {
  const table = kids(body, 'tbl').find(t => textOf(t).includes('Các vấn đề vướng mắc'));
  const row = table && kids(table, 'tr')[1];
  return row ? kids(row, 'tc') : [];
};

function readInfoXml(doc, kind, out) {
  const body = doc.getElementsByTagNameNS(W, 'body')[0];
  const put = (key, value) => { if (out[key] === undefined) out[key] = value; };
  for (const f of INFO_TEXT[kind]) {
    if (f.upper) continue;                       // tiêu đề in hoa, lấy giá trị ở chỗ viết thường
    const p = matchPara(body, f.re);
    if (p) put(f.key, f.re.exec(textOf(p))[1]);
  }
  for (const [key, p] of Object.entries(signatureParas(body))) put(key, textOf(p).trim());
  if (kind === 'nk') {
    const [issue, solution] = issueCells(body);
    if (issue) put('issues', textOf(issue));
    if (solution) put('solution', textOf(solution));
  } else {
    for (const m of INFO_MEMBERS) {
      const table = memberTable(body, m.label);
      if (!table) continue;
      put(m.key, kids(table, 'tr').map(tr => {
        const [who, role] = kids(tr, 'tc').map(tc => textOf(tc).trim());
        const found = /^-?\s*(Ông|Bà)\s*:\s*(.*)$/.exec(who || '');
        return {
          title: found ? found[1] : 'Ông',
          name: found ? found[2] : who || '',
          role: (role || '').replace(/^Chức vụ\s*:\s*/, ''),
        };
      }));
    }
  }
}

async function readDefaults(templates) {
  const out = {};
  for (const kind of ['nk', 'bb']) {
    const files = await unzip(templates[kind]);
    readInfoXml(parseXml(files.get('word/document.xml')), kind, out);
  }
  return out;
}

function applyInfo(doc, kind, info) {
  const body = doc.getElementsByTagNameNS(W, 'body')[0];
  const has = key => info[key] !== undefined;
  const oneLine = v => String(v).replace(/\s*\n\s*/g, ' ').trim();

  for (const f of INFO_TEXT[kind]) {
    if (!has(f.key)) continue;
    const p = matchPara(body, f.re);
    if (!p) continue;
    let value = oneLine(info[f.key]);
    if (f.dot) value = value.replace(/\s*\.+$/, '');
    if (f.upper) value = value.toLocaleUpperCase('vi');
    replaceAcross(tNodes(p), f.re, value);
  }
  for (const [key, p] of Object.entries(signatureParas(body)))
    if (has(key)) replaceAcross(tNodes(p), /^\s*(.*?)\s*$/, oneLine(info[key]));

  if (kind === 'nk') {
    const cells = issueCells(body);
    const style = cells[0]?.getElementsByTagNameNS(W, 'rPr');      // kiểu chữ của ô "vướng mắc"
    const model = style && [...style].find(x => x.parentNode.localName === 'r');
    cells.forEach((tc, i) => {
      const key = ['issues', 'solution'][i];
      if (!key || !has(key)) return;
      const wasEmpty = !tc.getElementsByTagNameNS(W, 'r').length;
      fillCell(tc, oneLine(info[key]));
      const run = tc.getElementsByTagNameNS(W, 'r')[0];
      // Ô trống trong mẫu mang định dạng riêng (chữ xanh): dùng kiểu của ô bên cạnh.
      if (wasEmpty && run && model) {
        const own = kid(run, 'rPr');
        if (own) run.removeChild(own);
        run.insertBefore(model.cloneNode(true), run.firstChild);
      }
    });
    return;
  }
  for (const m of INFO_MEMBERS) {
    if (!has(m.key)) continue;
    const table = memberTable(body, m.label);
    if (!table) continue;
    const rows = kids(table, 'tr');
    if (!info[m.key].length) { table.parentNode.removeChild(table); continue; }
    const proto = rows[0].cloneNode(true);
    rows.forEach(r => table.removeChild(r));
    for (const person of info[m.key]) {
      const tr = proto.cloneNode(true);
      const [who, role] = kids(tr, 'tc');
      replaceAcross(tNodes(who), /:\s*(.*?)\s*$/, oneLine(person.name));
      replaceAcross(tNodes(who), /(Ông|Bà)\s*:/, person.title);
      if (role) replaceAcross(tNodes(role), /Chức vụ\s*:\s*(.*?)\s*$/, oneLine(person.role));
      table.appendChild(tr);
    }
  }
}
