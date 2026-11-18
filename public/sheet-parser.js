// อ่านชีตรูปแบบ "Tax Year" (เหมือนชีต 2025/2026) แล้วแปลงเป็นข้อมูลสำหรับ Firestore
export const TH = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
const SAVING = ['เงินออม','เก็บเงิน','เงินเก็บ','เงินสำรอง','โบนัส'];
const CASH = ['เงินมาทำงาน','เงินเก็บห้ามใช้'];

const clean = v => (typeof v === 'string' ? v.replace(/\*/g, '').trim() : v);
const n = v => (typeof v === 'number' && isFinite(v) ? v : (typeof v === 'string' && v.trim() !== '' && isFinite(+v) ? +v : 0));
function lev(a, b) { const d = Array.from({ length: a.length + 1 }, (_, i) => [i]); for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); return d[a.length][b.length]; }
// รองรับชื่อเดือนที่พิมพ์ผิดเล็กน้อย เช่น "มิถุยายน"
const monthOf = v => { const t = clean(v); if (typeof t !== 'string' || t.length < 4) return 0; const i = TH.indexOf(t); if (i >= 0) return i + 1;
  const near = TH.map((m, k) => [lev(m, t), k]).sort((a, b) => a[0] - b[0])[0]; return near[0] <= 1 ? near[1] + 1 : 0; };
const norm = s => s.replace(/็+/g, '็').replace(/\s+/g, ' ').trim();
const pad = x => String(x).padStart(2, '0');

function toISODate(XLSX, v) {
  if (v instanceof Date) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
  if (typeof v === 'number') { const d = XLSX.SSF.parse_date_code(v); if (d) return `${d.y}-${pad(d.m)}-${pad(d.d)}`; }
  return '';
}

export function parseWorkbook(XLSX, wb) {
  return wb.SheetNames.map(name => {
    try { return { sheet: name, ...parseSheet(XLSX, wb.Sheets[name], name) }; }
    catch (e) { return { sheet: name, ok: false, reason: e.message }; }
  });
}

export function parseSheet(XLSX, ws, name) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  // sheet_to_json เริ่มจากเซลล์แรกที่มีข้อมูล ปรับให้ตรงกับตำแหน่งจริงของชีต
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const R = (r, c) => { const row = rows[r - range.s.r]; return row ? row[c - range.s.c] ?? null : null; };
  const maxR = range.e.r, maxC = range.e.c;
  const find = test => { for (let r = range.s.r; r <= maxR; r++) for (let c = range.s.c; c <= maxC; c++) if (test(R(r, c), r, c)) return [r, c]; return null; };

  // ปี
  let year = null;
  const yc = find(v => typeof v === 'string' && /Tax Year\s*\d{4}/i.test(v));
  if (yc) year = +R(...yc).match(/\d{4}/)[0];
  if (!year && /^\d{4}$/.test(String(name).trim())) year = +String(name).trim();

  // อ่านสูตรของเซลล์ แล้วคืนชุดตำแหน่งเซลล์ที่สูตรอ้างถึง
  const F = (r, c) => { const cell = ws[XLSX.utils.encode_cell({ r, c })]; return cell && cell.f ? String(cell.f) : null; };
  const refs = f => { const set = new Set(); if (!f || f.includes('!')) return set;
    for (const m of f.matchAll(/\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g)) {
      const a = XLSX.utils.decode_cell(m[1] + m[2]), b = m[3] ? XLSX.utils.decode_cell(m[3] + m[4]) : a;
      for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r++) for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c++) set.add(r + ',' + c);
    } return set; };
  const noteText = v => { if (v === null || v === '') return ''; if (typeof v === 'number' && v > 30000 && v < 80000) { const d = XLSX.SSF.parse_date_code(v); return `${d.d}/${d.m}/${d.y}`; } return String(v).trim(); };

  // ตารางเงินเดือน: หาคอลัมน์จากชื่อหัวตาราง
  const hc = find(v => clean(v) === 'เงินเดือนที่ได้');
  if (!hc) throw new Error('ไม่พบหัวตาราง "เงินเดือนที่ได้"');
  if (!year) throw new Error('ไม่พบปีของชีต');
  const [hr, gc] = hc, mc = gc - 1;
  const col = {};
  for (let c = gc; c <= Math.min(maxC, gc + 10); c++) {
    const h = clean(R(hr, c)) || '';
    if (h === 'ประกันสังคม') col.sso = c; else if (h.startsWith('กองทุน')) col.fund = c; else if (h === 'ภาษี') col.tax = c;
    else if (h === 'หมายเหตุ') { col.note = c; break; }
    else if (h && h !== 'เงินเดือนที่ได้' && !/^เงินท[ีี่]*ได้/.test(h) && col.tax !== undefined) { col.other = c; col.otherName = h; }
  }
  const leaveCols = [];
  if (col.note !== undefined) for (let c = col.note; c <= col.note + 3; c++) {
    const lb = clean(R(hr + 1, c)); if (c === col.note || ['ลาป่วย','ลากิจ','ลาพักร้อน','อื่นๆ'].includes(lb)) leaveCols.push([c, ['ลาป่วย','ลากิจ','ลาพักร้อน','อื่นๆ'].includes(lb) ? lb : '']);
  }
  const income = [];
  for (let r = hr + 1; r <= maxR; r++) {
    const label = clean(R(r, mc));
    if (label === 'รวม') break;
    const m = monthOf(label);
    if (!m) continue;
    const notes = leaveCols.map(([c, lb]) => { const t = noteText(R(r, c)); return t ? (lb ? `${lb} ${t}` : t) : ''; }).filter(Boolean);
    const row = { id: 'r' + (r + 1), month: m, label: '', gross: n(R(r, gc)), sso: n(R(r, col.sso)), fund: n(R(r, col.fund)), tax: n(R(r, col.tax)), note: notes.join(', ') };
    if (col.other !== undefined) { row.other = n(R(r, col.other)); row.otherName = col.otherName; }
    income.push(row);
  }
  if (!income.length) throw new Error('ไม่พบข้อมูลเงินเดือนรายเดือน');
  // เดือนที่ซ้ำ: แถวที่ไม่หักประกันสังคมและกองทุนถือเป็นโบนัส
  for (let m = 1; m <= 12; m++) {
    const same = income.filter(x => x.month === m);
    if (same.length < 2) continue;
    let bonus = same.filter(x => !x.sso && !x.fund);
    if (!bonus.length || bonus.length === same.length) bonus = same.slice(1);
    bonus.forEach(x => (x.label = 'โบนัส'));
  }

  // ภาษี: แถวที่มี 100000 ตามด้วย 60000
  const tax = { personal: 60000, others: [] };
  const tc = find((v, r, c) => n(v) === 100000 && n(R(r, c + 1)) === 60000);
  if (tc) { const other = n(R(tc[0], tc[1] + 5)); if (other) tax.others.push({ name: 'ลดหย่อนอื่นๆ', amount: other }); }

  // กฎแบ่งเงิน
  let rule = { base: 0, parts: [] };
  const rc = find(v => clean(v) === 'ความจำเป็น');
  if (rc) {
    const [r0, c0] = rc;
    rule.base = n(R(r0 - 1, c0 + 2));
    for (let r = r0; r <= maxR; r++) { const nm = clean(R(r, c0)); if (!nm) break; rule.parts.push({ name: nm, pct: Math.round(n(R(r, c0 + 1)) * 1000) / 10 }); }
  }

  // งบรายเดือน: เซลล์ "เดือน" ที่ช่องถัดไปเป็นชื่อเดือน
  const budgets = {};
  for (let r = range.s.r; r <= maxR; r++) for (let c = range.s.c; c <= maxC; c++) {
    if (clean(R(r, c)) !== 'เดือน') continue;
    const m = monthOf(R(r, c + 1)); if (!m) continue;
    const items = [], cells = [];
    for (let rr = r + 2; rr <= Math.min(maxR, r + 30); rr++) {
      const raw = R(rr, c), nm = typeof raw === 'string' ? norm(clean(raw)) : '';
      if (nm.includes('กดเงินสด')) { cells.cashRow = rr; break; }
      if (nm === 'เดือน') break;
      if (!nm) continue;
      items.push({ name: nm, amount: n(R(rr, c + 1)), saving: SAVING.some(s => nm.startsWith(s)), cash: CASH.some(s => nm.startsWith(s)) });
      cells.push((rr) + ',' + (c + 1));
    }
    // ใช้สูตร "คงเหลือ" และ "กดเงินสด" ใน sheet เพื่อให้ตัวเลขตรงกับที่คำนวณไว้
    const leftRefs = refs(F(r + 1, c + 3));
    if (cells.some(k => leftRefs.has(k))) items.forEach((it, i) => { if (!leftRefs.has(cells[i])) it.offBudget = true; });
    if (cells.cashRow !== undefined) { const cashRefs = refs(F(cells.cashRow, c + 1));
      if (cells.some(k => cashRefs.has(k))) items.forEach((it, i) => { it.cash = cashRefs.has(cells[i]); }); }
    budgets[String(m)] = { date: toISODate(XLSX, R(r, c + 2)), income: n(R(r + 1, c + 1)), items };
  }

  return { ok: true, year, data: { year, income, budgets, tax, rule } };
}
