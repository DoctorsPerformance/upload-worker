// Upload worker (runs on a schedule in the public "upload-worker" repository).
//
// The online dashboard cannot write files itself: a signed-in person's browser encrypts the
// workbook to this worker's public key and posts it to a private ntfy.sh topic. Here we read the
// topic, decrypt each new upload, check it strictly against its page's template (the same check the
// page runs: public/template-check.js and public/templates/ of the data repository — every column,
// the figures above the table filled in, a real branch / month / year on every row), and commit it
// into that page's folder of the private data repository ("<branch>/Allied/<department>/…xlsx"). That push starts the dashboard's own publish workflow,
// which merges the department's rows into its month files and rebuilds the site.
// A short status ({ ref, ok, note }, no personal data) goes back so the page can tell the uploader.
//
// Secrets: UPLOAD_TOPIC, UPLOAD_PRIVATE (JWK, base64url), DATA_REPO (owner/name), and an SSH deploy
// key with write access to DATA_REPO (configured by the workflow).
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const ExcelJS = require('exceljs');

const RELAY = 'https://ntfy.sh';
const { UPLOAD_TOPIC: TOPIC, UPLOAD_PRIVATE: PRIV, DATA_REPO } = process.env;
const DATA_DIR = 'Source of the Incentive';
const BRANCHES = ['SGH-Cairo', 'SGH-Alex'];
// What may be uploaded from the online page: category -> folder, and its departments.
// units: the only departments allowed (one per upload); anyUnit: a whole-month file, or one specialty.
const ALLOWED = {
  allied: { folder: 'Allied', units: ['Anesthesia', 'Emergency Room', 'Intensive Care Unit', 'Laboratory', 'Radiation Oncology', 'Medical Imaging', 'Physiotherapy', 'Pharmacy', 'Home care'] },
  bd: { folder: 'Separate/Business Development', units: ['Corporate Sales', 'External Doctors', 'IVP Visiting Professors'] },
  opd: { folder: 'OPD', anyUnit: true },
  residents: { folder: 'Residents', anyUnit: true },
  admission: { folder: 'Separate/Admission & Discharge', anyUnit: true },
  patientrel: { folder: 'Separate/Patient Relation', anyUnit: true },
  legal: { folder: 'Separate/Legal Affairs', anyUnit: true },
  collection: { folder: 'Separate/CAM Collection', anyUnit: true },
  quarterly: { folder: 'Quarterly', anyUnit: true },
};
// letters, digits, spaces and a few marks; no slashes and no ".." (it becomes a folder name)
const SAFE_UNIT = /^[\p{L}\p{N}][\p{L}\p{N} ,&()'-]{0,59}$/u;
const ID_NAMES = ['id', 'employeeid', 'empid', 'code', 'doctorid', 'doctorcode', 'staffid', 'empno', 'employeeno', 'employeecode', 'empcode'];
const NAME_NAMES = ['name', 'doctor', 'doctorname', 'employeename', 'fullname'];
const PAY_NAMES = ['incentives', 'incentive', 'totalsalary', 'salary', 'basicsalary', 'monthlyrevenue', 'revenue', 'grossincentives', 'finalincentives'];
const MAX_BYTES = 5 * 1024 * 1024;
const work = path.join(process.cwd(), 'data-repo');

const git = (...args) => execFileSync('git', args, { cwd: work, stdio: ['ignore', 'pipe', 'pipe'] }).toString().trim();
const norm = (s) => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');

function decrypt(body) {
  const m = JSON.parse(body);
  const jwk = JSON.parse(Buffer.from(PRIV, 'base64url').toString('utf8'));
  const ecdh = crypto.createECDH('prime256v1');
  ecdh.setPrivateKey(Buffer.from(jwk.d, 'base64url'));
  const key = ecdh.computeSecret(Buffer.from(m.e, 'base64'));
  const data = Buffer.from(m.c, 'base64');
  const d = crypto.createDecipheriv('aes-256-gcm', key, Buffer.from(m.i, 'base64'));
  d.setAuthTag(data.subarray(data.length - 16));
  return JSON.parse(Buffer.concat([d.update(data.subarray(0, data.length - 16)), d.final()]).toString('utf8'));
}

/**
 * The workbook must open and have the template's table: a header (row 1, or under the department
 * figures) with ID and Name — or, like Home care's, Month and Incentives only — and at least one row.
 */
async function inspect(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  for (const ws of wb.worksheets) {
    for (let r = 1; r <= Math.min(15, ws.rowCount); r++) {
      const cols = {};
      ws.getRow(r).eachCell((c, i) => { const k = norm(c.text); if (!cols[k]) cols[k] = i; });
      const find = (names) => names.map((n) => cols[n]).find(Boolean);
      const idCol = find(ID_NAMES); const nameCol = find(NAME_NAMES); const payCol = find(PAY_NAMES);
      const people = idCol && nameCol;
      if (!people && !(cols.month && payCol)) continue;
      if (!payCol) return { error: 'the file has no Incentives / Total Salary column' };
      let rows = 0;
      const months = new Set();
      for (let k = r + 1; k <= ws.rowCount; k++) {
        const row = ws.getRow(k);
        const has = people ? row.getCell(idCol).text.trim() || row.getCell(nameCol).text.trim() : row.getCell(payCol).text.trim();
        if (!has) continue;
        rows++;
        const m = cols.month ? row.getCell(cols.month).text.trim() : '';
        if (m) months.add(`${m}${cols.year ? ` ${row.getCell(cols.year).text.trim()}` : ''}`);
      }
      return rows ? { rows, months: [...months].filter(Boolean) } : { error: 'the table in the file is empty' };
    }
  }
  return { error: 'the file is not a dashboard template (no ID and Name columns)' };
}

// A sheet as rows of cell values (what template-check.js reads).
const cellValue = (v) => {
  if (v && typeof v === 'object' && !(v instanceof Date)) {
    if ('result' in v) return v.result;
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if ('text' in v) return v.text;
    return null;
  }
  return v;
};
function gridOf(ws) {
  const g = [];
  ws.eachRow({ includeEmpty: true }, (row, r) => { const a = []; row.eachCell({ includeEmpty: false }, (c, i) => { a[i - 1] = cellValue(c.value); }); g[r - 1] = a; });
  for (let i = 0; i < g.length; i++) if (!g[i]) g[i] = [];
  return g;
}
async function firstSheet(src) {
  const wb = new ExcelJS.Workbook();
  if (Buffer.isBuffer(src)) await wb.xlsx.load(src); else await wb.xlsx.readFile(src);
  if (!wb.worksheets.length) throw new Error('no sheet');
  return gridOf(wb.worksheets[0]);
}
/** The page's template check (from the data repository), or null when the page has no template. */
async function templateCheck(TC, p, buffer) {
  const key = TC.templateKey(p.category, p.unit);
  if (!key) return null;
  const tpl = await firstSheet(path.join(work, 'public', 'templates', TC.TEMPLATES[key]));
  let grid;
  try { grid = await firstSheet(buffer); } catch { return { ok: false, errors: ['the file could not be opened as Excel'] }; }
  return TC.validate(grid, tpl, { branch: p.branch, unit: p.unit || null, name: TC.TEMPLATES[key].replace(/\.xlsx$/i, '') });
}

async function status(list) {
  for (const st of list) {
    await fetch(`${RELAY}/${TOPIC}s`, { method: 'POST', body: JSON.stringify(st), headers: { 'Content-Type': 'text/plain' } }).catch(() => {});
  }
}

(async () => {
  if (!TOPIC || !PRIV || !DATA_REPO) { console.log('Not configured (UPLOAD_TOPIC / UPLOAD_PRIVATE / DATA_REPO).'); return; }
  const text = await (await fetch(`${RELAY}/${TOPIC}/json?poll=1&since=12h`)).text();
  const messages = text.split('\n').filter(Boolean).map((l) => { try { return JSON.parse(l); } catch { return null; } })
    .filter((m) => m && m.event === 'message');
  if (!messages.length) { console.log('No uploads.'); return; }

  // DATA_REPO_URL: only for testing against a local copy
  execFileSync('git', ['clone', '-q', '--depth', '1', process.env.DATA_REPO_URL || `git@github.com:${DATA_REPO}.git`, work], { stdio: 'inherit' });
  git('config', 'user.name', 'dashboard-upload');
  git('config', 'user.email', 'dashboard-upload@users.noreply.github.com');
  const doneFile = path.join(work, 'uploads', 'processed.json');
  const TC = require(path.join(work, 'public', 'template-check.js'));
  let done = [];
  try { done = JSON.parse(fs.readFileSync(doneFile, 'utf8')); } catch { /* first upload */ }
  const fresh = messages.filter((m) => !done.includes(m.id));
  if (!fresh.length) { console.log(`${messages.length} upload(s) on the relay, all already filed.`); return; }

  const statuses = [];
  const filed = [];
  for (const msg of fresh) {
    done.push(msg.id);
    let p;
    try {
      const body = msg.attachment ? await (await fetch(msg.attachment.url)).text() : msg.message;
      p = decrypt(body);
    } catch (err) {
      console.log(`${msg.id}: not a readable upload (${err.message})`);
      continue; // not ours / expired: nothing to report back to
    }
    const fail = (note) => { statuses.push({ ref: p.ref, ok: false, note }); console.log(`${msg.id}: rejected — ${note}`); };
    const allow = ALLOWED[p.category];
    if (!allow) { fail('this page does not take uploads from the online dashboard yet'); continue; }
    if (!BRANCHES.includes(p.branch)) { fail(`unknown branch ${p.branch}`); continue; }
    if (allow.units && !allow.units.includes(p.unit)) { fail(`unknown department ${p.unit}`); continue; }
    if (allow.anyUnit && p.unit && !SAFE_UNIT.test(p.unit)) { fail('the specialty name has characters that are not allowed'); continue; }
    const buffer = Buffer.from(String(p.data || ''), 'base64');
    if (buffer.length > MAX_BYTES) { fail('the file is larger than 5 MB'); continue; }
    if (buffer.subarray(0, 2).toString('latin1') !== 'PK') { fail('not an Excel .xlsx file'); continue; }
    let check;
    try { check = await templateCheck(TC, p, buffer); } catch (err) { check = { ok: false, errors: [`the template could not be checked (${err.message})`] }; }
    if (check && !check.ok) { fail(`not accepted — ${TC.summary(check.errors, 5).replace(/\n/g, ' · ')}`); continue; }
    let info;
    if (check) info = { rows: check.rows, months: [check.month] };
    else {
      try { info = await inspect(buffer); } catch { info = { error: 'the file could not be opened as Excel' }; }
      if (info.error) { fail(info.error); continue; }
    }
    const stamp = new Date(msg.time * 1000).toISOString().slice(0, 16).replace(/[-:T]/g, '');
    const safe = String(p.name || 'upload.xlsx').replace(/[^\w .()&-]+/g, '_').replace(/\.xlsx?$/i, '').slice(0, 80);
    const rel = path.posix.join(DATA_DIR, p.branch, allow.folder, ...(p.unit ? [p.unit] : []), `${stamp} ${safe}.xlsx`);
    fs.mkdirSync(path.join(work, path.dirname(rel)), { recursive: true });
    fs.writeFileSync(path.join(work, rel), buffer);
    filed.push(`${p.unit || allow.folder} (${String(p.by || 'unknown').slice(0, 40)})`);
    statuses.push({ ref: p.ref, ok: true, note: `${info.rows} rows${info.months.length ? ` · ${info.months.join(', ')}` : ''}` });
    console.log(`${msg.id}: ${rel} · ${info.rows} rows`);
  }
  fs.mkdirSync(path.dirname(doneFile), { recursive: true });
  fs.writeFileSync(doneFile, `${JSON.stringify(done.slice(-1000), null, 0)}\n`);
  git('add', '-A');
  git('commit', '-q', '-m', filed.length ? `Upload from the dashboard: ${filed.join(', ')}` : 'Upload inbox: nothing filed');
  // The publish workflow may have pushed meanwhile: rebase and retry.
  for (let i = 0; i < 4; i++) {
    try { git('pull', '-q', '--rebase'); git('push', '-q'); break; } catch (err) { if (i === 3) throw err; await new Promise((r) => setTimeout(r, 5000)); }
  }
  await status(statuses);
  console.log(`${filed.length} filed, ${statuses.filter((s) => !s.ok).length} rejected.`);
})().catch((err) => { console.error(err); process.exit(1); });
