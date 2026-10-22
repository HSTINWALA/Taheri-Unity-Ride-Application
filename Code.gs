/** Taheri Unity Ride 2026 – Google Apps Script backend.
 *  Bound to a Google Sheet: Extensions > Apps Script > paste > Deploy > Web app (Execute as: Me, Access: Anyone). */
const ADMIN = {user: 'admin', pass: 'ChangeMe123'};   // <-- change this
const COLS = {
  Participants: ['id','name','category','phone','verifiedPos'],
  Volunteers:   ['username','password','name','phone','checkpointId'],
  Checkpoints:  ['id','name','order'],
  Scans:        ['time','participantId','checkpointId','by']
};
const sh_ = n => { const ss = SpreadsheetApp.getActiveSpreadsheet(); let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(COLS[n]); s.setFrozenRows(1); } return s; };
const rows_ = n => { const v = sh_(n).getDataRange().getValues(); const h = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((k,i) => [k, String(r[i])]))); };
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const cache_ = () => CacheService.getScriptCache();

function doGet() { return out_({ok:true,msg:'Unity Ride API running'}); }

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return out_(handle_(JSON.parse(e.postData.contents))); }
  catch (err) { return out_({ok:false,error:String(err)}); }
  finally { lock.releaseLock(); }
}

function handle_(p) {
  Object.keys(COLS).forEach(sh_);
  if (p.action === 'login') {
    let sess = null;
    if (p.role === 'admin') { if (p.user === ADMIN.user && p.pass === ADMIN.pass) sess = {role:'admin',user:'admin',name:'Admin'}; }
    else { const v = rows_('Volunteers').find(x => x.username === p.user && x.password === p.pass);
      if (v) sess = {role:'volunteer',user:v.username,name:v.name,checkpointId:v.checkpointId}; }
    if (!sess) return {ok:false,error:'Wrong username or password'};
    const token = Utilities.getUuid(); cache_().put(token, JSON.stringify(sess), 21600);
    return {ok:true,token,session:sess};
  }
  const raw = cache_().get(p.token); if (!raw) return {ok:false,error:'auth'};
  const s = JSON.parse(raw), admin = s.role === 'admin';

  if (p.action === 'state') {
    if (admin) return {ok:true,data:{Participants:rows_('Participants'),Volunteers:rows_('Volunteers'),Checkpoints:rows_('Checkpoints'),Scans:rows_('Scans')}};
    const cp = rows_('Checkpoints').find(c => c.id === s.checkpointId) || null;
    const parts = rows_('Participants'), mine = rows_('Scans').filter(x => x.checkpointId === s.checkpointId);
    const recent = mine.slice(-15).reverse().map(x => ({time:x.time,id:x.participantId,name:(parts.find(q => q.id === x.participantId)||{}).name||''}));
    return {ok:true,cp,total:parts.length,count:mine.length,recent};
  }
  if (p.action === 'scan') {
    const cpId = admin ? p.checkpointId : s.checkpointId;           // volunteers are locked to their checkpoint
    if (!cpId) return {ok:false,error:'No checkpoint assigned. Ask admin.'};
    const id = String(p.code || '').trim().toUpperCase();
    const part = rows_('Participants').find(x => x.id.toUpperCase() === id);
    if (!part) return {ok:true,status:'unknown',msg:'Participant not found: ' + id};
    if (rows_('Scans').some(x => x.participantId === part.id && x.checkpointId === cpId))
      return {ok:true,status:'dup',msg:'Already scanned here',name:part.name,id:part.id};
    sh_('Scans').appendRow([new Date().toISOString(), part.id, cpId, s.user]);
    return {ok:true,status:'ok',name:part.name,id:part.id,category:part.category};
  }
  if (!admin) return {ok:false,error:'Admin only'};
  if (p.action === 'save') {
    const cols = COLS[p.sheet]; if (!cols || p.sheet === 'Scans') return {ok:false,error:'Bad sheet'};
    const sheet = sh_(p.sheet), vals = sheet.getDataRange().getValues();
    const line = cols.map(c => p.row[c] === undefined ? '' : p.row[c]);
    const idx = vals.findIndex((r,i) => i > 0 && String(r[0]) === String(line[0]));
    if (idx > 0) sheet.getRange(idx + 1, 1, 1, cols.length).setValues([line]); else sheet.appendRow(line);
    return {ok:true};
  }
  if (p.action === 'remove') {
    if (p.sheet === 'Scans') { // remove one scan by participant+checkpoint
      const sheet = sh_('Scans'), v = sheet.getDataRange().getValues();
      for (let i = v.length - 1; i > 0; i--) if (v[i][1] === p.participantId && v[i][2] === p.checkpointId) sheet.deleteRow(i + 1);
      return {ok:true};
    }
    const sheet = sh_(p.sheet), v = sheet.getDataRange().getValues();
    for (let i = v.length - 1; i > 0; i--) if (String(v[i][0]) === String(p.key)) sheet.deleteRow(i + 1);
    return {ok:true};
  }
  return {ok:false,error:'Unknown action'};
}
