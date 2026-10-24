/** Taheri Unity Ride 2026 – Google Apps Script backend.
 *  Bound to a Google Sheet: Extensions > Apps Script > paste > Deploy > Web app (Execute as: Me, Access: Anyone).
 *  After ANY edit: Deploy > Manage deployments > Edit > New version > Deploy. */
const ADMIN = {user: 'admin', pass: 'ChangeMe123'};   // <-- change this
const COLS = {
  Participants:  ['id','name','category','phone','verifiedPos'],
  Volunteers:    ['username','password','name','phone','checkpointId'],
  Checkpoints:   ['id','name','order'],
  Scans:         ['time','participantId','checkpointId','by'],
  ActivityUsers: ['username','password','name','phone','activityId'],   // activityId: A1 = Weight, A2 = Clues
  ActivityLogs:  ['participantId','activityId','start','end','result','notes','by']
};
const val_ = v => v instanceof Date ? v.toISOString() : String(v);
const sh_ = n => { const ss = SpreadsheetApp.getActiveSpreadsheet(); let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(COLS[n]); s.setFrozenRows(1); } return s; };
const rows_ = n => { const v = sh_(n).getDataRange().getValues(); const h = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((k,i) => [k, val_(r[i])]))); };
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const cache_ = () => CacheService.getScriptCache();
const cpsSorted_ = () => rows_('Checkpoints').sort((a,b) => a.order - b.order);   // ascending order

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
    else {
      const v = rows_('Volunteers').find(x => x.username === p.user && x.password === p.pass);
      if (v) sess = {role:'volunteer',user:v.username,name:v.name,checkpointId:v.checkpointId};
      else { const a = rows_('ActivityUsers').find(x => x.username === p.user && x.password === p.pass);
        if (a) sess = {role:'activity',user:a.username,name:a.name,activityId:a.activityId}; }
    }
    if (!sess) return {ok:false,error:'Wrong username or password'};
    const token = Utilities.getUuid(); cache_().put(token, JSON.stringify(sess), 21600);
    return {ok:true,token,session:sess};
  }
  const raw = cache_().get(p.token); if (!raw) return {ok:false,error:'auth'};
  const s = JSON.parse(raw), admin = s.role === 'admin';
  if (s.role === 'activity') return act_(p, s);

  if (p.action === 'state') {
    if (admin) return {ok:true,data:{Participants:rows_('Participants'),Volunteers:rows_('Volunteers'),Checkpoints:rows_('Checkpoints'),
      Scans:rows_('Scans'),ActivityUsers:rows_('ActivityUsers'),ActivityLogs:rows_('ActivityLogs')}};
    const cps = cpsSorted_(), all = rows_('Scans'), parts = rows_('Participants');
    const cp = cps.find(c => c.id === s.checkpointId) || null;
    const mine = all.filter(x => x.checkpointId === s.checkpointId);
    const recent = mine.slice(-15).reverse().map(x => ({time:x.time,id:x.participantId,name:(parts.find(q => q.id === x.participantId)||{}).name||''}));
    const report = cps.map(c => ({id:c.id,name:c.name,count:all.filter(x => x.checkpointId === c.id).length}));
    return {ok:true,cp,total:parts.length,count:mine.length,recent,report};
  }
  if (p.action === 'scan') {
    const cpId = admin ? p.checkpointId : s.checkpointId;           // volunteers are locked to their checkpoint
    if (!cpId) return {ok:false,error:'No checkpoint assigned. Ask admin.'};
    const id = String(p.code || '').trim().toUpperCase();
    const part = rows_('Participants').find(x => x.id.toUpperCase() === id);
    if (!part) return {ok:true,status:'unknown',msg:'Participant not found: ' + id};
    const scans = rows_('Scans').filter(x => x.participantId === part.id);
    if (scans.some(x => x.checkpointId === cpId)) return {ok:true,status:'dup',msg:'Already scanned here',name:part.name,id:part.id};
    const cps = cpsSorted_(), idx = cps.findIndex(c => c.id === cpId);
    if (idx > 0 && !(admin && p.force)) {                            // previous checkpoint must be scanned first
      const prev = cps[idx - 1];
      if (!scans.some(x => x.checkpointId === prev.id))
        return {ok:true,status:'blocked',msg:'BLOCKED – not scanned at previous checkpoint: ' + prev.name,name:part.name,id:part.id};
    }
    sh_('Scans').appendRow([new Date().toISOString(), part.id, cpId, s.user]);
    return {ok:true,status:'ok',name:part.name,id:part.id,category:part.category};
  }
  if (!admin) return {ok:false,error:'Admin only'};
  if (p.action === 'save') {
    const cols = COLS[p.sheet]; if (!cols || p.sheet === 'Scans' || p.sheet === 'ActivityLogs') return {ok:false,error:'Bad sheet'};
    const sheet = sh_(p.sheet), vals = sheet.getDataRange().getValues();
    const line = cols.map(c => p.row[c] === undefined ? '' : p.row[c]);
    const idx = vals.findIndex((r,i) => i > 0 && String(r[0]) === String(line[0]));
    if (idx > 0) sheet.getRange(idx + 1, 1, 1, cols.length).setValues([line]); else sheet.appendRow(line);
    return {ok:true};
  }
  if (p.action === 'actSave') {
    const r = p.row, sheet = sh_('ActivityLogs'), v = sheet.getDataRange().getValues();
    const line = [r.participantId, r.activityId, r.start || '', r.end || '', r.result === undefined ? '' : r.result, r.notes || '', s.user];
    const i = v.findIndex((x,k) => k > 0 && String(x[0]) === r.participantId && String(x[1]) === r.activityId);
    if (i > 0) sheet.getRange(i + 1, 1, 1, 7).setValues([line]); else sheet.appendRow(line);
    return {ok:true};
  }
  if (p.action === 'actRemove') {
    const sheet = sh_('ActivityLogs'), v = sheet.getDataRange().getValues();
    for (let i = v.length - 1; i > 0; i--) if (String(v[i][0]) === p.participantId && String(v[i][1]) === p.activityId) sheet.deleteRow(i + 1);
    return {ok:true};
  }
  if (p.action === 'remove') {
    const sheet = sh_(p.sheet), v = sheet.getDataRange().getValues();
    if (p.sheet === 'Scans') {
      for (let i = v.length - 1; i > 0; i--) if (v[i][1] === p.participantId && v[i][2] === p.checkpointId) sheet.deleteRow(i + 1);
      return {ok:true};
    }
    for (let i = v.length - 1; i > 0; i--) if (String(v[i][0]) === String(p.key)) sheet.deleteRow(i + 1);
    return {ok:true};
  }
  return {ok:false,error:'Unknown action'};
}

/** Activity staff: scan start, then scan again to finish with result (weight / clues completed). */
function act_(p, s) {
  const aid = s.activityId, parts = rows_('Participants');
  if (p.action === 'state') {
    const logs = rows_('ActivityLogs').filter(x => x.activityId === aid);
    const recent = logs.slice(-15).reverse().map(x => ({id:x.participantId,name:(parts.find(q => q.id === x.participantId)||{}).name||'',start:x.start,end:x.end,result:x.result}));
    return {ok:true,activityId:aid,started:logs.length,completed:logs.filter(x => x.end).length,recent};
  }
  if (p.action !== 'actScan' && p.action !== 'actEnd') return {ok:false,error:'Not allowed'};
  const id = String(p.code || '').trim().toUpperCase();
  const part = parts.find(x => x.id.toUpperCase() === id);
  if (!part) return {ok:true,status:'unknown',msg:'Participant not found: ' + id};
  const sheet = sh_('ActivityLogs'), v = sheet.getDataRange().getValues();
  const i = v.findIndex((r,k) => k > 0 && String(r[0]) === part.id && String(r[1]) === aid);
  const now = new Date().toISOString();
  if (p.action === 'actScan') {
    if (i < 0) { sheet.appendRow([part.id, aid, now, '', '', '', s.user]); return {ok:true,status:'started',name:part.name,id:part.id,time:now}; }
    if (!v[i][3]) return {ok:true,status:'open',name:part.name,id:part.id,start:val_(v[i][2])};
    return {ok:true,status:'done',msg:'Already completed',name:part.name,id:part.id,result:String(v[i][4])};
  }
  if (i < 0) return {ok:false,error:'Start not recorded for this participant'};
  if (v[i][3]) return {ok:false,error:'Already completed'};
  if (p.result === '' || isNaN(p.result)) return {ok:false,error:'Result must be a number'};
  sheet.getRange(i + 1, 4, 1, 3).setValues([[now, Number(p.result), p.notes || '']]);
  return {ok:true};
}
