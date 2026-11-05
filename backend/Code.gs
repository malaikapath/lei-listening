/**
 * MALAIKA PATH · Listening Tool — backend (Google Apps Script)
 * For the Directorate of Labour, Employment Services, Industrial Relations & OSH
 * Ministry of Gender, Labour and Social Development (Uganda)
 *
 * What this does
 *  - Stores every conversation in a Google Sheet in the owner's Drive (tabs: Responses, Settings, Vouchers, Staff, Audit)
 *  - One response per email address. Progress is saved after every question, so people can stop and resume.
 *  - Scores how fully each person answered, unlocks the right reward, hands out a voucher code from the pool
 *    and emails it from the owner's Gmail (MailApp).
 *  - Serves the web page (index.html) and a JSON API so the page can also be hosted elsewhere (e.g. Vercel).
 *
 * First-time setup: run setup() once from the editor, then Deploy > New deployment > Web app
 *   Execute as: Me   ·   Who has access: Anyone
 */

const PROPS = PropertiesService.getScriptProperties();
const TABS = {
  responses: ['id','created','updated','status','mode','interviewer','name','email','title','unit','level','phone','ref','verified','dupFlag',
              'a1','a2','a3','a4','a5','a6','w1','w2','w3','w4','w5','w6','depth','tier','rewardStatus','voucher','rewardSentAt','device','code'],
  settings:  ['key','value'],
  vouchers:  ['code','tier','status','assignedTo','assignedName','assignedAt','note'],
  staff:     ['name','email','unit','title'],
  audit:     ['at','actor','action','details']
};

/* ------------------------------------------------------------------ setup */
function setup(){
  let id = PROPS.getProperty('SHEET_ID');
  let ss = id ? SpreadsheetApp.openById(id) : null;
  if(!ss){
    ss = SpreadsheetApp.create('Malaika Path · Listening Tool · Responses');
    PROPS.setProperty('SHEET_ID', ss.getId());
  }
  Object.keys(TABS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if(!sh){ sh = ss.insertSheet(name); }
    sh.getRange(1,1,1,TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const def = ss.getSheetByName('Sheet1'); if(def && ss.getSheets().length > 1) ss.deleteSheet(def);
  if(!PROPS.getProperty('ADMIN_KEY')) PROPS.setProperty('ADMIN_KEY', 'malaika-' + Math.random().toString(36).slice(2,8));
  log_('setup','setup','Sheet ready: ' + ss.getUrl());
  Logger.log('Sheet: ' + ss.getUrl());
  Logger.log('Malaika Path dashboard password (change it in Project Settings > Script properties > ADMIN_KEY): ' + PROPS.getProperty('ADMIN_KEY'));
}

/* ------------------------------------------------------------------ web entry points */
function doGet(e){
  const p = (e && e.parameter) || {};
  if(p.api){ return json_(route_(p.api, p.args ? JSON.parse(p.args) : [])); }
  return json_({ok:true, service:'Malaika Path · Listening to Our Team API', page:'https://rebekahzadokmugisha-boop.github.io/lei-listening/'});
}
function doPost(e){
  let body = {};
  try{ body = JSON.parse(e.postData.contents || '{}'); }catch(err){ return json_({error:'Bad request'}); }
  return json_(route_(body.fn, body.args || []));
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function route_(fn, args){
  const allowed = {apiAdminRelease:1, apiAdminDelete:1, apiSignup:1, apiLogin:1, apiChangeCode:1, apiResendCode:1, apiPublic:1, apiStart:1, apiSave:1, apiSubmit:1, apiAdmin:1, apiAdminSaveSettings:1, apiAdminApprove:1, apiAdminAddVouchers:1, apiAdminResend:1, apiAdminSetStatus:1};
  if(!allowed[fn]) return {error:'Unknown action'};
  const API = {apiAdminRelease, apiAdminDelete, apiSignup, apiLogin, apiChangeCode, apiResendCode, apiPublic, apiStart, apiSave, apiSubmit, apiAdmin, apiAdminSaveSettings, apiAdminApprove, apiAdminAddVouchers, apiAdminResend, apiAdminSetStatus};
  try{ return {ok:true, data: API[fn].apply(null, args)}; }
  catch(err){ return {error: err.message || String(err)}; }
}

/* ------------------------------------------------------------------ helpers */
function ss_(){ return SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID')); }
function sh_(n){ return ss_().getSheetByName(n); }
function rows_(n){
  const sh = sh_(n), v = sh.getDataRange().getValues(); let h = v.shift();
  if(h.length < TABS[n].length || h.join() !== TABS[n].join()){ sh.getRange(1,1,1,TABS[n].length).setValues([TABS[n]]).setFontWeight('bold'); h = TABS[n]; }
  return v.map((r,i) => { const o = {_row:i+2}; h.forEach((k,j) => o[k] = r[j] instanceof Date ? r[j].toISOString() : r[j]); return o; });
}
function writeRow_(n, obj, rowNum){
  const h = TABS[n], vals = [h.map(k => obj[k] === undefined ? '' : obj[k])];
  const sh = sh_(n);
  if(rowNum) sh.getRange(rowNum,1,1,h.length).setValues(vals);
  else sh.appendRow(vals[0]);
}
function log_(actor, action, details){ try{ sh_('audit').appendRow([new Date().toISOString(), actor, action, String(details||'').slice(0,500)]); }catch(e){} }
function now_(){ return new Date().toISOString(); }
function normEmail_(e){ return String(e||'').trim().toLowerCase(); }
function normName_(n){ return String(n||'').toLowerCase().replace(/[^a-z ]/g,' ').split(/\s+/).filter(Boolean).sort().join(' '); }
function words_(t){ return String(t||'').trim().split(/\s+/).filter(Boolean).length; }
function isAdmin_(key){ if(!key || key !== PROPS.getProperty('ADMIN_KEY')) throw new Error('That password is not correct.'); }
function lock_(fn){ const l = LockService.getScriptLock(); l.waitLock(20000); try{ return fn(); } finally { l.releaseLock(); } }

function settings_(){
  const s = {};
  rows_('settings').forEach(r => { try{ s[r.key] = JSON.parse(r.value); }catch(e){ s[r.key] = r.value; } });
  return s; // the page merges these over its own defaults
}
function setting_(k, d){ const s = settings_(); return s[k] === undefined || s[k] === '' ? d : s[k]; }

function levelFromTitle_(t){
  t = String(t||'').toLowerCase();
  if(/assistant commissioner|principal|head of|manager|senior/.test(t)) return 'Management';
  if(/commissioner|director|permanent secretary|minister/.test(t)) return 'Leadership';
  if(/record|\bit\b|ict|driver|secretar|clerk|attendant|account|typist|steno|librar|reception|store|admin assistant|office assistant|askari|security/.test(t)) return 'Support / IT / Records';
  return 'Officer';
}
function verified_(email){
  email = normEmail_(email);
  const domains = String(setting_('allowedDomains','mglsd.go.ug')).split(/[,\s]+/).filter(Boolean).map(d => d.toLowerCase());
  if(domains.some(d => email.endsWith('@'+d))) return true;
  return rows_('staff').some(r => normEmail_(r.email) === email);
}
function depth_(r){
  const targets = setting_('wordTargets',[70,50,70,60,50,0]);
  let score = 0;
  for(let i=0;i<5;i++){ score += Math.min((Number(r['w'+(i+1)])||0) / (targets[i]||50), 1) * 20; }
  return Math.round(score);
}


/* ------------------------------------------------------------------ sign up and personal codes */
function newCode_(){ const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; let o = ''; for(let i=0;i<6;i++) o += c[Math.floor(Math.random()*c.length)]; return o; }
function apiSignup(profile){
  const res = lock_(() => {
    const email = normEmail_(profile.email);
    if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Please check your email address.');
    if(!String(profile.name||'').trim() || !String(profile.title||'').trim() || !profile.unit) throw new Error('Please fill in your name, role and unit.');
    const isVer = verified_(email);
    if(setting_('accessMode','open') === 'list' && !isVer) throw new Error('This email is not on the Directorate staff list. Use your work email, or call the project coordinator.');
    const rs = rows_('responses');
    if(rs.some(r => normEmail_(r.email) === email)) throw new Error('This email is already signed up. Tap "I have my code" to continue, or ask us to email your code again.');
    const dupName = rs.some(r => normName_(r.name) && normName_(r.name) === normName_(profile.name));
    const row = { id: 'r' + Utilities.getUuid().replace(/-/g,'').slice(0,9), created: now_(), updated: now_(), status:'in_progress',
      mode: profile.mode === 'interviewer' ? 'Interviewer-led' : 'Self-completed', interviewer: profile.interviewer || '',
      name: String(profile.name).trim(), email, title: String(profile.title).trim(), unit: profile.unit, level: levelFromTitle_(profile.title),
      phone: '', ref: profile.ref || '', verified: isVer ? 'yes' : 'no', dupFlag: dupName ? 'Same name as another response' : '',
      device: String(profile.device||'').slice(0,120), code: newCode_() };
    writeRow_('responses', row);
    log_(email, 'signed up', row.unit + ' · ' + row.title);
    return row;
  });
  emailCode_(res);
  return {id: res.id, code: res.code, emailed: true};
}
function apiLogin(email, code){
  const r = rows_('responses').find(x => normEmail_(x.email) === normEmail_(email));
  if(!r || String(r.code).toUpperCase() !== String(code||'').trim().toUpperCase()) throw new Error('That email and code do not match. Check the code in your email, or tap "Email me my code".');
  if(r.status === 'submitted') return {done:true, tier:r.tier, rewardStatus:r.rewardStatus, voucher:r.voucher, name:r.name};
  return {id:r.id, answers:[1,2,3,4,5,6].map(i => r['a'+i] || ''), profile:{name:r.name, email:r.email, title:r.title, unit:r.unit}};
}
function apiChangeCode(id, code){
  code = String(code||'').trim().toUpperCase();
  if(!/^[A-Z0-9]{4,20}$/.test(code)) throw new Error('Use 4 to 20 letters or numbers, with no spaces.');
  return lock_(() => { const r = rows_('responses').find(x => x.id === id); if(!r) throw new Error('Session not found.');
    r.code = code; writeRow_('responses', r, r._row); log_(r.email, 'changed code', ''); return true; });
}
function apiResendCode(email){
  const r = rows_('responses').find(x => normEmail_(x.email) === normEmail_(email));
  if(r) emailCode_(r);
  return true; // same answer either way, so nobody can test which emails are signed up
}
function emailCode_(r){
  const s = settings_();
  const html = `<div style="font-family:Arial,sans-serif;max-width:560px;color:#231A15;line-height:1.55">
    <p style="font-size:12px;letter-spacing:1px;text-transform:uppercase;color:#6B5F57"><b>${s.shortName || 'LEISH'}</b> · Ministry of Gender, Labour and Social Development · powered by Malaika Path</p>
    <h2 style="margin:6px 0 12px">Your personal code</h2>
    <p>Hello ${String(r.name||'').split(' ')[0]}, thank you for joining <i>Listening to Our Team</i>. Use this code with your email to come back and continue at any time:</p>
    <p style="font-size:28px;letter-spacing:6px;font-family:monospace;background:#F7ECD3;padding:12px 18px;border-radius:8px;display:inline-block">${r.code}</p>
    <p><a href="https://rebekahzadokmugisha-boop.github.io/lei-listening/">Open Listening to Our Team</a></p>
    <p style="background:#F6F4F0;padding:10px 12px;border-radius:8px;font-size:13px">Please do not reply to this email; this inbox is not monitored. For any question, call the ${s.contactName || 'Malaika Path project coordinator'} on <b>${s.contactPhone || '0772 550 134'}</b>.</p></div>`;
  try{ MailApp.sendEmail({to:r.email, subject:'Your code · Listening to Our Team', htmlBody:html, name:'Malaika Path (no reply)'}); log_('system','code email', r.email); }
  catch(e){ log_('system','code email failed', r.email + ' · ' + e.message); }
}

/* ------------------------------------------------------------------ public API */
function apiPublic(){
  const rs = rows_('responses'), units = {};
  rs.filter(r => r.status === 'submitted').forEach(r => units[r.unit] = (units[r.unit]||0) + 1);
  const fullUnlocked = rs.filter(r => r.status === 'submitted' && r.tier && r.tier !== 'none').length;
  const appUrl = 'https://rebekahzadokmugisha-boop.github.io/lei-listening/';
  return { settings: settings_(), units, total: rs.filter(r => r.status === 'submitted').length, fullUnlocked, appUrl };
}

/** Called once, after the "About you" page. Returns the session to continue. */
function apiStart(profile, clientId){
  return lock_(() => {
    const email = normEmail_(profile.email);
    if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Please check your email address. We send your reward there.');
    if(!String(profile.name||'').trim() || !String(profile.title||'').trim() || !profile.unit) throw new Error('Please fill in your name, title and unit.');
    const isVer = verified_(email);
    if(setting_('accessMode','open') === 'list' && !isVer) throw new Error('This email is not on the Directorate staff list. Use your work email, or ask your Commissioner to add you.');
    const rs = rows_('responses');
    const mine = rs.find(r => normEmail_(r.email) === email);
    if(mine && mine.status === 'submitted') return {done:true, id:mine.id, tier:mine.tier, rewardStatus:mine.rewardStatus};
    if(mine){
      const same = (clientId && clientId === mine.id) || normName_(mine.name) === normName_(profile.name);
      mine.updated = now_(); mine.title = profile.title; mine.unit = profile.unit; mine.phone = profile.phone ? "'" + String(profile.phone).replace(/^'/,'') : mine.phone; mine.level = levelFromTitle_(profile.title);
      writeRow_('responses', mine, mine._row);
      return {id:mine.id, resumed:true, answers: same ? [1,2,3,4,5,6].map(i => mine['a'+i] || '') : []};
    }
    const dupName = rs.some(r => normName_(r.name) && normName_(r.name) === normName_(profile.name));
    const row = { id: 'r' + Utilities.getUuid().replace(/-/g,'').slice(0,9), created: now_(), updated: now_(), status:'in_progress',
      mode: profile.mode === 'interviewer' ? 'Interviewer-led' : 'Self-completed', interviewer: profile.interviewer || '',
      name: String(profile.name).trim(), email, title: String(profile.title).trim(), unit: profile.unit, level: levelFromTitle_(profile.title),
      phone: profile.phone ? "'" + String(profile.phone).replace(/^'/,'') : '', ref: profile.ref || '', verified: isVer ? 'yes' : 'no', dupFlag: dupName ? 'Same name as another response' : '',
      device: String(profile.device||'').slice(0,120) };
    writeRow_('responses', row);
    log_(email, 'started', row.unit + ' · ' + row.title);
    return {id: row.id, resumed:false, answers:[]};
  });
}

/** Saves one answer. Called after every question so nothing is lost. */
function apiSave(id, q, text){
  return lock_(() => {
    const r = rows_('responses').find(x => x.id === id);
    if(!r) throw new Error('We could not find your session. Please start again from the first page.');
    if(r.status === 'submitted') throw new Error('You have already submitted your answers. Thank you.');
    q = Number(q); if(!(q >= 1 && q <= 6)) throw new Error('Unknown question');
    r['a'+q] = String(text||'').slice(0, 20000); r['w'+q] = words_(text); r.updated = now_();
    r.depth = depth_(r);
    writeRow_('responses', r, r._row);
    return {saved: r.updated, depth: r.depth};
  });
}

/** Final submit: locks the response, unlocks the reward and emails it. */
function apiSubmit(id){
  const res = lock_(() => {
    const r = rows_('responses').find(x => x.id === id);
    if(!r) throw new Error('We could not find your session.');
    if(r.status === 'submitted') return {tier:r.tier, rewardStatus:r.rewardStatus, voucher:r.voucher};
    for(let i=1;i<=5;i++){ if(!String(r['a'+i]||'').trim()) throw new Error('Question ' + i + ' still needs an answer.'); }
    r.depth = depth_(r);
    r.status = 'submitted'; r.updated = now_();
    if(r.mode === 'Interviewer-led'){ r.tier = 'none'; r.rewardStatus = 'Not eligible (interviewer-led)'; }
    else if(setting_('rewardsLive', false) !== true){ r.tier = ''; r.rewardStatus = 'Rewards not open yet'; }
    else {
      r.tier = pickRung_(r.depth >= Number(setting_('fullThreshold', 75)));
      const needsApproval = r.verified !== 'yes' && setting_('approveUnverified', true) !== false;
      if(needsApproval){ r.rewardStatus = 'Awaiting approval'; }
      else assignVoucher_(r);
    }
    writeRow_('responses', r, r._row);
    log_(r.email, 'submitted', 'depth ' + r.depth + ' · ' + r.tier + ' · ' + r.rewardStatus);
    return r;
  });
  if(res.rewardStatus && res.tier) emailReward_(res);
  return {tier:res.tier, rewardStatus:res.rewardStatus, voucher:res.voucher, depth:res.depth};
}

/* The reward ladder: rewards go out in order of completion. Full answers take the next
   rung that still has places; short answers go to the last rung. count 0 = no limit. */
const LADDER_DEFAULT = [
  {key:'cj', name:'Café Javas breakfast combo: a croissant and hot chocolate', count:5, howTo:'Show this code at any Café Javas branch to redeem it.'},
  {key:'endiro', name:'Endiro Coffee breakfast combo', count:10, howTo:'Show this code at any Endiro Coffee branch to redeem it.'},
  {key:'office', name:'A thank-you gift from the Minister’s office', count:35, howTo:'Take this code to the Minister’s office to collect your gift. The project coordinator will confirm the time.', auto:true}
];
function ladder_(){ const l = setting_('rewards', LADDER_DEFAULT); return (Array.isArray(l) && l.length) ? l : LADDER_DEFAULT; }
function rung_(key){ return ladder_().find(x => x.key === key) || ladder_()[ladder_().length-1]; }
function pickRung_(full){
  const L = ladder_(); if(!full) return L[L.length-1].key;
  const taken = {}; rows_('responses').filter(x => x.status === 'submitted').forEach(x => taken[x.tier] = (taken[x.tier]||0) + 1);
  const r = L.find(x => !Number(x.count) || (taken[x.key]||0) < Number(x.count));
  return (r || L[L.length-1]).key;
}
function assignVoucher_(r){
  const rg = rung_(r.tier);
  if(rg.auto){ r.voucher = 'LEISH-' + Utilities.getUuid().replace(/-/g,'').slice(0,5).toUpperCase(); r.rewardStatus = 'Voucher sent'; r.rewardSentAt = now_(); return; }
  const v = rows_('vouchers').find(x => x.tier === r.tier && (!x.status || x.status === 'available'));
  if(v){
    v.status = 'issued'; v.assignedTo = r.email; v.assignedName = r.name; v.assignedAt = now_();
    writeRow_('vouchers', v, v._row);
    r.voucher = v.code; r.rewardStatus = 'Voucher sent'; r.rewardSentAt = now_();
  } else {
    r.rewardStatus = 'Unlocked, voucher to follow';
  }
}

function emailReward_(r){
  if(!r.email) return;
  const s = settings_();
  const dir = s.directorateName || 'Directorate of Labour, Employment Services, Industrial Relations and OSH';
  const rg = rung_(r.tier), reward = rg.name;
  const first = String(r.name||'').split(' ')[0] || 'colleague';
  let line;
  if(r.tier === 'none') return;
  if(r.rewardStatus === 'Voucher sent') line = `<p>Your reward: <b>${reward}</b>.</p><p style="font-size:22px;letter-spacing:2px;font-family:monospace;background:#F3EEDD;padding:12px 16px;border-radius:8px;display:inline-block">${r.voucher}</p><p>${rg.howTo || ''}</p>`;
  else if(r.rewardStatus === 'Awaiting approval') line = `<p>You have unlocked <b>${reward}</b>. Because you used a personal email address, the Malaika Path team will confirm you are part of the Directorate and then send your voucher here.</p>`;
  else line = `<p>You have unlocked <b>${reward}</b>. Your voucher code is being prepared and will reach this inbox shortly.</p>`;
  const more = '';
  const html = `<div style="font-family:Arial,sans-serif;max-width:560px;color:#1d2a26;line-height:1.55">
    <p style="font-size:12px;letter-spacing:1px;text-transform:uppercase;color:#6B5F57"><b>${s.shortName || 'LEISH'}</b> · ${dir} · powered by Malaika Path</p>
    <h2 style="margin:6px 0 12px">Thank you, ${first}. You have unlocked a reward.</h2>
    <p>We received your answers to <i>Listening to Our Team</i>. What you shared will help shape a better working environment for everyone in the Directorate, and better service for the workers and employers who come to us.</p>
    ${line}${more}
    <p>Please tell a colleague: everyone who takes part finds a thank-you at the end.</p>
    <p style="background:#F6F4F0;padding:10px 12px;border-radius:8px;font-size:13px">Please do not reply to this email; this inbox is not monitored. For any question, call the ${s.contactName || 'Malaika Path project coordinator'} on <b>${s.contactPhone || '0772 550 134'}</b>.</p>
    <p style="color:#666;font-size:13px">Your answers are confidential and are used only for this improvement programme.<br>Malaika Path · Kalema Close, Bukasa Road, Kampala</p></div>`;
  try{
    MailApp.sendEmail({to:r.email, subject:'You unlocked a reward · Listening to Our Team', htmlBody:html, name:'Malaika Path (no reply)'});
    log_('system','email', r.email + ' · ' + r.rewardStatus);
  }catch(e){ log_('system','email failed', r.email + ' · ' + e.message); }
}

/* ------------------------------------------------------------------ Malaika Path API */
function apiAdmin(key){
  isAdmin_(key);
  log_('admin','viewed dashboard','');
  return { settings: settings_(), responses: rows_('responses').map(r => { delete r._row; delete r.code; return r; }),
           vouchers: rows_('vouchers').map(v => { delete v._row; return v; }), staff: rows_('staff').length,
           audit: rows_('audit').slice(-60).reverse(), sheetUrl: ss_().getUrl() };
}
function apiAdminSaveSettings(key, obj){
  isAdmin_(key);
  return lock_(() => saveSettings_(obj));
}
function saveSettings_(obj){
  {
    const sh = sh_('settings'); const existing = rows_('settings');
    Object.keys(obj).forEach(k => {
      const val = JSON.stringify(obj[k]);
      if(val.length > 49000) throw new Error('"' + k + '" is too large to save. For a logo, use a smaller image.');
      const row = existing.find(r => r.key === k);
      if(row) sh.getRange(row._row, 2).setValue(val); else sh.appendRow([k, val]);
    });
    log_('admin','saved settings', Object.keys(obj).join(', '));
    return true;
  }
}
function apiAdminApprove(key, id){
  isAdmin_(key);
  const r = lock_(() => {
    const r = rows_('responses').find(x => x.id === id);
    if(!r || r.status !== 'submitted') throw new Error('Response not found.');
    r.verified = 'approved'; assignVoucher_(r); writeRow_('responses', r, r._row);
    log_('admin','approved reward', r.email + ' · ' + r.rewardStatus); return r;
  });
  emailReward_(r); return r.rewardStatus;
}
function apiAdminSetStatus(key, id, status){
  isAdmin_(key);
  return lock_(() => { const r = rows_('responses').find(x => x.id === id); if(!r) throw new Error('Not found');
    r.rewardStatus = status; writeRow_('responses', r, r._row); log_('admin','set reward status', r.email+' · '+status); return true; });
}
function apiAdminResend(key, id){
  isAdmin_(key);
  const r = rows_('responses').find(x => x.id === id); if(!r) throw new Error('Not found');
  emailReward_(r); return true;
}
/** Adds voucher codes (one per line) and gives them straight to anyone already waiting. */
function apiAdminAddVouchers(key, tier, text){
  isAdmin_(key);
  const codes = String(text||'').split(/[\n,]+/).map(s => s.trim()).filter(Boolean);
  const sent = [];
  lock_(() => {
    const have = new Set(rows_('vouchers').map(v => String(v.code)));
    codes.forEach(c => { if(!have.has(c)) writeRow_('vouchers', {code:"'" + c, tier, status:'available'}); });
    rows_('responses').filter(r => r.status === 'submitted' && r.tier === tier && r.rewardStatus === 'Unlocked, voucher to follow')
      .forEach(r => { assignVoucher_(r); if(r.voucher){ writeRow_('responses', r, r._row); sent.push(r); } });
    log_('admin','added vouchers', codes.length + ' ' + tier + ' codes; ' + sent.length + ' sent to waiting people');
  });
  sent.forEach(emailReward_);
  return {added: codes.length, sent: sent.length};
}

/** Testing is over: switch rewards on. Anyone who already finished gets a reward, in the order they finished. */
function apiAdminRelease(key){
  isAdmin_(key);
  const out = lock_(() => {
    saveSettings_({rewardsLive: true});
    const done = [];
    rows_('responses').filter(r => r.status === 'submitted' && r.rewardStatus === 'Rewards not open yet')
      .sort((a,b) => String(a.updated).localeCompare(String(b.updated)))
      .forEach(r => {
        r.tier = pickRung_(Number(r.depth) >= Number(setting_('fullThreshold', 75)));
        if(r.verified !== 'yes' && setting_('approveUnverified', true) !== false) r.rewardStatus = 'Awaiting approval'; else assignVoucher_(r);
        writeRow_('responses', r, r._row); done.push(r);
      });
    log_('admin','rewards switched on', done.length + ' early finishers rewarded');
    return done;
  });
  out.forEach(emailReward_);
  return {released: out.length};
}
/** Removes one response for good, e.g. a test. */
function apiAdminDelete(key, id){
  isAdmin_(key);
  return lock_(() => { const r = rows_('responses').find(x => x.id === id); if(!r) throw new Error('Not found');
    sh_('responses').deleteRow(r._row); log_('admin','deleted response', r.email + ' · ' + r.name); return true; });
}
