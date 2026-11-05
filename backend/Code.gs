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
  const first = String(r.name||'').split(' ')[0] || 'colleague';
  const body = `<p style="margin:0 0 14px">Dear ${first},</p>
    <p style="margin:0 0 14px">Here is your personal code. Use it with your email to pick up where you left off.</p>
    <table role="presentation" cellpadding="0" cellspacing="0" style="margin:8px 0 18px"><tr><td style="background:#F7ECD3;border-radius:10px;padding:16px 26px;font-family:'Courier New',monospace;font-size:32px;letter-spacing:8px;color:#5A3D05;font-weight:bold">${r.code}</td></tr></table>
    <p style="margin:0 0 18px">Speak freely. There are no wrong answers.</p>
    <p style="margin:0 0 6px"><a href="${PAGE_URL}" style="display:inline-block;background:#B4441B;color:#ffffff;text-decoration:none;font-family:Arial,sans-serif;font-weight:bold;padding:13px 22px;border-radius:10px">Continue the conversation</a></p>`;
  sendMail_(r.email, 'Your personal code · Listening to Our Team', emailShell_({kicker:'Listening to Our Team', title:'Your voice, in your own words', standfirst:'Six questions. Fifteen minutes. Speak or type.', body}));
}
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
  {key:'cj', name:'Café Javas breakfast voucher', count:5, howTo:'Show this code at any Café Javas branch to redeem it.'},
  {key:'endiro', name:'Endiro breakfast combo', count:10, howTo:'Show this code at any Endiro Coffee branch to redeem it.'},
  {key:'office', name:'Freedom voice and data voucher (UGX 8,000)', count:35, howTo:'Load this voucher code on your phone using the instructions in your email.'}
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
  if(!r.email || !r.tier || r.tier === 'none') return;
  const rg = rung_(r.tier), reward = rg.name, first = String(r.name||'').split(' ')[0] || 'colleague';
  let line;
  if(r.rewardStatus === 'Voucher sent') line = `<p style="margin:0 0 10px">Your thank-you: <b>${reward}</b>.</p>
      <table role="presentation" cellpadding="0" cellspacing="0" style="margin:6px 0 12px"><tr><td style="background:#F7ECD3;border:2px dashed #C98A12;border-radius:12px;padding:16px 24px;font-family:'Courier New',monospace;font-size:26px;letter-spacing:5px;color:#5A3D05;font-weight:bold">${r.voucher}</td></tr></table>
      <p style="margin:0 0 14px;font-size:14px;color:#6B5F57">${rg.howTo || ''}</p>`;
  else if(r.rewardStatus === 'Awaiting approval') line = `<p style="margin:0 0 14px">You have unlocked <b>${reward}</b>. We will confirm your work details, then send your voucher here.</p>`;
  else line = `<p style="margin:0 0 14px">You have unlocked <b>${reward}</b>. Your voucher will reach this inbox shortly.</p>`;
  const body = `<p style="margin:0 0 14px">Dear ${first},</p>
    <p style="margin:0 0 14px">Your answers are in. Thank you for your time and honesty.</p>
    ${line}
    <p style="margin:0 0 6px">Tell a colleague: there is a thank-you for everyone who takes part.</p>`;
  sendMail_(r.email, 'A thank-you for you · Listening to Our Team', emailShell_({kicker:'With our thanks', title:'We heard you, ' + first + '.', standfirst:'Your voice now shapes what comes next.', body}));
}
/* ------------------------------------------------------------------ emails: Malaika Path magazine layout */
const PAGE_URL = 'https://rebekahzadokmugisha-boop.github.io/lei-listening/';
const LEISH_FULL = 'LEISH (Labour, Employment Services, Industrial Relations and OSH)';
function emailShell_(o){
  const s = settings_(), ink = '#231A15', muted = '#6B5F57', brand = '#B4441B', gold = '#C98A12', cream = '#F7F1E8';
  const phone = s.contactPhone || '0772 550 134', who = s.contactName || 'Malaika Path project coordinator';
  return `<!doctype html><html><body style="margin:0;padding:0;background:${cream}">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:${cream}"><tr><td align="center" style="padding:24px 12px">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:14px;overflow:hidden;font-family:Georgia,'Times New Roman',serif;color:${ink}">
  <tr><td style="background:#2B1A13;padding:18px 28px">
    <table role="presentation" width="100%"><tr>
      <td style="vertical-align:middle"><img src="${PAGE_URL}crest.png" width="40" height="43" alt="Coat of arms of Uganda" style="display:block"></td>
      <td style="vertical-align:middle;padding-left:12px;font-family:Arial,sans-serif;color:#F6EEE7;font-size:12px;line-height:1.4"><b style="font-size:14px">${LEISH_FULL}</b><br>Ministry of Gender, Labour and Social Development</td>
      <td align="right" style="vertical-align:middle;font-family:Arial,sans-serif;color:#C9B3A4;font-size:10px;letter-spacing:1.5px;text-transform:uppercase">in partnership with<br><img src="${PAGE_URL}logo.png" width="54" height="54" alt="Malaika Path" style="display:inline-block;margin-top:6px;background:#fff;border-radius:8px"></td>
    </tr></table></td></tr>
  <tr><td style="height:5px;background:linear-gradient(90deg,#111 0 33%,#F2C300 33% 66%,#D21F26 66%);font-size:0">&nbsp;</td></tr>
  <tr><td style="padding:34px 32px 10px">
    <p style="margin:0;font-family:Arial,sans-serif;font-size:11px;letter-spacing:2.5px;text-transform:uppercase;color:${brand};font-weight:bold">${o.kicker}</p>
    <h1 style="margin:10px 0 14px;font-size:34px;line-height:1.12;font-weight:normal">${o.title}</h1>
    <p style="margin:0;font-size:18px;line-height:1.55;color:${muted};font-style:italic">${o.standfirst}</p>
  </td></tr>
  <tr><td style="padding:22px 32px 8px;font-size:16px;line-height:1.7">${o.body}</td></tr>
  <tr><td style="padding:10px 32px 6px">
    <table role="presentation" width="100%" style="border-top:2px solid ${ink}"><tr><td style="padding-top:18px">
      <p style="margin:0;font-family:Arial,sans-serif;font-size:11px;letter-spacing:2.5px;text-transform:uppercase;color:${gold};font-weight:bold">The programme</p>
      <h2 style="margin:8px 0 8px;font-size:21px;font-weight:normal">Listening first, then building together</h2>
      <p style="margin:0 0 8px;font-size:15px;line-height:1.6">The Minister’s vision, delivered with Malaika Path: better tools, support and recognition for every officer. It starts with your voice.</p>
    </td></tr></table>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:8px 0 6px;font-family:Arial,sans-serif"><tr>
      ${[['1','Listen','Your words'],['2','Map','What gets in the way'],['3','Design','Better ways of working'],['4','Build','Better days, better service']].map(x => `<td width="25%" style="vertical-align:top;padding:10px 6px 10px 0"><div style="font-size:22px;color:${brand};font-family:Georgia,serif">${x[0]}</div><div style="font-size:13px;font-weight:bold;margin-top:2px">${x[1]}</div><div style="font-size:12px;color:${muted};line-height:1.45;margin-top:2px">${x[2]}</div></td>`).join('')}
    </tr></table>
  </td></tr>
  <tr><td style="padding:14px 32px 26px">
    <table role="presentation" width="100%" style="background:${cream};border-left:4px solid ${brand};border-radius:6px"><tr><td style="padding:16px 18px;font-size:19px;line-height:1.5;font-style:italic">“Listen first. Act next. Speak last.”<br><span style="font-family:Arial,sans-serif;font-style:normal;font-size:12px;color:${muted};letter-spacing:1px;text-transform:uppercase">The Malaika Path way</span></td></tr></table>
  </td></tr>
  <tr><td style="background:#2B1A13;padding:22px 32px;font-family:Arial,sans-serif;color:#C9B3A4;font-size:12px;line-height:1.6">
    <b style="color:#F6EEE7;font-size:13px">Malaika Path</b> · <i>Empowering Africa, one story at a time.</i><br>
    No reply needed; this inbox is not monitored. Questions? Call <b style="color:#F6EEE7">${phone}</b>.<br>
    Your answers are confidential.
  </td></tr>
</table></td></tr></table></body></html>`;
}
function sendMail_(to, subject, html){
  try{ MailApp.sendEmail({to, subject, htmlBody:html, name:'Malaika Path (no reply)'}); log_('system','email', to + ' · ' + subject); }
  catch(e){ log_('system','email failed', to + ' · ' + e.message); }
}

/* ------------------------------------------------------------------ invitation email (drafts only; a person clicks Send) */
function inviteHtml_(test){
  const ink = '#231A15', muted = '#6B5F57', brand = '#B4441B', gold = '#C98A12', cream = '#F7F1E8';
  const steps = [['1','Listen','Your words'],['2','Map','What gets in the way'],['3','Design','Better ways of working'],['4','Build','Better days, better service']].map(x => `<td width="25%" style="vertical-align:top;padding:10px 6px 10px 0"><div style="font-size:22px;color:${brand};font-family:Georgia,serif">${x[0]}</div><div style="font-size:13px;font-weight:bold;margin-top:2px">${x[1]}</div><div style="font-size:12px;color:${muted};line-height:1.45;margin-top:2px">${x[2]}</div></td>`).join('');
  const li = t => `<tr><td style="vertical-align:top;padding:4px 10px 4px 0;color:${brand};font-family:Arial,sans-serif;font-weight:bold">•</td><td style="padding:4px 0;font-size:15px;line-height:1.55">${t}</td></tr>`;
  const testbar = test ? '<tr><td style="background:#FBEBD5;padding:12px 28px;font-family:Arial,sans-serif;font-size:13px;color:#7A4300;line-height:1.5"><b>TEST EMAIL.</b> Please go through the tool once only. We have five test places before launch, and rewards stay switched off during testing.</td></tr>' : '';
  return `<!doctype html><html><body style="margin:0;padding:0;background:${cream}">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:${cream}"><tr><td align="center" style="padding:24px 12px">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:14px;overflow:hidden;font-family:Georgia,'Times New Roman',serif;color:${ink}">
${testbar}
<tr><td style="background:#2B1A13;padding:18px 28px"><table role="presentation" width="100%"><tr>
<td style="vertical-align:middle;width:44px"><img src="cid:crest" width="40" height="43" alt="Coat of arms of Uganda" style="display:block"></td>
<td style="vertical-align:middle;padding-left:12px;font-family:Arial,sans-serif;color:#F6EEE7;font-size:12px;line-height:1.4"><b style="font-size:14px">${LEISH_FULL}</b><br>Ministry of Gender, Labour and Social Development</td>
<td align="right" style="vertical-align:middle;font-family:Arial,sans-serif;color:#C9B3A4;font-size:10px;letter-spacing:1.5px;text-transform:uppercase">in partnership with<br><img src="cid:logo" width="54" height="54" alt="Malaika Path" style="display:inline-block;margin-top:6px;background:#fff;border-radius:8px"></td>
</tr></table></td></tr>
<tr><td style="height:5px;background:#F2C300;font-size:0;line-height:0"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td width="33%" style="height:5px;background:#111111"></td><td width="34%" style="height:5px;background:#F2C300"></td><td width="33%" style="height:5px;background:#D21F26"></td></tr></table></td></tr>
<tr><td style="padding:34px 32px 8px">
<p style="margin:0;font-family:Arial,sans-serif;font-size:11px;letter-spacing:2.5px;text-transform:uppercase;color:${brand};font-weight:bold">An invitation · Listening to Our Team</p>
<h1 style="margin:10px 0 10px;font-size:32px;line-height:1.12;font-weight:normal">Your work, in your own words</h1>
<p style="margin:0;font-size:17px;line-height:1.5;color:${muted};font-style:italic">Six questions. Fifteen minutes. Speak or type.</p></td></tr>
<tr><td style="padding:20px 32px 6px;font-size:16px;line-height:1.7">
<p style="margin:0 0 14px">Dear colleagues,</p>
<p style="margin:0 0 14px">Following the Honourable Minister’s communication, we invite you to the first step of our improvement programme: a short listening conversation.</p>
<p style="margin:0 0 14px">This is not an assessment, and it is not about jobs. We want to hear how your work really happens and what gets in the way, so we can fix it together.</p>
</td></tr>
<tr><td style="padding:4px 32px 4px"><table role="presentation" width="100%" style="background:${cream};border-radius:10px"><tr><td style="padding:16px 20px">
<p style="margin:0 0 6px;font-family:Arial,sans-serif;font-size:11px;letter-spacing:2.5px;text-transform:uppercase;color:${gold};font-weight:bold">How it works</p>
<table role="presentation" cellpadding="0" cellspacing="0">${li("On your phone or computer, in English or Kiswahili.")}${li("Confidential. Reports show units, never names.")}${li("A personal code by email lets you stop and come back.")}</table>
<p style="margin:12px 0 6px;font-family:Arial,sans-serif;font-size:11px;letter-spacing:2.5px;text-transform:uppercase;color:${gold};font-weight:bold">Simple rules</p>
<table role="presentation" cellpadding="0" cellspacing="0">${li("One conversation per person, with your own email.")}${li("Use your @mglsd.go.ug email if you have one. Personal emails are confirmed first.")}${li("Answer in full. Detail helps most.")}</table>
</td></tr></table></td></tr>
<tr><td style="padding:18px 32px 4px"><table role="presentation" width="100%" style="border:2px dashed ${gold};border-radius:12px;background:#FFFBF2"><tr><td style="padding:16px 20px;font-size:17px;line-height:1.5"><span style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:2.5px;text-transform:uppercase;color:${gold};font-weight:bold">A little surprise</span><br>There is a reward waiting for the first people to finish. <i>How many? We’re not saying.</i> Don’t wait.</td></tr></table></td></tr>
<tr><td align="center" style="padding:22px 32px 10px"><a href="${PAGE_URL}" style="display:inline-block;background:${brand};color:#ffffff;text-decoration:none;font-family:Arial,sans-serif;font-weight:bold;font-size:16px;padding:14px 28px;border-radius:10px">Start the conversation</a></td></tr>
<tr><td style="padding:12px 32px 6px"><table role="presentation" width="100%" style="border-top:2px solid ${ink}"><tr><td style="padding-top:16px">
<p style="margin:0;font-family:Arial,sans-serif;font-size:11px;letter-spacing:2.5px;text-transform:uppercase;color:${gold};font-weight:bold">The programme</p>
<h2 style="margin:8px 0 8px;font-size:21px;font-weight:normal">Listening first, then building together</h2>
<p style="margin:0 0 8px;font-size:15px;line-height:1.6">The Minister’s vision, delivered with Malaika Path: better tools, support and recognition for every officer. It starts with your voice.</p>
</td></tr></table>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:6px 0 6px;font-family:Arial,sans-serif"><tr>${steps}</tr></table></td></tr>
<tr><td style="padding:12px 32px 24px"><table role="presentation" width="100%" style="background:${cream};border-left:4px solid ${brand};border-radius:6px"><tr><td style="padding:16px 18px;font-size:19px;line-height:1.5;font-style:italic">“Listen first. Act next. Speak last.”<br><span style="font-family:Arial,sans-serif;font-style:normal;font-size:12px;color:${muted};letter-spacing:1px;text-transform:uppercase">The Malaika Path way</span></td></tr></table></td></tr>
<tr><td style="padding:0 32px 26px;font-size:15px;line-height:1.6">With appreciation,<br><b>Rebekah Zadok Mugisha</b><br><span style="color:${muted}">Lead Consultant, Malaika Path</span></td></tr>
<tr><td style="background:#2B1A13;padding:20px 32px;font-family:Arial,sans-serif;color:#C9B3A4;font-size:12px;line-height:1.6">
<b style="color:#F6EEE7;font-size:13px">Malaika Path</b> · <i>Empowering Africa, one story at a time.</i><br>
${LEISH_FULL}, in partnership with Malaika Path.<br>Questions? Call the project coordinator on <b style="color:#F6EEE7">0772 550 134</b>.</td></tr>
</table></td></tr></table></body></html>`;
}
/** Run from the editor: puts the TEST invitation in the Malaika Path Drafts folder. Nothing is sent. */
function createTestInviteDraft(){
  const imgs = { crest: UrlFetchApp.fetch(PAGE_URL + 'crest.png').getBlob().setName('crest.png'), logo: UrlFetchApp.fetch(PAGE_URL + 'logo.png').getBlob().setName('malaika-path.png') };
  const to = 'rebekahzadokmugisha@gmail.com,dru.nkyeswa@gmail.com,loyoenock@gmail.com,mlawgeorge@gmail.com';
  GmailApp.createDraft(to, 'TEST · Your voice first: Listening to Our Team', 'Please view this email in HTML.', {htmlBody: inviteHtml_(true), inlineImages: imgs, name: 'Malaika Path'});
  Logger.log('Draft created in Malaika Path Drafts for: ' + to);
}
/** Run from the editor when ready: the invitation for the PS, without the test banner. Nothing is sent. */
function createPSInviteDraft(){
  const imgs = { crest: UrlFetchApp.fetch(PAGE_URL + 'crest.png').getBlob().setName('crest.png'), logo: UrlFetchApp.fetch(PAGE_URL + 'logo.png').getBlob().setName('malaika-path.png') };
  GmailApp.createDraft('', 'Your voice first: Listening to Our Team', 'Please view this email in HTML.', {htmlBody: inviteHtml_(false), inlineImages: imgs, name: 'Malaika Path'});
  Logger.log('PS invitation draft created in Malaika Path Drafts (add the PS address before sending).');
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
