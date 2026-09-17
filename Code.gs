/**
 * DAILY REPORT PRO FINAL - COMPANY EDITION
 * Google Apps Script + Google Sheets + Google Drive
 * Roles: USER, ADMIN, SUPER_ADMIN
 */

const CONFIG = {
  APP_NAME: 'Daily Report Pro',
  COMPANY_NAME: 'Chicken Crush',
  SPREADSHEET_ID: '1Gi_C_ctKUz_kLnQxdfMnmZrCpJ0jwtXRoQAOhwFJceA',
  SUPER_ADMIN_EMAIL: 'chickencrush.support@gmail.com',
  TIMEZONE: 'Asia/Jakarta',
  DRIVE_FOLDER_ID: '',
  SHEETS: {
    REPORTS: 'REPORTS', USERS: 'USERS', DIVISIONS: 'DIVISIONS',
    SETTINGS: 'SETTINGS', AUDIT: 'AUDIT_LOG', SESSIONS: 'SESSIONS'
  }
};

const STATUS = ['Selesai', 'Proses', 'Hold'];
const ROLES = ['USER', 'ADMIN', 'SUPER_ADMIN'];

function doGet(e) {
  var page = HtmlService.createHtmlOutputFromFile('Index');
  if (!/<html[\s>]/i.test(page.getContent())) {
    return HtmlService.createHtmlOutput('<h2>Periksa file Index.html</h2><p>Isi Index.html harus berasal dari file Index.html dalam paket, dimulai dengan &lt;!DOCTYPE html&gt;. Isi Code.gs ditempel ke file Script.</p>');
  }
  return page
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupApp() {
  var temporarySuperAdminPassword = '';
  const ss = getDb_();
  ensureSheet_(ss, CONFIG.SHEETS.REPORTS, [
    'ID','TIMESTAMP','DATE','USER_EMAIL','USER_NAME','DIVISION','ROLE','STATUS',
    'ACTIVITY','DEADLINE','DRIVE_URL','NOTES','COMPLETED_AT','UPDATED_AT'
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.USERS, ['EMAIL','NAME','ROLE','DIVISION','ACTIVE','CREATED_AT','UPDATED_AT','SALT','PASSWORD_HASH']);
  ensureSheet_(ss, CONFIG.SHEETS.SESSIONS, ['TOKEN_HASH','EMAIL','CREATED_AT','EXPIRES_AT','ACTIVE']);
  ensureSheet_(ss, CONFIG.SHEETS.DIVISIONS, ['ID','NAME','ACTIVE','CREATED_AT']);
  ensureSheet_(ss, CONFIG.SHEETS.SETTINGS, ['KEY','VALUE']);
  ensureSheet_(ss, CONFIG.SHEETS.AUDIT, ['TIMESTAMP','ACTOR_EMAIL','ACTOR_NAME','ACTION','ENTITY','ENTITY_ID','DETAIL']);
  ensureUserHeaders_(ss.getSheetByName(CONFIG.SHEETS.USERS));

  const email = String(CONFIG.SUPER_ADMIN_EMAIL || '').trim().toLowerCase();
  const userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  if (email && email !== 'admin@yourcompany.com') {
    const users = getSheetObjects_(CONFIG.SHEETS.USERS);
    var existingSuper = users.find(function(u){return String(u.EMAIL||'').toLowerCase()===email;});
    if (existingSuper && !String(existingSuper.PASSWORD_HASH||'')) {
      var all= userSheet.getDataRange().getValues(), hh=all[0], ee=hh.indexOf('EMAIL'), saltI=hh.indexOf('SALT'), hashI=hh.indexOf('PASSWORD_HASH');
      for(var ui=1;ui<all.length;ui++){if(String(all[ui][ee]||'').toLowerCase()===email){var seedExisting=makePassword_('ChangeMe123!');temporarySuperAdminPassword='ChangeMe123!';userSheet.getRange(ui+1,saltI+1).setValue(seedExisting.salt);userSheet.getRange(ui+1,hashI+1).setValue(seedExisting.hash);break;}}
    }
    if (!users.some(u => String(u.EMAIL).toLowerCase() === email)) {
      var seed = makePassword_('ChangeMe123!');temporarySuperAdminPassword='ChangeMe123!';
      userSheet.appendRow([email, 'Super Admin', 'SUPER_ADMIN', '', true, new Date(), new Date(), seed.salt, seed.hash]);
    }
  }
  const divSheet = ss.getSheetByName(CONFIG.SHEETS.DIVISIONS);
  if (divSheet.getLastRow() === 1) {
    divSheet.getRange(2,1,4,4).setValues([
      ['DIV-001','BPO',true,new Date()],['DIV-002','Marketing',true,new Date()],
      ['DIV-003','Operasional',true,new Date()],['DIV-004','Finance',true,new Date()]
    ]);
  }
  const settings = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  const existing = getSheetObjects_(CONFIG.SHEETS.SETTINGS);
  const defaults = [
    ['APP_NAME', CONFIG.APP_NAME], ['COMPANY_NAME', CONFIG.COMPANY_NAME],
    ['TIMEZONE', CONFIG.TIMEZONE], ['VERSION', '3.0'], ['SETUP_AT', new Date()]
  ];
  defaults.forEach(([k,v]) => { if (!existing.some(x => String(x.KEY) === k)) settings.appendRow([k,v]); });
  styleSheet_(ss.getSheetByName(CONFIG.SHEETS.REPORTS));
  styleSheet_(ss.getSheetByName(CONFIG.SHEETS.USERS));
  styleSheet_(ss.getSheetByName(CONFIG.SHEETS.DIVISIONS));
  styleSheet_(ss.getSheetByName(CONFIG.SHEETS.AUDIT));
  audit_('SETUP','SYSTEM','', 'Application initialized / upgraded to V3');
  return {ok:true, message:'Daily Report Pro V3 berhasil disiapkan.'+(temporarySuperAdminPassword?' Password sementara Super Admin: '+temporarySuperAdminPassword:'')};
}

function registerUser(payload) {
  payload = payload || {};
  ensureCoreSheets_();
  var email = clean_(payload.email).toLowerCase(), name = clean_(payload.name), division = clean_(payload.division), password = String(payload.password || '');
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email / ID tidak valid.');
  if (!name || name.length < 2) throw new Error('Nama wajib diisi.');
  if (!division) throw new Error('Divisi wajib dipilih.');
  var validDivisions = getAllDivisions_(false);
  if (!validDivisions.some(function(d){ return d.name.toLowerCase() === division.toLowerCase(); })) throw new Error('Divisi tidak tersedia atau sudah dinonaktifkan. Silakan pilih divisi yang tersedia.');
  division = validDivisions.find(function(d){ return d.name.toLowerCase() === division.toLowerCase(); }).name;
  validatePassword_(password);
  var sh=getDb_().getSheetByName(CONFIG.SHEETS.USERS), users=getSheetObjects_(CONFIG.SHEETS.USERS);
  if (users.some(function(u){return String(u.EMAIL||'').toLowerCase()===email;})) throw new Error('Akun sudah terdaftar. Silakan login.');
  var role = email === String(CONFIG.SUPER_ADMIN_EMAIL||'').trim().toLowerCase() ? 'SUPER_ADMIN' : 'USER';
  var now=new Date(), pw=makePassword_(password);
  sh.appendRow([email,name,role,division,true,now,now,pw.salt,pw.hash]);
  audit_('REGISTER','USER',email,name+' | '+division);
  return loginUser({email:email,password:password,remember:true});
}

function loginUser(payload) {
  ensureCoreSheets_();
  payload=payload||{}; var email=clean_(payload.email).toLowerCase(), password=String(payload.password||'');
  if(!email||!password) throw new Error('ID dan password wajib diisi.');
  var users=getSheetObjects_(CONFIG.SHEETS.USERS), found=users.find(function(u){return String(u.EMAIL||'').toLowerCase()===email;});
  if(!found||!bool_(found.ACTIVE)) throw new Error('Akun tidak ditemukan atau nonaktif.');
  if(!verifyPassword_(password,String(found.SALT||''),String(found.PASSWORD_HASH||''))) throw new Error('ID atau password salah.');
  var token=issueSession_(email, payload.remember !== false);
  var user=userFromRow_(found);
  cacheUser_(token,user); setCurrentUser_(user);
  return {ok:true,token:token,user:user,expiresAt:sessionExpiry_(payload.remember !== false)};
}

function restoreSession(token) {
  var user=requireUser_(token);
  cacheUser_(token,user); return {ok:true,user:user,expiresAt:sessionExpiry_(true)};
}

function logoutUser(token) { if(token) revokeSession_(token); return {ok:true}; }

function changeMyPassword(payload,token){
  var user=requireUser_(token), p=payload||{}, old=String(p.currentPassword||''), next=String(p.newPassword||'');
  validatePassword_(next);
  var users=getSheetObjects_(CONFIG.SHEETS.USERS), found=users.find(function(u){return String(u.EMAIL||'').toLowerCase()===user.email;});
  if(!found||!verifyPassword_(old,String(found.SALT||''),String(found.PASSWORD_HASH||'')))throw new Error('Password saat ini salah.');
  var pw=makePassword_(next), sh=getDb_().getSheetByName(CONFIG.SHEETS.USERS), data=sh.getDataRange().getValues(), h=data[0];
  for(var i=1;i<data.length;i++)if(String(data[i][h.indexOf('EMAIL')]||'').toLowerCase()===user.email){sh.getRange(i+1,h.indexOf('SALT')+1).setValue(pw.salt);sh.getRange(i+1,h.indexOf('PASSWORD_HASH')+1).setValue(pw.hash);break;}
  audit_('CHANGE_PASSWORD','USER',user.email,'Password changed'); return {ok:true};
}

function getPublicDivisions() {
  ensureCoreSheets_();
  var rows = getAllDivisions_(false);
  if (!rows.length) { seedDefaultDivisions_(); rows = getAllDivisions_(false); }
  return rows;
}

function hashText_(text) { var d=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,text,Utilities.Charset.UTF_8); return Utilities.base64EncodeWebSafe(d); }
function makePassword_(password) { var salt=Utilities.getUuid().replace(/-/g,''); return {salt:salt,hash:hashText_(salt+'|'+password)}; }
function verifyPassword_(password,salt,hash) { return !!salt && !!hash && hashText_(salt+'|'+password)===hash; }
function validatePassword_(p) { if(p.length<8) throw new Error('Password minimal 8 karakter.'); if(!/[A-Za-z]/.test(p)||!/[0-9]/.test(p)) throw new Error('Password harus mengandung huruf dan angka.'); }
function issueSession_(email,remember) { var raw=Utilities.getUuid()+'-'+Utilities.getUuid(), now=new Date(), days=remember?30:1, exp=new Date(now.getTime()+days*86400000), sh=getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS); sh.appendRow([hashText_(raw),email,now,exp,true]); return raw; }
function sessionExpiry_(remember) { return new Date(new Date().getTime()+(remember?30:1)*86400000).toISOString(); }
function revokeSession_(token) { try{CacheService.getScriptCache().remove('sess_'+hashText_(String(token)));}catch(e){} var sh=getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS),data=sh.getDataRange().getValues(),h=data[0],ti=h.indexOf('TOKEN_HASH'); var th=hashText_(String(token)); for(var i=1;i<data.length;i++){if(String(data[i][ti])===th){sh.getRange(i+1,h.indexOf('ACTIVE')+1).setValue(false);break;}} }
function userFromRow_(u) { return {email:String(u.EMAIL||'').toLowerCase(),name:String(u.NAME||u.EMAIL||''),role:String(u.ROLE||'USER'),division:String(u.DIVISION||'')}; }
function ensureUserHeaders_(sh) { var expected=['EMAIL','NAME','ROLE','DIVISION','ACTIVE','CREATED_AT','UPDATED_AT','SALT','PASSWORD_HASH']; var headers=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),expected.length)).getValues()[0]; expected.forEach(function(x,i){if(headers[i]!==x)sh.getRange(1,i+1).setValue(x);}); }

function getBootstrapData(token) {
  const user = requireUser_(token);
  return {
    appName: CONFIG.APP_NAME, companyName: CONFIG.COMPANY_NAME,
    user, divisions: listVisibleDivisions_(user), today: today_(),
    dashboard: getDashboardData({period:'month'}, token),
    features: {isAdmin:user.role !== 'USER', isSuperAdmin:user.role === 'SUPER_ADMIN'}
  };
}

function saveReport(payload, token) {
  const user = requireUser_(token); payload = payload || {};
  const activity = clean_(payload.activity), date = clean_(payload.date) || today_();
  const deadline = clean_(payload.deadline), status = 'Hold';
  if (!activity) throw new Error('Aktivitas pekerjaan wajib diisi.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal laporan tidak valid.');
  if (deadline && !/^\d{4}-\d{2}-\d{2}$/.test(deadline)) throw new Error('Deadline tidak valid.');
  if (STATUS.indexOf(status) < 0) throw new Error('Status pekerjaan tidak valid.');
  if (deadline && deadline < date) throw new Error('Deadline tidak boleh sebelum tanggal laporan.');
  const id = 'RPT-' + Utilities.getUuid().slice(0,8).toUpperCase();
  const now = new Date(), completedAt = status === 'Selesai' ? now : '';
  let driveUrl = clean_(payload.driveUrl);
  if (payload.file && payload.file.data) driveUrl = saveUploadedFile_(payload.file, user);
  ensureEvidenceColumn_();
  getDb_().getSheetByName(CONFIG.SHEETS.REPORTS).appendRow([
    id, now, date, user.email, user.name, user.division, user.role, status,
    activity, deadline, driveUrl, clean_(payload.notes), completedAt, now
  ]);
  audit_('CREATE_REPORT','REPORT',id,activity + ' | ' + status);
  return {ok:true,id,message:'Laporan berhasil disimpan.'};
}

function ensureEvidenceColumn_(){
  var sh=getDb_().getSheetByName(CONFIG.SHEETS.REPORTS),headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  if(headers.indexOf('EVIDENCE_URL')<0)sh.getRange(1,headers.length+1).setValue('EVIDENCE_URL');
}
function validateEvidence_(url){
  url=clean_(url);
  if(!/^https:\/\/drive\.google\.com\/(?:file\/d\/[A-Za-z0-9_-]+(?:[/?#]|$)|(?:open|uc)\?[^#]*\bid=[A-Za-z0-9_-]+)/i.test(url))throw new Error('Masukkan link file foto Google Drive yang valid, bukan link folder.');
  return url;
}
function updateReportStatus(id, payload, token) {
  var user=requireUser_(token);id=clean_(id);payload=payload||{};
  var status=clean_(payload.status),evidence=clean_(payload.evidenceUrl);
  if(!id||STATUS.indexOf(status)<0)throw new Error('Data status tidak valid.');
  var lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    ensureEvidenceColumn_();
    var sh=getDb_().getSheetByName(CONFIG.SHEETS.REPORTS),data=sh.getDataRange().getValues(),h=data[0],ix=k=>h.indexOf(k);
    for(var i=1;i<data.length;i++)if(String(data[i][ix('ID')])===id){
      authorizeReportAccess_(user,String(data[i][ix('USER_EMAIL')]).toLowerCase(),String(data[i][ix('DIVISION')]));
      var current=String(data[i][ix('STATUS')]),next=current==='Hold'?'Proses':current==='Proses'?'Selesai':'';
      if(status!==next)throw new Error('Status hanya dapat maju: Hold → Proses → Selesai. Muat ulang laporan.');
      if(status==='Selesai')evidence=validateEvidence_(evidence);
      var now=new Date(),row=data[i].slice();
      row[ix('STATUS')]=status;row[ix('UPDATED_AT')]=now;
      if(status==='Selesai'){row[ix('EVIDENCE_URL')]=evidence;row[ix('COMPLETED_AT')]=now;}
      sh.getRange(i+1,1,1,h.length).setValues([row]);
      audit_('UPDATE_STATUS','REPORT',id,current+' → '+status);
      return {ok:true,message:'Status diperbarui.'};
    }
    throw new Error('Laporan tidak ditemukan.');
  }finally{lock.releaseLock();}
}

function addReportNote(id,note,token){
  const user=requireUser_(token); note=clean_(note); if(!note)throw new Error('Catatan tidak boleh kosong.');
  const sh=getDb_().getSheetByName(CONFIG.SHEETS.REPORTS),data=sh.getDataRange().getValues(),h=data[0],ix=k=>h.indexOf(k);
  for(let i=1;i<data.length;i++) if(String(data[i][ix('ID')])===clean_(id)){
    authorizeReportAccess_(user,String(data[i][ix('USER_EMAIL')]),String(data[i][ix('DIVISION')]));
    const old=String(data[i][ix('NOTES')]||''), stamp=Utilities.formatDate(new Date(),CONFIG.TIMEZONE,'dd/MM/yyyy HH:mm');
    sh.getRange(i+1,ix('NOTES')+1).setValue(old?old+'\n['+stamp+'] '+note:'['+stamp+'] '+note);
    sh.getRange(i+1,ix('UPDATED_AT')+1).setValue(new Date());
    audit_('ADD_NOTE','REPORT',id,note); return {ok:true,message:'Catatan ditambahkan.'};
  }
  throw new Error('Laporan tidak ditemukan.');
}

function getDashboardData(filters,token){
  const user=requireUser_(token), rows=getReportObjects_().filter(r=>canSeeReport_(user,r));
  const range=periodRange_(filters&&filters.period,filters&&filters.startDate,filters&&filters.endDate);
  const scoped=rows.filter(r=>r.date>=range.start&&r.date<=range.end), today=today_(), todayRows=rows.filter(r=>r.date===today);
  const counts=countStatuses_(scoped), todayCounts=countStatuses_(todayRows);
  const overdue=scoped.filter(r=>r.deadline&&r.deadline<today&&r.status!=='Selesai');
  const byDate={}; scoped.forEach(r=>byDate[r.date]=(byDate[r.date]||0)+1);
  const trend=Object.keys(byDate).sort().map(d=>({date:d,value:byDate[d]}));
  const byDivision={}; scoped.forEach(r=>{const k=r.division||'Tanpa Divisi';byDivision[k]=(byDivision[k]||0)+1;});
  const divisionStats=Object.keys(byDivision).map(k=>({name:k,value:byDivision[k]})).sort((a,b)=>b.value-a.value);
  const byEmployee={}; scoped.forEach(r=>{const k=r.userName||r.userEmail; if(!byEmployee[k])byEmployee[k]={name:k,division:r.division||'-',total:0,done:0,progress:0,hold:0}; byEmployee[k].total++; byEmployee[k][statusKey_(r.status)]++;});
  const employeeStats=Object.keys(byEmployee).map(k=>byEmployee[k]).sort((a,b)=>b.total-a.total).slice(0,30);
  const recent=scoped.slice().sort((a,b)=>(b.timestamp||'').localeCompare(a.timestamp||'')).slice(0,25);
  return {period:range,total:scoped.length,todayTotal:todayRows.length,counts,todayCounts,overdueCount:overdue.length,
    overdue:overdue.slice().sort((a,b)=>a.deadline.localeCompare(b.deadline)).slice(0,10).map(safeReport_),trend,divisionStats,employeeStats,
    recent:recent.map(safeReport_),myTotal:rows.filter(r=>r.userEmail===user.email).length,employeeCount:getVisibleUserCount_(user)};
}

function getReports(filters,token){
  const user=requireUser_(token); filters=filters||{}; const range=periodRange_(filters.period||'month',filters.startDate,filters.endDate),q=clean_(filters.query).toLowerCase();
  return getReportObjects_().filter(r=>canSeeReport_(user,r)).filter(r=>r.date>=range.start&&r.date<=range.end)
    .filter(r=>!filters.status||r.status===filters.status).filter(r=>!filters.division||r.division===filters.division)
    .filter(r=>!q||[r.userName,r.division,r.activity,r.status,r.id].join(' ').toLowerCase().indexOf(q)>=0)
    .sort((a,b)=>(b.date+b.timestamp).localeCompare(a.date+a.timestamp)).map(safeReport_);
}

function getReportDetail(id,token){const user=requireUser_(token),row=getReportObjects_().find(r=>r.id===clean_(id));if(!row)throw new Error('Laporan tidak ditemukan.');if(!canSeeReport_(user,row))throw new Error('Anda tidak memiliki akses ke laporan ini.');return safeReport_(row);}

function exportReportsPdf(filters,token){
  const user=requireUser_(token), rows=getReports(filters||{},token), range=periodRange_(filters&&filters.period||'month',filters&&filters.startDate,filters&&filters.endDate);
  const doc=DocumentApp.create('Daily Report - '+range.start+' - '+range.end),body=doc.getBody();
  body.appendParagraph(CONFIG.COMPANY_NAME).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(CONFIG.APP_NAME+' | Laporan Aktivitas Pekerjaan').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Periode: '+range.start+' s/d '+range.end);
  body.appendParagraph('Dicetak oleh: '+user.name+' | '+Utilities.formatDate(new Date(),CONFIG.TIMEZONE,'dd/MM/yyyy HH:mm'));
  body.appendParagraph('');
  const c=countStatuses_(rows.map(r=>({status:r.status}))); body.appendParagraph('Ringkasan: Total '+rows.length+' | Selesai '+c.Selesai+' | Proses '+c.Proses+' | Hold '+c.Hold);
  const data=[['Tanggal','Karyawan','Divisi','Status','Aktivitas','Deadline','Durasi']]; rows.forEach(r=>data.push([r.date,r.userName,r.division||'-',r.status,r.activity,r.deadline||'-',r.duration||'-']));
  if(rows.length)body.appendTable(data);else body.appendParagraph('Tidak ada data pada periode ini.');
  doc.saveAndClose(); const file=DriveApp.getFileById(doc.getId()),blob=file.getBlob().getAs(MimeType.PDF).setName('Daily_Report_'+range.start+'_'+range.end+'.pdf');
  const result={name:blob.getName(),mimeType:'application/pdf',base64:Utilities.base64Encode(blob.getBytes())};file.setTrashed(true);audit_('EXPORT_PDF','REPORT','',range.start+' to '+range.end);return result;
}

function exportReportsXlsx(filters,token){
  const user=requireUser_(token),rows=getReports(filters||{},token),range=periodRange_(filters&&filters.period||'month',filters&&filters.startDate,filters&&filters.endDate);
  const temp=SpreadsheetApp.create('Daily Report Export '+range.start+' '+range.end),sh=temp.getSheets()[0];sh.setName('Laporan');
  const values=[['Tanggal','Karyawan','Email','Divisi','Status','Aktivitas','Deadline','Durasi','Google Drive','Catatan']];
  rows.forEach(r=>values.push([r.date,r.userName,r.userEmail,r.division||'-',r.status,r.activity,r.deadline||'-',r.duration||'-',r.driveUrl||'',r.notes||'']));
  sh.getRange(1,1,values.length,values[0].length).setValues(values);sh.getRange(1,1,1,values[0].length).setFontWeight('bold');sh.setFrozenRows(1);sh.autoResizeColumns(1,values[0].length);
  SpreadsheetApp.flush();Utilities.sleep(700);
  const url='https://docs.google.com/spreadsheets/d/'+temp.getId()+'/export?format=xlsx';
  const response=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
  if(response.getResponseCode()!==200){DriveApp.getFileById(temp.getId()).setTrashed(true);throw new Error('Export Excel gagal. Coba lagi.');}
  const blob=response.getBlob().setName('Daily_Report_'+range.start+'_'+range.end+'.xlsx');
  const result={name:blob.getName(),mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',base64:Utilities.base64Encode(blob.getBytes())};
  DriveApp.getFileById(temp.getId()).setTrashed(true);audit_('EXPORT_XLSX','REPORT','',range.start+' to '+range.end);return result;
}

function listUsers(token){requireRole_(['SUPER_ADMIN'],token);return getSheetObjects_(CONFIG.SHEETS.USERS).map(u=>({email:String(u.EMAIL||'').toLowerCase(),name:String(u.NAME||''),role:String(u.ROLE||'USER'),division:String(u.DIVISION||''),active:bool_(u.ACTIVE),createdAt:dateTime_(u.CREATED_AT),hasPassword:!!String(u.PASSWORD_HASH||'')}));}
function saveUser(payload,token){
  requireRole_(['SUPER_ADMIN'],token);payload=payload||{};const email=clean_(payload.email).toLowerCase(),name=clean_(payload.name),role=clean_(payload.role).toUpperCase(),division=clean_(payload.division);
  if(!email||!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new Error('Email tidak valid.');if(!name)throw new Error('Nama wajib diisi.');if(ROLES.indexOf(role)<0)throw new Error('Role tidak valid.');if(role!=='SUPER_ADMIN'&&!division)throw new Error('Divisi wajib diisi.');
  const sh=getDb_().getSheetByName(CONFIG.SHEETS.USERS),data=sh.getDataRange().getValues(),h=data[0],ei=h.indexOf('EMAIL');let row=-1;for(let i=1;i<data.length;i++)if(String(data[i][ei]).toLowerCase()===email){row=i+1;break;}
  const now=new Date(),existingRow=row>0?data[row-1]:null; var salt=existingRow?String(existingRow[7]||''):''; var ph=existingRow?String(existingRow[8]||''):''; if(payload.password){validatePassword_(String(payload.password));var pw=makePassword_(String(payload.password));salt=pw.salt;ph=pw.hash;} const values=[email,name,role,role==='SUPER_ADMIN'?'':division,payload.active!==false,row>0?data[row-1][5]:now,now,salt,ph];if(row>0)sh.getRange(row,1,1,9).setValues([values]);else sh.appendRow(values);
  audit_('SAVE_USER','USER',email,name+' | '+role+' | '+division);return{ok:true};
}
function listDivisions(token){requireRole_(['SUPER_ADMIN'],token);return getAllDivisions_(true);}
function saveDivision(payload,token){
  requireRole_(['SUPER_ADMIN'],token);payload=payload||{};const id=clean_(payload.id)||'DIV-'+Utilities.getUuid().slice(0,6).toUpperCase(),name=clean_(payload.name);if(!name)throw new Error('Nama divisi wajib diisi.');
  const sh=getDb_().getSheetByName(CONFIG.SHEETS.DIVISIONS),data=sh.getDataRange().getValues();let row=-1;for(let i=1;i<data.length;i++)if(String(data[i][0])===id){row=i+1;break;}
  const values=[id,name,payload.active!==false,row>0?data[row-1][3]:new Date()];if(row>0)sh.getRange(row,1,1,4).setValues([values]);else sh.appendRow(values);audit_('SAVE_DIVISION','DIVISION',id,name);return{ok:true,id};
}
function getAuditLogs(token){requireRole_(['SUPER_ADMIN'],token);return getSheetObjects_(CONFIG.SHEETS.AUDIT).slice(-200).reverse().map(x=>({timestamp:dateTime_(x.TIMESTAMP),actor:String(x.ACTOR_NAME||x.ACTOR_EMAIL||''),action:String(x.ACTION||''),entity:String(x.ENTITY||''),entityId:String(x.ENTITY_ID||''),detail:String(x.DETAIL||'')}));}
function getCompanySummary(token){requireRole_(['SUPER_ADMIN'],token);return getDashboardData({period:'year'},token);}

function saveUploadedFile_(file,user){const bytes=Utilities.base64Decode(String(file.data).split(',').pop()),blob=Utilities.newBlob(bytes,file.mimeType||MimeType.PLAIN_TEXT,file.name||'lampiran');const folder=CONFIG.DRIVE_FOLDER_ID?DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID):DriveApp.getRootFolder();const f=folder.createFile(blob);f.setDescription('Daily Report attachment - '+user.email);return f.getUrl();}
function setCurrentUser_(u){ CURRENT_USER_=u; }
var CURRENT_USER_=null;
function requireUser_(token){
  token=clean_(token); if(!token) throw new Error('Sesi login tidak ditemukan. Silakan login kembali.');
  var cached=CacheService.getScriptCache().get('sess_'+hashText_(token));
  if(cached){var cu=JSON.parse(cached);setCurrentUser_(cu);return cu;}
  var th=hashText_(token), sh=getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS), data=sh.getDataRange().getValues(), h=data[0], ti=h.indexOf('TOKEN_HASH'), ei=h.indexOf('EMAIL'), xi=h.indexOf('EXPIRES_AT'), ai=h.indexOf('ACTIVE'), now=new Date(), row=null;
  for(var i=1;i<data.length;i++){ if(String(data[i][ti])===th){row=data[i]; break;} }
  if(!row||!bool_(row[ai])||!row[xi]||new Date(row[xi]).getTime()<now.getTime()){ if(row)sh.getRange(data.indexOf(row)+1,ai+1).setValue(false); throw new Error('Sesi sudah berakhir. Silakan login kembali.'); }
  var email=String(row[ei]||'').toLowerCase(), users=getSheetObjects_(CONFIG.SHEETS.USERS), found=users.find(function(u){return String(u.EMAIL||'').toLowerCase()===email;});
  if(!found||!bool_(found.ACTIVE)) throw new Error('Akun tidak aktif.');
  var user=userFromRow_(found); setCurrentUser_(user); cacheUser_(token,user); return user;
}
function cacheUser_(token,user){try{CacheService.getScriptCache().put('sess_'+hashText_(token),JSON.stringify(user),21600);}catch(e){}}
function requireRole_(roles,token){var u=requireUser_(token);if(roles.indexOf(u.role)<0)throw new Error('Akses ditolak.');return u;}
function canSeeReport_(u,r){if(u.role==='SUPER_ADMIN')return true;if(u.role==='ADMIN')return r.division===u.division;return r.userEmail===u.email;}
function authorizeReportAccess_(u,ownerEmail,division){if(u.role==='SUPER_ADMIN'||(u.role==='ADMIN'&&division===u.division)||u.email===String(ownerEmail).toLowerCase())return;throw new Error('Anda tidak memiliki akses ke laporan ini.');}
function listVisibleDivisions_(u){return getAllDivisions_(false).filter(d=>u.role==='SUPER_ADMIN'||d.name===u.division);}
function getAllDivisions_(includeInactive){return getSheetObjects_(CONFIG.SHEETS.DIVISIONS).map(d=>({id:String(d.ID),name:String(d.NAME),active:bool_(d.ACTIVE)})).filter(d=>includeInactive||d.active);}

function ensureCoreSheets_(){
  var ss=getDb_();
  ensureSheet_(ss,CONFIG.SHEETS.USERS,['EMAIL','NAME','ROLE','DIVISION','ACTIVE','CREATED_AT','UPDATED_AT','SALT','PASSWORD_HASH']);
  ensureSheet_(ss,CONFIG.SHEETS.DIVISIONS,['ID','NAME','ACTIVE','CREATED_AT']);
  ensureSheet_(ss,CONFIG.SHEETS.SESSIONS,['TOKEN_HASH','EMAIL','CREATED_AT','EXPIRES_AT','ACTIVE']);
  ensureSheet_(ss,CONFIG.SHEETS.REPORTS,['ID','TIMESTAMP','DATE','USER_EMAIL','USER_NAME','DIVISION','ROLE','STATUS','ACTIVITY','DEADLINE','DRIVE_URL','NOTES','COMPLETED_AT','UPDATED_AT']);
  ensureSheet_(ss,CONFIG.SHEETS.SETTINGS,['KEY','VALUE']);
  ensureSheet_(ss,CONFIG.SHEETS.AUDIT,['TIMESTAMP','ACTOR_EMAIL','ACTOR_NAME','ACTION','ENTITY','ENTITY_ID','DETAIL']);
}
function seedDefaultDivisions_(){
  var sh=getDb_().getSheetByName(CONFIG.SHEETS.DIVISIONS);
  if(sh.getLastRow()===1){
    sh.getRange(2,1,5,4).setValues([
      ['DIV-001','BPO',true,new Date()],
      ['DIV-002','Marketing',true,new Date()],
      ['DIV-003','Operasional',true,new Date()],
      ['DIV-004','Finance',true,new Date()],
      ['DIV-005','IT',true,new Date()]
    ]);
  }
}

function getVisibleUserCount_(u){const users=getSheetObjects_(CONFIG.SHEETS.USERS).filter(x=>bool_(x.ACTIVE));return u.role==='SUPER_ADMIN'?users.length:users.filter(x=>String(x.DIVISION||'')===u.division).length;}
function getReportObjects_(){return getSheetObjects_(CONFIG.SHEETS.REPORTS).map(r=>({id:String(r.ID||''),timestamp:dateTime_(r.TIMESTAMP),date:formatDate_(r.DATE),userEmail:String(r.USER_EMAIL||'').toLowerCase(),userName:String(r.USER_NAME||''),division:String(r.DIVISION||''),role:String(r.ROLE||''),status:String(r.STATUS||'Proses'),activity:String(r.ACTIVITY||''),deadline:formatDate_(r.DEADLINE),driveUrl:String(r.DRIVE_URL||''),evidenceUrl:String(r.EVIDENCE_URL||''),notes:String(r.NOTES||''),completedAt:dateTime_(r.COMPLETED_AT),updatedAt:dateTime_(r.UPDATED_AT)}));}
function safeReport_(r){const now=new Date();let duration='';if(r.completedAt&&r.timestamp)duration=durationText_(new Date(r.completedAt)-new Date(r.timestamp));else if(r.timestamp)duration=durationText_(now-new Date(r.timestamp));return Object.assign({},r,{duration});}
function countStatuses_(rows){return{Selesai:rows.filter(r=>r.status==='Selesai').length,Proses:rows.filter(r=>r.status==='Proses').length,Hold:rows.filter(r=>r.status==='Hold').length};}
function statusKey_(s){return s==='Selesai'?'done':s==='Hold'?'hold':'progress';}
function periodRange_(period,start,end){if(start&&end)return{start:String(start),end:String(end)};const d=new Date(),y=d.getFullYear(),m=d.getMonth();if(period==='today')return{start:today_(),end:today_()};if(period==='week'){const day=d.getDay()||7;const s=new Date(y,m,d.getDate()-day+1);return{start:fmtDate_(s),end:today_()};}if(period==='year')return{start:y+'-01-01',end:y+'-12-31'};return{start:y+'-'+pad_(m+1)+'-01',end:today_()};}
function getDb_(){if(!CONFIG.SPREADSHEET_ID||CONFIG.SPREADSHEET_ID==='PASTE_SPREADSHEET_ID_HERE')throw new Error('CONFIG.SPREADSHEET_ID belum diisi di Code.gs.');return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);}
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);return sh;}
function getSheetObjects_(name){const sh=getDb_().getSheetByName(name);if(!sh||sh.getLastRow()<2)return[];const v=sh.getDataRange().getValues(),h=v[0];return v.slice(1).filter(r=>r.some(x=>x!==''&&x!==null)).map(row=>{const o={};h.forEach((k,i)=>o[k]=row[i]);return o;});}
function clean_(v){return String(v==null?'':v).trim();}
function bool_(v){return v===true||String(v).toLowerCase()==='true'||String(v)==='1';}
function pad_(n){return('0'+n).slice(-2);}
function today_(){return Utilities.formatDate(new Date(),CONFIG.TIMEZONE,'yyyy-MM-dd');}
function fmtDate_(d){return Utilities.formatDate(d,CONFIG.TIMEZONE,'yyyy-MM-dd');}
function formatDate_(v){if(!v)return'';if(v instanceof Date)return fmtDate_(v);const s=String(v);return/^\d{4}-\d{2}-\d{2}/.test(s)?s.slice(0,10):s;}
function dateTime_(v){if(!v)return'';if(v instanceof Date)return Utilities.formatDate(v,CONFIG.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ss");return String(v);}
function durationText_(ms){if(ms<0)ms=0;const min=Math.floor(ms/60000),d=Math.floor(min/1440),h=Math.floor((min%1440)/60),m=min%60;return(d?d+'h ':'')+(h?h+'j ':'')+m+'m';}
function audit_(action,entity,id,detail){try{const u=Session.getActiveUser().getEmail()||'system';const users=getSheetObjects_(CONFIG.SHEETS.USERS);const me=users.find(x=>String(x.EMAIL||'').toLowerCase()===String(u).toLowerCase());getDb_().getSheetByName(CONFIG.SHEETS.AUDIT).appendRow([new Date(),u,me?String(me.NAME||u):u,action,entity,id,detail]);}catch(e){}}
function styleSheet_(sh){if(!sh)return;sh.setFrozenRows(1);const last=sh.getLastColumn();if(last){sh.getRange(1,1,1,last).setFontWeight('bold').setBackground('#172554').setFontColor('#ffffff');sh.autoResizeColumns(1,last);}}
