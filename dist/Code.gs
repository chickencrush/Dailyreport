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

// Kolom tambahan pada USERS & REPORTS untuk modul profil, reset password, dan bukti laporan.
const USER_EXTRA_HEADERS_ = ['RESET_CODE_HASH','RESET_EXPIRES_AT','PHOTO_URL'];
const REPORT_EXTRA_HEADER_ = 'EVIDENCE_URL';

// Sheet modul tambahan (kalender/meeting, presensi, dan izin/cuti) beserta header-nya.
const MODULE_SHEET_HEADERS_ = {
  'MEETINGS': ['ID','DATE','START_TIME','END_TIME','TITLE','LOCATION','MEETING_LINK','PARTICIPANTS','NOTES','CREATED_BY_EMAIL','CREATED_BY_NAME','CREATED_AT'],
  'ATTENDANCE': ['ID','DATE','EMAIL','USER_NAME','DIVISION','OUTLET_ID','OUTLET_NAME','SHIFT_ID','SHIFT_NAME','CHECK_IN_AT','CHECK_OUT_AT','IN_LAT','IN_LNG','IN_ACCURACY','OUT_LAT','OUT_LNG','OUT_ACCURACY','IN_PHOTO_URL','OUT_PHOTO_URL','STATUS','WORK_MINUTES','NOTES'],
  'LEAVE_REQUESTS': ['ID','EMAIL','USER_NAME','DIVISION','TYPE','START_DATE','END_DATE','REASON','STATUS','APPROVER_EMAIL','APPROVER_NAME','APPROVER_NOTE','CREATED_AT','DECIDED_AT'],
  'ATT_OUTLETS': ['ID','NAME','LATITUDE','LONGITUDE','RADIUS','ACTIVE','CREATED_AT'],
  'ATT_SHIFTS': ['ID','NAME','START_TIME','END_TIME','TOLERANCE','ACTIVE','CREATED_AT'],
  'ATT_SCHEDULES': ['ID','EMAIL','DATE','SHIFT_ID','OUTLET_ID','CREATED_AT']
};

// Daftar operasi yang boleh dipanggil lewat doPost. Fungsi yang belum ada di
// deployment otomatis tidak diregistrasi (fail closed).
const API_ROUTE_NAMES_ = ['loginUser','registerUser','getPublicDivisions','restoreSession','logoutUser',
  'changeMyPassword','requestPasswordReset','confirmPasswordReset','getBootstrapData','getDashboardData',
  'getReports','getReportDetail','saveReport','updateReportStatus','addReportNote','listUsers','saveUser',
  'listDivisions','saveDivision','getAuditLogs','exportReportsPdf','exportReportsXlsx','updateMyProfile',
  'getMyProfilePhoto','listMeetings','saveMeeting','deleteMeeting','getAttendanceOverview',
  'checkInAttendance','checkOutAttendance','getAttendancePhoto','submitLeaveRequest','listLeaveRequests',
  'decideLeaveRequest','getAttendanceSettings','saveAttendanceOutlet','saveAttendanceShift','saveEmployeeSchedule'];

// Hierarchy wewenang: SUPER_ADMIN > ADMIN > USER.
const ROLE_RANK_={'USER':1,'ADMIN':2,'SUPER_ADMIN':3};
function isSuperAdmin_(u){return u.role==='SUPER_ADMIN';}
function isAdminRole_(u){return u.role==='ADMIN'||u.role==='SUPER_ADMIN';}
function requireAdminRole_(token){var u=requireUser_(token);if(!isAdminRole_(u))throw new Error('Akses ditolak.');return u;}
function requireSuperRole_(token){var u=requireUser_(token);if(u.role!=='SUPER_ADMIN')throw new Error('Akses ditolak.');return u;}
function canEditUser_(actor,target){
  if(actor.role==='SUPER_ADMIN')return true;
  if(actor.role==='ADMIN'){
    // Admin tidak boleh mengubah akun dengan peran setara/lebih tinggi darinya.
    if(roleRank_(target)>=roleRank_(actor))return false;
    return String(target.DIVISION||'')===actor.division;
  }
  return false;
}
function roleRank_(row){return ROLE_RANK_[String(row.ROLE||'USER')]||1;}
function updateUserRole_(sh,rowIndex,role){sh.getRange(rowIndex,4+1).setValue(role);}

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
    'ACTIVITY','DEADLINE','DRIVE_URL','NOTES','COMPLETED_AT','UPDATED_AT',REPORT_EXTRA_HEADER_
  ]);
  ensureSheet_(ss, CONFIG.SHEETS.USERS, ['EMAIL','NAME','ROLE','DIVISION','ACTIVE','CREATED_AT','UPDATED_AT','SALT','PASSWORD_HASH'].concat(USER_EXTRA_HEADERS_));
  ensureSheet_(ss, CONFIG.SHEETS.SESSIONS, ['TOKEN_HASH','EMAIL','CREATED_AT','EXPIRES_AT','ACTIVE']);
  ensureSheet_(ss, CONFIG.SHEETS.DIVISIONS, ['ID','NAME','ACTIVE','CREATED_AT']);
  ensureSheet_(ss, CONFIG.SHEETS.SETTINGS, ['KEY','VALUE']);
  ensureSheet_(ss, CONFIG.SHEETS.AUDIT, ['TIMESTAMP','ACTOR_EMAIL','ACTOR_NAME','ACTION','ENTITY','ENTITY_ID','DETAIL']);
  Object.keys(MODULE_SHEET_HEADERS_).forEach(function(n){ensureSheet_(ss,n,MODULE_SHEET_HEADERS_[n]);});
  ensureUserHeaders_(ss.getSheetByName(CONFIG.SHEETS.USERS));
  seedAttendanceDefaults_(ss);

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
  updateUserPassword_(user.email,next);
  audit_('CHANGE_PASSWORD','USER',user.email,'Password changed'); return {ok:true};
}

// Reset password mandiri (tanpa SMTP): kode 6 digit ditampilkan sekali di layar reset.
function requestPasswordReset(payload){
  ensureCoreSheets_();
  payload=payload||{}; var email=clean_(payload.email).toLowerCase();
  if(!email) throw new Error('Email wajib diisi.');
  var users=getSheetObjects_(CONFIG.SHEETS.USERS), found=users.find(function(u){return String(u.EMAIL||'').toLowerCase()===email;});
  var generic={ok:false,message:'Jika akun terdaftar, kode reset 6 digit akan muncul di layar berikutnya.'};
  if(!found||!bool_(found.ACTIVE)) return generic;
  var code=''; var digits='0123456789';
  for(var i=0;i<6;i++) code+=digits.charAt(Math.floor(Math.random()*10));
  var sh=getDb_().getSheetByName(CONFIG.SHEETS.USERS),data=sh.getDataRange().getValues(),h=data[0];
  for(var r=1;r<data.length;r++){
    if(String(data[r][h.indexOf('EMAIL')]||'').toLowerCase()!==email) continue;
    sh.getRange(r+1,h.indexOf('RESET_CODE_HASH')+1).setValue(hashText_(email+'|'+code));
    sh.getRange(r+1,h.indexOf('RESET_EXPIRES_AT')+1).setValue(new Date(Date.now()+15*60000));
    break;
  }
  audit_('RESET_REQUEST','USER',email,'Kode reset dibuat');
  return {ok:true,message:'Kode reset Anda: '+code+' (berlaku 15 menit).',maskedEmail:maskEmail_(email)};
}

function confirmPasswordReset(payload){
  ensureCoreSheets_();
  payload=payload||{}; var email=clean_(payload.email).toLowerCase(), code=clean_(payload.code), next=String(payload.newPassword||'');
  if(!email||!code||!next) throw new Error('Email, kode, dan password baru wajib diisi.');
  validatePassword_(next);
  var users=getSheetObjects_(CONFIG.SHEETS.USERS), found=users.find(function(u){return String(u.EMAIL||'').toLowerCase()===email;});
  if(!found||!bool_(found.ACTIVE)) throw new Error('Kode tidak valid atau sudah kedaluwarsa.');
  var storedHash=String(found.RESET_CODE_HASH||''), exp=found.RESET_EXPIRES_AT;
  if(!storedHash||hashText_(email+'|'+code)!==storedHash) throw new Error('Kode tidak valid atau sudah kedaluwarsa.');
  if(!exp||new Date(exp).getTime()<Date.now()) throw new Error('Kode sudah kedaluwarsa. Minta kode baru.');
  updateUserPassword_(email,next);
  clearResetCode_(email); revokeSessionsForEmail_(email);
  audit_('RESET_CONFIRM','USER',email,'Password direset dengan kode');
  return {ok:true,message:'Password berhasil diganti. Silakan login dengan password baru.'};
}

function updateUserPassword_(email,password){
  var pw=makePassword_(password), sh=getDb_().getSheetByName(CONFIG.SHEETS.USERS), data=sh.getDataRange().getValues(), h=data[0];
  for(var i=1;i<data.length;i++) if(String(data[i][h.indexOf('EMAIL')]||'').toLowerCase()===email){
    sh.getRange(i+1,h.indexOf('SALT')+1).setValue(pw.salt);
    sh.getRange(i+1,h.indexOf('PASSWORD_HASH')+1).setValue(pw.hash);
    break;
  }
}
function clearResetCode_(email){
  try{var sh=getDb_().getSheetByName(CONFIG.SHEETS.USERS),data=sh.getDataRange().getValues(),h=data[0];
  for(var i=1;i<data.length;i++) if(String(data[i][h.indexOf('EMAIL')]||'').toLowerCase()===email){
    sh.getRange(i+1,h.indexOf('RESET_CODE_HASH')+1,1,2).setValues([['','']]);break;}}catch(e){}
}
function revokeSessionsForEmail_(email){
  try{var sh=getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS),data=sh.getDataRange().getValues(),h=data[0],ei=h.indexOf('EMAIL'),ai=h.indexOf('ACTIVE');
  for(var i=1;i<data.length;i++) if(String(data[i][ei]||'').toLowerCase()===email) sh.getRange(i+1,ai+1).setValue(false);}catch(e){}
}
function maskEmail_(email){var at=email.indexOf('@');if(at<2)return email;return email.slice(0,2)+'***'+email.slice(at);}

function updateMyProfile(payload,token){
  var user=requireUser_(token); payload=payload||{};
  var name=clean_(payload.name)||user.name;
  if(name.length<2) throw new Error('Nama minimal 2 karakter.');
  var photoUrl=user.photoUrl||'';
  if(payload.photo&&payload.photo.data) photoUrl=storeDataBlob_(payload.photo,'dr_profile',user.email);
  var sh=getDb_().getSheetByName(CONFIG.SHEETS.USERS),data=sh.getDataRange().getValues(),h=data[0];
  for(var i=1;i<data.length;i++) if(String(data[i][h.indexOf('EMAIL')]||'').toLowerCase()===user.email){
    sh.getRange(i+1,h.indexOf('NAME')+1).setValue(name);
    if(h.indexOf('PHOTO_URL')>=0) sh.getRange(i+1,h.indexOf('PHOTO_URL')+1).setValue(photoUrl);
    sh.getRange(i+1,h.indexOf('UPDATED_AT')+1).setValue(new Date());
    break;
  }
  var updated={email:user.email,name:name,role:user.role,division:user.division,photoUrl:photoUrl};
  cacheUser_(String(token||''),updated); setCurrentUser_(updated);
  audit_('UPDATE_PROFILE','USER',user.email,'Profil diperbarui');
  return updated;
}

function getMyProfilePhoto(token){
  var user=requireUser_(token);
  if(!user.photoUrl) return {data:''};
  return {data:loadDataBlob_(user.photoUrl)};
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
function userFromRow_(u) { return {email:String(u.EMAIL||'').toLowerCase(),name:String(u.NAME||u.EMAIL||''),role:String(u.ROLE||'USER'),division:String(u.DIVISION||''),photoUrl:String(u.PHOTO_URL||'')}; }
function ensureUserHeaders_(sh) { var expected=['EMAIL','NAME','ROLE','DIVISION','ACTIVE','CREATED_AT','UPDATED_AT','SALT','PASSWORD_HASH']; var headers=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),expected.length)).getValues()[0]; expected.forEach(function(x,i){if(headers[i]!==x)sh.getRange(1,i+1).setValue(x);}); }

function getBootstrapData(token) {
  const user = requireUser_(token);
  return {
    appName: CONFIG.APP_NAME, companyName: CONFIG.COMPANY_NAME,
    user, divisions: listVisibleDivisions_(user), today: today_(),
    dashboard: getDashboardData({period:'month'}, token),
    features: featureFlags_(user)
  };
}

function featureFlags_(user){
  const canManageUsers=user.role==='SUPER_ADMIN'||user.role==='OWNER';
  return {isAdmin:user.role!=='USER',isSuperAdmin:user.role==='SUPER_ADMIN',isOwner:user.role==='OWNER',
    canManageUsers,canConfigureAttendance:canManageUsers,attendanceEnabled:true};
}

// NOTE: single definition of restoreSession lives further below (returns {user, bootstrap}).

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
  ensureCoreSheets_();
  var sh=getDb_().getSheetByName(CONFIG.SHEETS.REPORTS),headers=shRangeHeaders_(sh);
  var row=new Array(headers.length).fill('');
  var put=function(k,v){var i=headers.indexOf(k);if(i>=0)row[i]=v;};
  put('ID',id);put('TIMESTAMP',now);put('DATE',date);put('USER_EMAIL',user.email);put('USER_NAME',user.name);
  put('DIVISION',user.division);put('ROLE',user.role);put('STATUS',status);put('ACTIVITY',activity);
  put('DEADLINE',deadline);put('DRIVE_URL',driveUrl);put('NOTES',clean_(payload.notes));
  put('COMPLETED_AT',completedAt);put('UPDATED_AT',now);put(REPORT_EXTRA_HEADER_,clean_(payload.evidenceUrl));
  sh.appendRow(row);
  audit_('CREATE_REPORT','REPORT',id,activity + ' | ' + status);
  return {ok:true,id,message:'Laporan berhasil disimpan.'};
}

function shRangeHeaders_(sh){return sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String);}

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

function listUsers(token){var actor=requireAdminRole_(token);return getSheetObjects_(CONFIG.SHEETS.USERS).map(u=>({email:String(u.EMAIL||'').toLowerCase(),name:String(u.NAME||''),role:String(u.ROLE||'USER'),division:String(u.DIVISION||''),active:bool_(u.ACTIVE),createdAt:dateTime_(u.CREATED_AT),hasPassword:!!String(u.PASSWORD_HASH||''),canEdit:canEditUser_(actor,u)}));}
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
  ensureSheet_(ss,CONFIG.SHEETS.USERS,['EMAIL','NAME','ROLE','DIVISION','ACTIVE','CREATED_AT','UPDATED_AT','SALT','PASSWORD_HASH'].concat(USER_EXTRA_HEADERS_));
  ensureSheet_(ss,CONFIG.SHEETS.DIVISIONS,['ID','NAME','ACTIVE','CREATED_AT']);
  ensureSheet_(ss,CONFIG.SHEETS.SESSIONS,['TOKEN_HASH','EMAIL','CREATED_AT','EXPIRES_AT','ACTIVE']);
  ensureSheet_(ss,CONFIG.SHEETS.REPORTS,['ID','TIMESTAMP','DATE','USER_EMAIL','USER_NAME','DIVISION','ROLE','STATUS','ACTIVITY','DEADLINE','DRIVE_URL','NOTES','COMPLETED_AT','UPDATED_AT',REPORT_EXTRA_HEADER_]);
  ensureSheet_(ss,CONFIG.SHEETS.SETTINGS,['KEY','VALUE']);
  ensureSheet_(ss,CONFIG.SHEETS.AUDIT,['TIMESTAMP','ACTOR_EMAIL','ACTOR_NAME','ACTION','ENTITY','ENTITY_ID','DETAIL']);
  Object.keys(MODULE_SHEET_HEADERS_).forEach(function(name){ensureSheet_(ss,name,MODULE_SHEET_HEADERS_[name]);});
  // Seed outlet/shift presensi default agar check-in langsung bisa dipakai tanpa setupApp manual.
  try{seedAttendanceDefaults_();}catch(e){}
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

/* ================= MEETING (KALENDER) ================= */
function listMeetings(payload,token){
  requireUser_(token); payload=payload||{};
  var month=clean_(payload.month), rows=getSheetObjects_('MEETINGS');
  if(/^\d{4}-\d{2}$/.test(month)) rows=rows.filter(function(x){return formatDate_(x.DATE).slice(0,7)===month;});
  return rows.map(function(x){return {
    id:String(x.ID||''), date:formatDate_(x.DATE), startTime:String(x.START_TIME||''), endTime:String(x.END_TIME||''),
    title:String(x.TITLE||''), location:String(x.LOCATION||''), meetingLink:String(x.MEETING_LINK||''),
    participants:String(x.PARTICIPANTS||''), notes:String(x.NOTES||''), createdByName:String(x.CREATED_BY_NAME||'')
  };}).sort(function(a,b){return (a.date+a.startTime).localeCompare(b.date+b.startTime);});
}

function saveMeeting(payload,token){
  var user=requireAdminRole_(token); payload=payload||{};
  var title=clean_(payload.title), date=clean_(payload.date), startTime=clean_(payload.startTime), endTime=clean_(payload.endTime);
  if(!title) throw new Error('Judul meeting wajib diisi.');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal meeting tidak valid.');
  if(!/^\d{2}:\d{2}$/.test(startTime)||!/^\d{2}:\d{2}$/.test(endTime)) throw new Error('Jam meeting tidak valid.');
  if(endTime<=startTime) throw new Error('Jam selesai harus setelah jam mulai.');
  ensureCoreSheets_();
  var sh=getDb_().getSheetByName('MEETINGS'), h=MODULE_SHEET_HEADERS_.MEETINGS;
  sh.appendRow([Utilities.getUuid(),date,startTime,endTime,title,clean_(payload.location)||'Online',
    clean_(payload.meetingLink),clean_(payload.participants),clean_(payload.notes),user.email,user.name,new Date()]);
  audit_('CREATE_MEETING','MEETING',String(sh.getRange(sh.getLastRow(),1).getValue()),title+' | '+date);
  return {ok:true,message:'Jadwal meeting tersimpan.'};
}

function deleteMeeting(id,token){
  var user=requireAdminRole_(token); id=clean_(id); if(!id) throw new Error('ID meeting tidak valid.');
  var sh=getDb_().getSheetByName('MEETINGS'), data=sh.getDataRange().getValues(), h=data[0], ii=h.indexOf('ID');
  for(var i=1;i<data.length;i++) if(String(data[i][ii])===id){
    sh.deleteRow(i+1); audit_('DELETE_MEETING','MEETING',id,'Dihapus oleh '+user.email); return {ok:true};
  }
  throw new Error('Meeting tidak ditemukan.');
}

/* ================= PRESANSI & IZIN/CUTI ================= */
function nowInTz_(tz){
  // Waktu saat ini dalam zona aplikasi sebagai objek Date absolut (memakai offset tz).
  var s=Utilities.formatDate(new Date(),tz||CONFIG.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX");
  var d=new Date(s);
  return isNaN(d.getTime())?new Date():d;
}
function timeText_(v,tz){if(!v)return'';if(v instanceof Date)return Utilities.formatDate(v,tz||CONFIG.TIMEZONE,'HH:mm');return String(v).slice(11,16);}
function clockMinutes_(hhmm){var p=String(hhmm||'').split(':');return (Number(p[0])||0)*60+(Number(p[1])||0);}
function haversineMeters_(lat1,lng1,lat2,lng2){
  var R=6371000,toRad=Math.PI/180;
  var dLat=(lat2-lat1)*toRad,dLng=(lng2-lng1)*toRad;
  var a=Math.sin(dLat/2)*Math.sin(dLat/2)+Math.cos(lat1*toRad)*Math.cos(lat2*toRad)*Math.sin(dLng/2)*Math.sin(dLng/2);
  return 2*R*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}
function findAttendanceRowById_(id){
  var sh=getDb_().getSheetByName('ATTENDANCE'),data=sh.getDataRange().getValues(),h=data[0],ii=h.indexOf('ID');
  for(var i=1;i<data.length;i++) if(String(data[i][ii])===String(id)) return {sh:sh,data:data,h:h,row:i+1,index:i};
  return null;
}
function attRowView_(r){
  return {id:String(r.ID||''),date:formatDate_(r.DATE),userEmail:String(r.EMAIL||'').toLowerCase(),userName:String(r.USER_NAME||''),
    division:String(r.DIVISION||''),outletId:String(r.OUTLET_ID||''),outletName:String(r.OUTLET_NAME||''),
    shiftId:String(r.SHIFT_ID||''),shiftName:String(r.SHIFT_NAME||''),
    checkIn:dateTime_(r.CHECK_IN_AT),checkOut:dateTime_(r.CHECK_OUT_AT),status:String(r.STATUS||''),
    workMinutes:Number(r.WORK_MINUTES||0)||0,notes:String(r.NOTES||''),
    hasCheckInPhoto:!!String(r.IN_PHOTO_URL||''),hasCheckOutPhoto:!!String(r.OUT_PHOTO_URL||''),
    checkInPhotoUrl:bool_(r.IN_PHOTO_URL)?'available':'',checkOutPhotoUrl:bool_(r.OUT_PHOTO_URL)?'available':''};
}
function attendanceRowsFor_(user,filters){
  filters=filters||{};
  var range=periodRange_(filters.period||'month',filters.startDate,filters.endDate);
  var rows=getSheetObjects_('ATTENDANCE').map(function(r){return {raw:r,v:attRowView_(r)};});
  if(user.role==='ADMIN') rows=rows.filter(function(x){return x.v.division===user.division;});
  else if(user.role!=='SUPER_ADMIN') rows=rows.filter(function(x){return x.v.userEmail===user.email;});
  rows=rows.filter(function(x){return x.v.date>=range.start&&x.v.date<=range.end;});
  if(filters.division) rows=rows.filter(function(x){return x.v.division===String(filters.division);});
  if(filters.outletId) rows=rows.filter(function(x){return x.v.outletId===String(filters.outletId);});
  if(filters.email) rows=rows.filter(function(x){return x.v.userEmail===String(filters.email).toLowerCase();});
  return rows;
}
function getAttendanceOverview(payload,token){
  var user=requireUser_(token); payload=payload||{};
  var today=today_();
  var scoped=attendanceRowsFor_(user,Object.assign({},payload,{period:payload.period||'month'}));
  var mine=scoped.filter(function(x){return x.v.userEmail===user.email;});
  var current=mine.find(function(x){return x.v.date===today;});
  var schedule=findScheduleForToday_(user.email,today);
  return {
    period:periodRange_(payload.period||'month',payload.startDate,payload.endDate),
    today:today,
    rows:scoped.map(function(x){return x.v;}).sort(function(a,b){return (b.date+b.checkIn).localeCompare(a.date+a.checkIn);}),
    outlets:getActiveOutlets_(),
    shifts:getActiveShifts_(),
    schedule:schedule,
    current:current?current.v:null,
    summary:{total:mine.length,onTime:mine.filter(function(x){return x.v.status==='Tepat Waktu';}).length,
      late:mine.filter(function(x){return x.v.status==='Terlambat';}).length,
      open:mine.filter(function(x){return !x.v.checkOut;}).length}
  };
}
function checkInAttendance(payload,token){
  var user=requireUser_(token); payload=payload||{}; ensureCoreSheets_();
  var outlet=findOutletById_(payload.outletId), shift=findShiftById_(payload.shiftId);
  if(!outlet) throw new Error('Outlet tidak ditemukan atau tidak aktif.');
  if(!shift) throw new Error('Shift tidak ditemukan atau tidak aktif.');
  var lat=Number(payload.latitude),lng=Number(payload.longitude);
  if(!isFinite(lat)||!isFinite(lng)) throw new Error('Lokasi GPS tidak valid. Ambil lokasi terlebih dahulu.');
  var distance=haversineMeters_(lat,lng,Number(outlet.lat),Number(outlet.lng));
  var radius=Number(outlet.radius)||100;
  if(distance>radius) throw new Error('Di luar jangkauan outlet ('+Math.round(distance)+' m dari target, batas '+Math.round(radius)+' m).');
  var tz=CONFIG.TIMEZONE, today=today_(), now=nowInTz_(tz);
  var existing=findMyAttendanceToday_(user.email,today);
  if(existing&&existing.CHECK_IN_AT) throw new Error('Anda sudah check-in hari ini. Silakan check-out.');
  var photoRef=payload.photo&&payload.photo.data?storeDataBlob_(payload.photo,'dr_attendance',user.email):'';
  var status=distanceFromSchedule_(shift,timeText_(now,tz))<=clockMinutes_(shift.tolerance||0)?'Tepat Waktu':'Terlambat';
  var id='ATT-'+Utilities.getUuid().slice(0,8).toUpperCase();
  var row={ID:id,DATE:today,EMAIL:user.email,USER_NAME:user.name,DIVISION:user.division,
    OUTLET_ID:outlet.id,OUTLET_NAME:outlet.name,SHIFT_ID:shift.id,SHIFT_NAME:shift.name,
    CHECK_IN_AT:now,IN_LAT:lat,IN_LNG:lng,IN_ACCURACY:Number(payload.accuracy)||'',IN_PHOTO_URL:photoRef,
    STATUS:status,NOTES:clean_(payload.notes)};
  appendSheetObject_('ATTENDANCE',row);
  audit_('CHECK_IN','ATTENDANCE',id,outlet.name+' | '+shift.name+' | '+status);
  return {ok:true,id:id,status:status,message:'Check-in pukul '+timeText_(now,tz)+' · '+outlet.name};
}
function checkOutAttendance(payload,token){
  var user=requireUser_(token); payload=payload||{}; ensureCoreSheets_();
  var found=findMyAttendanceToday_(user.email,today_());
  if(!found) throw new Error('Anda belum check-in hari ini.');
  if(found.CHECK_OUT_AT) throw new Error('Anda sudah check-out hari ini.');
  var tz=CONFIG.TIMEZONE, now=nowInTz_(tz);
  var outlet=findOutletById_(payload.outletId||found.OUTLET_ID), shift=findShiftById_(payload.shiftId||found.SHIFT_ID);
  if(outlet){
    var lat=Number(payload.latitude),lng=Number(payload.longitude);
    if(isFinite(lat)&&isFinite(lng)){
      var distance=haversineMeters_(lat,lng,Number(outlet.lat),Number(outlet.lng)),radius=Number(outlet.radius)||100;
      if(distance>radius) throw new Error('Di luar jangkauan outlet saat check-out ('+Math.round(distance)+' m, batas '+Math.round(radius)+' m).');
    }
  }
  var photoRef=payload.photo&&payload.photo.data?storeDataBlob_(payload.photo,'dr_attendance',user.email):'';
  var inAt=new Date(found.CHECK_IN_AT), workMinutes=Math.max(0,Math.round((now-inAt)/60000));
  var row=findAttendanceRowById_(found.ID);
  if(!row) throw new Error('Data presensi tidak ditemukan. Muat ulang halaman.');
  var h=row.h, set=function(k,v){var i=h.indexOf(k);if(i>=0)row.sh.getRange(row.row,i+1).setValue(v);};
  set('CHECK_OUT_AT',now);
  if(photoRef)set('OUT_PHOTO_URL',photoRef);
  set('WORK_MINUTES',workMinutes);
  set('OUT_LAT',Number(payload.latitude)||'');set('OUT_LNG',Number(payload.longitude)||'');set('OUT_ACCURACY',Number(payload.accuracy)||'');
  if(clean_(payload.notes))set('NOTES',String(found.NOTES||'')+'\n'+clean_(payload.notes));
  audit_('CHECK_OUT','ATTENDANCE',String(found.ID),'Durasi '+workMinutes+' menit');
  return {ok:true,id:String(found.ID),message:'Check-out pukul '+timeText_(now,tz)+' · durasi '+attendanceDurationText_(workMinutes)};
}
function getAttendancePhoto(id,kind,token){
  var user=requireUser_(token);
  var row=findAttendanceRowById_(clean_(id)); if(!row) throw new Error('Data presensi tidak ditemukan.');
  var record={}; row.h.forEach(function(k,i){record[k]=row.data[row.index][i];});
  authorizeReportAccess_(user,String(record.EMAIL||'').toLowerCase(),String(record.DIVISION||''));
  var ref=String(kind==='out'?record.OUT_PHOTO_URL:record.IN_PHOTO_URL||'');
  if(!ref) throw new Error('Foto tidak tersedia.');
  return {data:loadDataBlob_(ref)};
}
function submitLeaveRequest(payload,token){
  var user=requireUser_(token); payload=payload||{}; ensureCoreSheets_();
  var type=clean_(payload.type), startDate=clean_(payload.startDate), endDate=clean_(payload.endDate)||startDate, reason=clean_(payload.reason);
  if(!type) throw new Error('Jenis pengajuan wajib dipilih.');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(startDate)) throw new Error('Tanggal mulai tidak valid.');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(endDate)) throw new Error('Tanggal selesai tidak valid.');
  if(endDate<startDate) throw new Error('Tanggal selesai tidak boleh sebelum tanggal mulai.');
  if(reason.length<5) throw new Error('Alasan minimal 5 karakter.');
  var id='LVQ-'+Utilities.getUuid().slice(0,8).toUpperCase();
  appendSheetObject_('LEAVE_REQUESTS',{ID:id,EMAIL:user.email,USER_NAME:user.name,DIVISION:user.division,TYPE:type,
    START_DATE:startDate,END_DATE:endDate,REASON:reason,STATUS:'Diajukan',CREATED_AT:new Date()});
  audit_('SUBMIT_LEAVE','LEAVE',id,type+' | '+startDate+' s/d '+endDate);
  return {ok:true,id:id,message:'Pengajuan dikirim dan menunggu persetujuan.'};
}
function listLeaveRequests(payload,token){
  var user=requireUser_(token); payload=payload||{};
  var rows=getSheetObjects_('LEAVE_REQUESTS');
  if(user.role==='ADMIN') rows=rows.filter(function(x){return String(x.DIVISION||'')===user.division;});
  else if(user.role!=='SUPER_ADMIN') rows=rows.filter(function(x){return String(x.EMAIL||'').toLowerCase()===user.email;});
  if(payload.status) rows=rows.filter(function(x){return String(x.STATUS||'')===String(payload.status);});
  return rows.map(function(x){return {id:String(x.ID||''),userEmail:String(x.EMAIL||'').toLowerCase(),userName:String(x.USER_NAME||''),
    division:String(x.DIVISION||''),type:String(x.TYPE||''),startDate:formatDate_(x.START_DATE),endDate:formatDate_(x.END_DATE),
    reason:String(x.REASON||''),status:String(x.STATUS||'Diajukan'),approverNote:String(x.APPROVER_NOTE||''),
    createdAt:dateTime_(x.CREATED_AT)};}).sort(function(a,b){return b.createdAt.localeCompare(a.createdAt);});
}
function decideLeaveRequest(id,payload,token){
  var user=requireAdminRole_(token); payload=payload||{};
  var status=clean_(payload.status);
  if(['Disetujui','Ditolak'].indexOf(status)<0) throw new Error('Keputusan tidak valid.');
  var sh=getDb_().getSheetByName('LEAVE_REQUESTS'),data=sh.getDataRange().getValues(),h=data[0],ii=h.indexOf('ID');
  for(var i=1;i<data.length;i++) if(String(data[i][ii])===clean_(id)){
    if(user.role==='ADMIN'&&String(data[i][h.indexOf('DIVISION')]||'')!==user.division) throw new Error('Akses ditolak.');
    var set=function(k,v){var idx=h.indexOf(k);if(idx>=0)sh.getRange(i+1,idx+1).setValue(v);};
    set('STATUS',status);set('APPROVER_EMAIL',user.email);set('APPROVER_NAME',user.name);
    set('APPROVER_NOTE',clean_(payload.note));set('DECIDED_AT',new Date());
    audit_('DECIDE_LEAVE','LEAVE',String(data[i][ii]),status);
    return {ok:true,message:'Pengajuan '+status.toLowerCase()+'.'};
  }
  throw new Error('Pengajuan tidak ditemukan.');
}
function getAttendanceSettings(token){
  requireAdminRole_(token); ensureCoreSheets_(); seedAttendanceDefaults_();
  var users=getSheetObjects_(CONFIG.SHEETS.USERS).filter(function(u){return bool_(u.ACTIVE);})
    .map(function(u){return {email:String(u.EMAIL||'').toLowerCase(),name:String(u.NAME||u.EMAIL||''),division:String(u.DIVISION||'')};});
  return {outlets:getAllOutlets_(true),shifts:getAllShifts_(true),users:users,today:today_()};
}
function saveAttendanceOutlet(payload,token){
  requireAdminRole_(token); payload=payload||{}; ensureCoreSheets_();
  var name=clean_(payload.name); if(!name) throw new Error('Nama outlet wajib diisi.');
  var lat=Number(payload.latitude!=null&&payload.latitude!==''?payload.latitude:payload.lat);
  var lng=Number(payload.longitude!=null&&payload.longitude!==''?payload.longitude:payload.lng);
  if(!isFinite(lat)||!isFinite(lng)) throw new Error('Koordinat latitude/longitude tidak valid.');
  var radius=Number(payload.radius)||100;
  var id=clean_(payload.id)||'OUL-'+Utilities.getUuid().slice(0,6).toUpperCase();
  upsertSheetRow_('ATT_OUTLETS','ID',id,['ID','NAME','LATITUDE','LONGITUDE','RADIUS','ACTIVE','CREATED_AT'],
    [id,name,lat,lng,radius,payload.active!==false,new Date()]);
  audit_('SAVE_OUTLET','ATTENDANCE',id,name);
  return {ok:true,id:id};
}
function saveAttendanceShift(payload,token){
  requireAdminRole_(token); payload=payload||{}; ensureCoreSheets_();
  var name=clean_(payload.name); if(!name) throw new Error('Nama shift wajib diisi.');
  var start=clean_(payload.start),end=clean_(payload.end);
  if(!/^\d{2}:\d{2}$/.test(start)||!/^\d{2}:\d{2}$/.test(end)) throw new Error('Jam shift tidak valid.');
  var tolerance=Number(payload.tolerance)||0,id=clean_(payload.id)||'SHF-'+Utilities.getUuid().slice(0,6).toUpperCase();
  upsertSheetRow_('ATT_SHIFTS','ID',id,['ID','NAME','START_TIME','END_TIME','TOLERANCE','ACTIVE','CREATED_AT'],
    [id,name,start,end,tolerance,payload.active!==false,new Date()]);
  audit_('SAVE_SHIFT','ATTENDANCE',id,name);
  return {ok:true,id:id};
}
function saveEmployeeSchedule(payload,token){
  requireAdminRole_(token); payload=payload||{}; ensureCoreSheets_();
  var email=clean_(payload.email).toLowerCase(),date=clean_(payload.date),shiftId=clean_(payload.shiftId),outletId=clean_(payload.outletId);
  if(!email) throw new Error('Karyawan wajib dipilih.');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal jadwal tidak valid.');
  if(!findShiftById_(shiftId)) throw new Error('Shift tidak ditemukan.');
  if(!findOutletById_(outletId)) throw new Error('Outlet tidak ditemukan.');
  upsertSheetRow_('ATT_SCHEDULES','EMAIL_DATE',email+'|'+date,['ID','EMAIL','DATE','SHIFT_ID','OUTLET_ID','CREATED_AT'],
    [Utilities.getUuid(),email,date,shiftId,outletId,new Date()],'EMAIL_DATE');
  audit_('SAVE_SCHEDULE','ATTENDANCE',email+'|'+date,shiftId+' @ '+outletId);
  return {ok:true,message:'Jadwal karyawan tersimpan.'};
}
function findMyAttendanceToday_(email,date){
  var rows=getSheetObjects_('ATTENDANCE').filter(function(x){return String(x.EMAIL||'').toLowerCase()===email&&formatDate_(x.DATE)===date;});
  return rows.length?rows[0]:null;
}
function findScheduleForToday_(email,date){
  var rows=getSheetObjects_('ATT_SCHEDULES').filter(function(x){return String(x.EMAIL||'').toLowerCase()===email&&formatDate_(x.DATE)===date;});
  if(!rows.length) return null;
  var s=rows[0],shift=findShiftById_(s.SHIFT_ID),outlet=findOutletById_(s.OUTLET_ID);
  if(!shift||!outlet) return null;
  return {date:date,shiftId:String(shift.id),outletId:String(outlet.id),shiftName:shift.name,outletName:outlet.name};
}
function getAllOutlets_(includeInactive){
  return getSheetObjects_('ATT_OUTLETS').map(function(x){return {id:String(x.ID),name:String(x.NAME),lat:Number(x.LATITUDE),lng:Number(x.LONGITUDE),
    radius:Number(x.RADIUS)||100,active:bool_(x.ACTIVE)};}).filter(function(x){return includeInactive||x.active;});
}
function getAllShifts_(includeInactive){
  return getSheetObjects_('ATT_SHIFTS').map(function(x){return {id:String(x.ID),name:String(x.NAME),start:String(x.START_TIME),end:String(x.END_TIME),
    tolerance:Number(x.TOLERANCE)||0,active:bool_(x.ACTIVE)};}).filter(function(x){return includeInactive||x.active;});
}
function getActiveOutlets_(){return getAllOutlets_(false);}
function getActiveShifts_(){return getAllShifts_(false);}
function findOutletById_(id){id=clean_(id);return getAllOutlets_(false).find(function(x){return x.id===id;})||getAllOutlets_(true).find(function(x){return x.id===id;})||null;}
function findShiftById_(id){id=clean_(id);return getAllShifts_(false).find(function(x){return x.id===id;})||getAllShifts_(true).find(function(x){return x.id===id;})||null;}
function distanceFromSchedule_(shift,hhmm){
  var m=clockMinutes_(hhmm),start=clockMinutes_(shift.start),end=clockMinutes_(shift.end);
  if(end<=start)end+=1440; // shift melewati tengah malam
  var candidates=[m,m+1440];
  var best=Infinity;
  candidates.forEach(function(c){best=Math.min(best,Math.abs(c-start),Math.abs(c-end));});
  return best;
}
function attendanceDurationText_(minutes){minutes=Number(minutes||0);return Math.floor(minutes/60)+'j '+(minutes%60)+'m';}
function appendSheetObject_(name,obj){
  var sh=getDb_().getSheetByName(name),headers=shRangeHeaders_(sh);
  var row=new Array(headers.length).fill('');
  Object.keys(obj).forEach(function(k){var i=headers.indexOf(k);if(i>=0)row[i]=obj[k];});
  sh.appendRow(row);
  return row;
}
function upsertSheetRow_(sheetName,keyCol,keyValue,headers,values,keyComposite){
  var sh=getDb_().getSheetByName(sheetName),data=sh.getDataRange().getValues(),h=data[0].map(String);
  var matchIndex=-1;
  if(keyComposite==='EMAIL_DATE'){
    var ei=h.indexOf('EMAIL'),di=h.indexOf('DATE');
    var parts=String(keyValue).split('|');
    for(var i=1;i<data.length;i++){if(String(data[i][ei]||'').toLowerCase()===parts[0]&&formatDate_(data[i][di])===parts[1]){matchIndex=i+1;break;}}
  }else{
    var ki=h.indexOf(keyCol);
    for(var j=1;j<data.length;j++){if(String(data[j][ki])===String(keyValue)){matchIndex=j+1;break;}}
  }
  if(matchIndex>0){
    var out=data[matchIndex-1].slice();
    headers.forEach(function(col,n){var idx=h.indexOf(col);if(idx>=0)out[idx]=values[n];});
    sh.getRange(matchIndex,1,1,out.length).setValues([out]);
  }else{
    var fresh=new Array(h.length).fill('');
    headers.forEach(function(col,n){var idx=h.indexOf(col);if(idx>=0)fresh[idx]=values[n];});
    sh.appendRow(fresh);
  }
}
function seedAttendanceDefaults_(){
  // Bisa dipanggil tanpa argumen (dari ensureCoreSheets_/API) maupun dengan ss (dari setupApp).
  var ss=getDb_();
  ['ATT_OUTLETS','ATT_SHIFTS','ATT_SCHEDULES','ATTENDANCE','LEAVE_REQUESTS','MEETINGS'].forEach(function(n){ensureSheet_(ss,n,MODULE_SHEET_HEADERS_[n]);});
  var outlets=ss.getSheetByName('ATT_OUTLETS');
  if(outlets&&outlets.getLastRow()<2){
    outlets.getRange(2,1,1,7).setValues([['OUL-KTR01','Kantor Pusat Chicken Crush',-6.200000,106.816666,150,true,new Date()]]);
  }
  var shifts=ss.getSheetByName('ATT_SHIFTS');
  if(shifts&&shifts.getLastRow()<2){
    shifts.getRange(2,1,3,7).setValues([
      ['SHF-PG01','Pagi','08:00','17:00',10,true,new Date()],
      ['SHF-SR02','Siang','12:00','21:00',10,true,new Date()]
    ]);
  }
}

/* Simpan muat blob base64 (foto profil/presensi) ke Drive lewat URL referensi "drive:<id>". */
function storeDataBlob_(dataUrl,prefix,email){
  var parts=String(dataUrl.data||'').split(',');
  var mime=(String(dataUrl.mimeType||'image/jpeg')).split(';')[0];
  if(!/^data:image\//i.test(mime)) mime='image/jpeg';
  var bytes=Utilities.base64Decode(parts.length>1?parts[parts.length-1]:String(dataUrl.data||''));
  var name=prefix+'-'+String(email||'anon').replace(/[^a-z0-9]/gi,'_')+'-'+Utilities.getUuid().slice(0,8)+'.jpg';
  var folder=CONFIG.DRIVE_FOLDER_ID?DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID):DriveApp.getRootFolder();
  var file=folder.createFile(Utilities.newBlob(bytes,mime,name));
  file.setDescription(prefix+' upload - '+email);
  return 'drive:'+file.getId();
}
function loadDataBlob_(ref){
  var id=String(ref||'');
  if(id.indexOf('drive:')===0)id=id.slice(6);
  try{
    var blob=DriveApp.getFileById(id).getBlob();
    return 'data:'+(blob.getContentType()||'image/jpeg')+';base64,'+Utilities.base64Encode(blob.getBytes());
  }catch(e){throw new Error('Berkas foto tidak dapat dimuat.');}
}

// Public JSON endpoint. Only operations listed in API_ROUTE_NAMES_ are reachable.
function getApiRoutes_(){
  // Peta ulang setiap permintaan agar fungsi baru langsung aktif tanpa restart runtime.
  // CATATAN: jangan pakai this[] (undefined pada Apps Script V8) dan jangan pakai eval
  // (dinonaktifkan sebagian pada runtime tertentu). Global object lookup yang andal:
  // this -> globalThis -> self -> window -> Function('return this')().
  var scope=(typeof globalThis!=='undefined'&&globalThis)?globalThis
    :(typeof self!=='undefined'&&self)?self
    :(typeof window!=='undefined'&&window)?window
    :null;
  if(!scope){ try{ scope=Function('return this')(); }catch(e){ scope=null; } }
  var routes={};
  for(var i=0;i<API_ROUTE_NAMES_.length;i++){
    var name=API_ROUTE_NAMES_[i];
    try{
      var fn=scope?scope[name]:null;
      if(typeof fn==='function') routes[name]=fn;
    }catch(e){}
  }
  return routes;
}
function doPost(e){
  try{
    if(!e||!e.postData||!e.postData.contents)throw new Error('Permintaan kosong.');
    var req=JSON.parse(e.postData.contents);
    if(!req||typeof req.action!=='string'||!Array.isArray(req.args))throw new Error('Permintaan tidak valid.');
    var routes=getApiRoutes_();
    if(!Object.prototype.hasOwnProperty.call(routes,req.action)){
      throw new Error('Operasi "'+req.action+'" belum tersedia di deployment. Simpan Code.gs lalu Deploy > Manage deployments > Edit > New version.');
    }
    var fn=routes[req.action],args=req.args.slice();
    while(args.length<fn.length)args.push(null);
    return apiJson_({ok:true,data:fn.apply(null,args)});
  }catch(err){return apiJson_({ok:false,error:String(err.message||err)});}
}
function apiJson_(value){return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);}
