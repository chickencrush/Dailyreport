let installPrompt=null;
window.addEventListener('beforeinstallprompt',event=>{event.preventDefault();installPrompt=event;});
document.getElementById('install').addEventListener('click',async()=>{if(installPrompt){const prompt=installPrompt;installPrompt=null;try{await prompt.prompt();await prompt.userChoice;}catch(e){document.getElementById('guide').showModal();}}else document.getElementById('guide').showModal();});
window.addEventListener('appinstalled',()=>{document.getElementById('install').classList.add('hidden');});
function validAppUrl(value){try{const url=new URL(value);return url.protocol==='https:'&&url.hostname==='script.google.com'&&/^\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url.pathname);}catch(e){return false;}}
const appUrl=window.DAILY_REPORT_URL||'';
if(validAppUrl(appUrl)){document.getElementById('setup').classList.add('hidden');const app=document.getElementById('app');app.src=appUrl;app.classList.remove('hidden');const link=document.getElementById('external');link.href=appUrl;link.classList.remove('hidden');}
else if(appUrl){document.getElementById('setupText').textContent='Alamat aplikasi belum valid. Gunakan URL deployment Google Apps Script berakhiran /exec.';}
function connectionStatus(){document.getElementById('offline').classList.toggle('hidden',navigator.onLine);}
window.addEventListener('online',connectionStatus);window.addEventListener('offline',connectionStatus);connectionStatus();
if('serviceWorker' in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});
