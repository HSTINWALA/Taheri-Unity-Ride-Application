const $=(s,r=document)=>r.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=t=>t?new Date(t).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit',second:'2-digit'}):'–';
async function api(action,d={},tokKey){
  try{const r=await fetch(window.API_URL,{method:'POST',body:JSON.stringify({action,token:localStorage.getItem(tokKey||window.TOK),...d})});return await r.json();}
  catch(e){return {ok:false,error:'Network error – check connection and API_URL in config.js'}}
}
function toast(msg,cls='ok',ms=2500){const t=$('#toast');t.className=cls;t.textContent=msg;t.style.display='block';clearTimeout(t._h);t._h=setTimeout(()=>t.style.display='none',ms)}
function beep(good){try{navigator.vibrate&&navigator.vibrate(good?80:[120,60,120]);const c=new AudioContext(),o=c.createOscillator();o.frequency.value=good?880:220;o.connect(c.destination);o.start();o.stop(c.currentTime+.15)}catch(e){}}
let _qr,_last='',_lt=0;
async function startScan(onCode){
  if(!_qr)_qr=new Html5Qrcode('reader');
  if(_qr.isScanning)return;
  try{await _qr.start({facingMode:'environment'},{fps:10,qrbox:{width:230,height:230}},t=>{const n=Date.now();if(t===_last&&n-_lt<4000)return;_last=t;_lt=n;onCode(t.trim())})}
  catch(e){toast('Camera blocked. Allow camera access (page must be on HTTPS).','err',5000)}
}
async function stopScan(){if(_qr&&_qr.isScanning)await _qr.stop()}
