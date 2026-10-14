/* FIWA Presence - kode aplikasi (dipindahkan dari template Blogger). */
(function(){
'use strict';
/* =====================================================================
   1. UTILITAS
   ===================================================================== */
const $=(s,r=document)=>r.querySelector(s), $$=(s,r=document)=>Array.from(r.querySelectorAll(s));
const esc=v=>String(v==null?'':v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const pad=n=>String(n).padStart(2,'0');
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const MON=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const MONS=['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
const DAYS=['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'];
const DAYS3=['Sen','Sel','Rab','Kam','Jum','Sab','Min'];
const TZMAP={'Asia/Jakarta':7,'Asia/Makassar':8,'Asia/Jayapura':9};
let TZH=7;
const D={
  parse(ds){const [y,m,d]=String(ds).slice(0,10).split('-').map(Number);return{y,m,d}},
  add(ds,n){const{y,m,d}=D.parse(ds);return new Date(Date.UTC(y,m-1,d+n)).toISOString().slice(0,10)},
  iso(ds){const{y,m,d}=D.parse(ds);const w=new Date(Date.UTC(y,m-1,d)).getUTCDay();return w===0?7:w},
  diff(a,b){const A=D.parse(a),B=D.parse(b);return Math.round((Date.UTC(B.y,B.m-1,B.d)-Date.UTC(A.y,A.m-1,A.d))/864e5)},
  epoch(ds,hm){const{y,m,d}=D.parse(ds);const t=String(hm||'00:00').split(':').map(Number);return Date.UTC(y,m-1,d,t[0]||0,t[1]||0,t[2]||0)-TZH*36e5},
  local(ms){const t=new Date(ms+TZH*36e5);return{date:t.toISOString().slice(0,10),hm:pad(t.getUTCHours())+':'+pad(t.getUTCMinutes()),hms:pad(t.getUTCHours())+':'+pad(t.getUTCMinutes())+':'+pad(t.getUTCSeconds())}},
  mStart(ds){const{y,m}=D.parse(ds);return y+'-'+pad(m)+'-01'},
  mEnd(ds){const{y,m}=D.parse(ds);return new Date(Date.UTC(y,m,0)).toISOString().slice(0,10)},
  range(a,b){const o=[];for(let d=a;d<=b;d=D.add(d,1))o.push(d);return o},
  hm(t){return t?String(t).slice(0,5):''}
};
const T=(x)=>x?Date.parse(x):null;
const fmtD=ds=>{if(!ds)return'-';const{y,m,d}=D.parse(ds);return pad(d)+'/'+pad(m)+'/'+y};
const fmtDL=ds=>{const{y,m,d}=D.parse(ds);return DAYS[D.iso(ds)-1]+', '+pad(d)+'/'+pad(m)+'/'+y};
const fmtM=ds=>{const{y,m}=D.parse(ds);return MON[m-1]+' '+y};
const fmtT=ms=>ms==null?'-':D.local(ms).hm;
const fmtDT=x=>{if(!x)return'-';const ms=typeof x==='number'?x:Date.parse(x);const l=D.local(ms);return fmtD(l.date)+' '+l.hm};
const nf=n=>new Intl.NumberFormat('id-ID',{maximumFractionDigits:2}).format(n==null?0:n);
const fmtMin=m=>{m=Math.round(m||0);if(!m)return'0m';const h=Math.floor(m/60),r=m%60;return(h?h+'j ':'')+(r||!h?r+'m':'').trim()};
const dayName=ds=>DAYS3[D.iso(ds)-1];
const uid=()=>Math.random().toString(36).slice(2,10);
const debounce=(fn,ms=250)=>{let t;return(...a)=>{clearTimeout(t);t=setTimeout(()=>fn(...a),ms)}};
const clone=o=>JSON.parse(JSON.stringify(o));
const sum=(a,f)=>a.reduce((s,x)=>s+(f?f(x):x),0);
const groupBy=(a,f)=>{const m={};for(const x of a){const k=f(x);(m[k]=m[k]||[]).push(x)}return m};
const indexBy=(a,k='id')=>{const m={};for(const x of a)m[x[k]]=x;return m};
const initials=n=>String(n||'?').trim().split(/\s+/).slice(0,2).map(w=>w[0]).join('').toUpperCase();
const loadScript=(src)=>new Promise((res,rej)=>{if(document.querySelector('script[data-src="'+src+'"]'))return res();const s=document.createElement('script');s.src=src;s.dataset.src=src;s.onload=res;s.onerror=()=>rej(new Error('Gagal memuat pustaka: '+src));document.head.appendChild(s)});
function saveFile(name,blob){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},500)}
function toCSV(headers,rows){const q=v=>{v=v==null?'':String(v);return/[",\n;]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v};return'﻿'+[headers,...rows].map(r=>r.map(q).join(',')).join('\r\n')}
async function exportRows(fname,sheet,headers,rows){
  try{await loadScript(XLSX_URL);const ws=XLSX.utils.aoa_to_sheet([headers,...rows]);ws['!cols']=headers.map((h,i)=>({wch:Math.min(40,Math.max(String(h).length,...rows.slice(0,200).map(r=>String(r[i]==null?'':r[i]).length))+2)}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,String(sheet).slice(0,30));XLSX.writeFile(wb,fname+'.xlsx');toast('File Excel diunduh','ok')}
  catch(e){saveFile(fname+'.csv',new Blob([toCSV(headers,rows)],{type:'text/csv;charset=utf-8'}));toast('Pustaka Excel tidak tersedia, file diunduh sebagai CSV (dapat dibuka di Excel)','warn')}
}
function printTable(title,sub,headers,rows){
  let p=$('#sk-print');if(!p){p=document.createElement('div');p.id='sk-print';document.body.appendChild(p)}
  const school=(S.settings&&S.settings['school.name'])||APP_NAME;
  p.innerHTML='<h2>'+esc(school)+' - '+esc(title)+'</h2><p>'+esc(sub||'')+' &middot; dicetak '+esc(fmtDT(NOW()))+' oleh '+esc(S.admin?S.admin.full_name:(S.me?S.me.full_name:''))+'</p><table><thead><tr>'+headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr>'+r.map(c=>'<td>'+esc(c==null?'':c)+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
  setTimeout(()=>window.print(),80);
}
const NOW=()=>Date.now()+(S.skew||0);
const TODAY=()=>D.local(NOW()).date;

/* =====================================================================
   2. IKON (SVG inline, tanpa pustaka luar)
   ===================================================================== */
const ICONS={
home:'<path d="M3 11.5 12 4l9 7.5"/><path d="M5 10v10h5v-6h4v6h5V10"/>',
users:'<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c.5-3.5 3.2-5.5 6.5-5.5s6 2 6.5 5.5"/><circle cx="17.5" cy="9" r="2.5"/><path d="M17 14.5c2.4.3 4 2 4.5 5"/>',
user:'<circle cx="12" cy="8" r="4"/><path d="M4 21c.7-4 4-6 8-6s7.3 2 8 6"/>',
building:'<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M9 7h2M13 7h2M9 11h2M13 11h2M9 15h2M13 15h2"/>',
calendar:'<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
clock:'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
check:'<path d="m5 12.5 4.5 4.5L19 7.5"/>',
x:'<path d="M6 6l12 12M18 6 6 18"/>',
plus:'<path d="M12 5v14M5 12h14"/>',
edit:'<path d="M4 20h4l10.5-10.5a2.1 2.1 0 0 0-3-3L5 17v3z"/>',
trash:'<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>',
search:'<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
bell:'<path d="M6 16V11a6 6 0 0 1 12 0v5l2 2H4z"/><path d="M10 20a2 2 0 0 0 4 0"/>',
logout:'<path d="M15 4h4a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1h-4"/><path d="M10 8l-4 4 4 4M6 12h10"/>',
settings:'<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9 7 7M17 17l2.1 2.1M4.9 19.1 7 17M17 7l2.1-2.1"/>',
file:'<path d="M6 3h8l5 5v13H6z"/><path d="M14 3v5h5M9 13h6M9 17h6"/>',
upload:'<path d="M12 16V5M7 10l5-5 5 5M4 19h16"/>',
download:'<path d="M12 4v11M7 11l5 5 5-5M4 20h16"/>',
chart:'<path d="M4 20V4M4 20h16"/><path d="M8 16v-4M12 16V8M16 16v-6"/>',
shield:'<path d="M12 3 5 6v5c0 5 3 8 7 10 4-2 7-5 7-10V6z"/>',
lock:'<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>',
key:'<circle cx="8" cy="15" r="4"/><path d="m11 12 8-8M16 7l3 3"/>',
refresh:'<path d="M20 12a8 8 0 1 1-2.5-5.8M20 4v5h-5"/>',
menu:'<path d="M4 7h16M4 12h16M4 17h16"/>',
grid:'<rect x="4" y="4" width="6.5" height="6.5" rx="1.8"/><rect x="13.5" y="4" width="6.5" height="6.5" rx="1.8"/><rect x="4" y="13.5" width="6.5" height="6.5" rx="1.8"/><rect x="13.5" y="13.5" width="6.5" height="6.5" rx="1.8"/>',
more:'<circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/>',
left:'<path d="m15 5-7 7 7 7"/>',right:'<path d="m9 5 7 7-7 7"/>',
alert:'<path d="M12 4 2.5 20h19z"/><path d="M12 10v4M12 17.5v.5"/>',
info:'<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.5"/>',
list:'<path d="M8 6h12M8 12h12M8 18h12M4 6h.5M4 12h.5M4 18h.5"/>',
inbox:'<path d="M3 13l3-8h12l3 8v6H3z"/><path d="M3 13h5l1 3h6l1-3h5"/>',
swap:'<path d="M4 8h14l-3-3M20 16H6l3 3"/>',
print:'<path d="M7 9V3h10v6M7 17H4v-7h16v7h-3M7 14h10v7H7z"/>',
eye:'<path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/>',
pin:'<path d="M12 21s7-6 7-11a7 7 0 0 0-14 0c0 5 7 11 7 11z"/><circle cx="12" cy="10" r="2.5"/>',
brief:'<rect x="3" y="7" width="18" height="13" rx="2"/><path d="M9 7V4h6v3M3 13h18"/>',
heart:'<path d="M12 20s-8-5-8-11a4.5 4.5 0 0 1 8-2.5A4.5 4.5 0 0 1 20 9c0 6-8 11-8 11z"/>',
finger:'<path d="M12 4a6 6 0 0 0-6 6v3M12 8a2.5 2.5 0 0 0-2.5 2.5V16M12 8a2.5 2.5 0 0 1 2.5 2.5V20M18 11v3c0 2-.5 4-1.5 6M9.5 20c0-1.5-.5-2.5-.5-4"/>',
sun:'<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M2 12h2M20 12h2M5 5l1.5 1.5M17.5 17.5 19 19M5 19l1.5-1.5M17.5 6.5 19 5"/>',
moon:'<path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"/>',
book:'<path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2z"/><path d="M4 19V5"/>',
copy:'<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2"/>',
history:'<path d="M3 12a9 9 0 1 0 3-6.7M3 4v5h5M12 8v4l3 2"/>',
school:'<path d="M2 9l10-5 10 5-10 5z"/><path d="M6 11.5V16c0 1.5 3 3 6 3s6-1.5 6-3v-4.5M22 9v6"/>'
};
const ic=(n,c='')=>'<svg class="sk-ic '+c+'" viewBox="0 0 24 24" aria-hidden="true">'+(ICONS[n]||'')+'</svg>';

/* =====================================================================
   3. STATE, STATUS, LABEL
   ===================================================================== */
const S={mode:null,admin:null,me:null,token:null,ctx:null,range:null,route:'',skew:0,settings:{},sb:null,menus:[],cache:{},booted:false};
const ST={present:['Hadir','g'],late:['Terlambat','a'],incomplete:['Belum Lengkap','o'],leave:['Cuti','b'],sick:['Sakit','v'],permit:['Izin','c'],official_duty:['Dinas','i'],holiday:['Libur','s'],day_off:['Libur Jadwal','s'],absent:['Alpha','r'],pending:['Menunggu','s'],no_schedule:['Tanpa Jadwal','s']};
const RST={waiting_l1:['Menunggu Kepala Dept.','a'],waiting_l2:['Menunggu HRD','b'],approved:['Disetujui','g'],rejected:['Ditolak','r'],cancelled:['Dibatalkan','s']};
const SWST={waiting_partner:['Menunggu pengganti','a'],waiting_coordinator:['Menunggu HRD','b'],approved:['Disetujui','g'],rejected:['Ditolak','r'],cancelled:['Dibatalkan','s']};
const ROLE={admin_sistem:'Super Admin',manajer_hrd:'HRD Manager',staf_hrd:'HRD Staff / Admin HR',kepala_dept:'Manager / Kepala Departemen',supervisor:'Supervisor / Atasan Langsung',pegawai:'Karyawan / Staff'};
const ROLE_INFO={admin_sistem:'Pengaturan sistem, akun, role, dan konfigurasi aplikasi.',manajer_hrd:'Approval akhir, monitoring seluruh karyawan, validasi kebijakan, dan laporan strategis HR.',staf_hrd:'Mengelola data karyawan, memeriksa absensi, memvalidasi pengajuan, dan memproses administrasi.',kepala_dept:'Memantau tim dan memberikan persetujuan tahap pertama sesuai kewenangan.',supervisor:'Memeriksa dan memberikan persetujuan awal jika alur pengajuan memerlukannya.',pegawai:'Edit data pribadi, absensi masuk-pulang, melihat riwayat, dan mengajukan izin atau cuti, dll.'};
const CAT={leave:'Cuti',permit:'Izin',sick:'Sakit',official_duty:'Dinas Luar',time_permit:'Izin Jam',correction:'Koreksi Absen',overtime:'Lembur'};
const TYPE={teacher:'Guru',staff:'Karyawan Umum'};
const EMPST={permanent:'Tetap',contract:'Kontrak',honorary:'Honorer'};
const GENDER={L:'Laki-laki',P:'Perempuan'};
const MARITAL={belum_menikah:'Belum Menikah',menikah:'Menikah',cerai_hidup:'Cerai Hidup',cerai_mati:'Cerai Mati'};
const EDU=['SD','SMP','SMA/SMK','D1','D2','D3','D4','S1','S2','S3'];
/* kolom tambahan tabel employees (butuh migrasi-karyawan.sql) */
const EMP_X=['gender','birth_place','birth_date','contract_end_date','education_level','education_institution','education_major','bpjs_kesehatan','bpjs_ketenagakerjaan','marital_status'];
const HOLT={national:'Libur Nasional',collective_leave:'Cuti Bersama',school:'Libur Lembaga',academic_break:'Libur Akademik'};
const bd=(t,c='s')=>'<span class="sk-bd '+c+'">'+esc(t)+'</span>';
const stBadge=s=>{const x=ST[s]||[s,'s'];return bd(x[0],x[1])};
const rstBadge=s=>{const x=RST[s]||[s,'s'];return bd(x[0],x[1])};

/* =====================================================================
   4. UI KIT
   ===================================================================== */
function toast(msg,type=''){const w=$('#sk-toasts');if(!w)return;const t=document.createElement('div');t.className='sk-toast '+type;t.textContent=msg;w.appendChild(t);setTimeout(()=>t.remove(),type==='bad'?6500:3800)}
function errMsg(e){let m=(e&&e.message)||String(e);m=m.replace(/^SESI_TIDAK_VALID:\s*/,'').replace(/^PERIOD_LOCKED:\s*/,'');return m}
function fail(e){console.error(e);const raw=(e&&e.message)||'';if(/SESI_TIDAK_VALID/.test(raw)){toast('Sesi berakhir, silakan login kembali','warn');doLogout(true);return}toast(errMsg(e),'bad')}
function modal(o){
  const ov=document.createElement('div');ov.className='sk-ov';
  ov.innerHTML='<div class="sk-md '+(o.size||'')+'" role="dialog" aria-modal="true"><div class="h"><h3>'+esc(o.title||'')+'</h3><button class="sk-ibtn" data-x style="width:34px;height:34px" aria-label="Tutup">'+ic('x')+'</button></div><div class="b">'+(o.body||'')+'</div>'+((o.actions&&o.actions.length)?'<div class="f"></div>':'')+'</div>';
  $('#sk-root').appendChild(ov);
  const api={el:ov,close(){ov.remove();if(o.onClose)o.onClose()},body:$('.b',ov),$:s=>$(s,ov),$$:s=>$$(s,ov)};
  $('[data-x]',ov).onclick=()=>api.close();
  if(!o.lock)ov.addEventListener('mousedown',e=>{if(e.target===ov)api.close()});
  const f=$('.f',ov);
  (o.actions||[]).forEach(a=>{const b=document.createElement('button');b.className='sk-b '+(a.c==null?'':a.c);b.innerHTML=(a.i?ic(a.i,'sm'):'')+esc(a.t);b.onclick=async()=>{if(b.disabled)return;b.disabled=true;try{const r=await a.fn?.(api);if(a.keep!==true&&r!==false)api.close()}catch(e){fail(e)}finally{b.disabled=false}};f.appendChild(b)});
  if(o.onMount)o.onMount(api);
  const first=$('input:not([type=hidden]),select,textarea',ov);if(first&&!o.noFocus)setTimeout(()=>first.focus(),30);
  return api;
}
function confirmBox(msg,o={}){return new Promise(res=>{let done=false;const m=modal({title:o.title||'Konfirmasi',body:'<p>'+esc(msg).replace(/\n/g,'<br>')+'</p>',onClose:()=>{if(!done){done=true;res(false)}},actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:o.ok||'Ya, lanjutkan',c:o.danger?'bad':'',i:o.danger?'trash':'check',fn:()=>{done=true;res(true)}}]})})}
function promptBox(title,label,o={}){return new Promise(res=>{let done=false;modal({title,body:'<label class="sk-f"><span>'+esc(label)+'</span>'+(o.multi===false?'<input class="sk-in" id="pb-v" value="'+esc(o.value||'')+'">':'<textarea class="sk-ta" id="pb-v">'+esc(o.value||'')+'</textarea>')+'</label>'+(o.hint?'<div class="sk-hint">'+esc(o.hint)+'</div>':''),onClose:()=>{if(!done){done=true;res(null)}},actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:o.ok||'Simpan',i:'check',c:o.danger?'bad':'',fn:m=>{const v=$('#pb-v',m.el).value.trim();if(o.required!==false&&!v){toast('Isian wajib diisi','warn');return false}done=true;res(v)}}]})})}
const fmtDur=ms=>{const s=Math.max(0,Math.round(ms/1000));return s<60?s+' dtk':Math.floor(s/60)+' mnt '+(s%60)+' dtk'};
const etaTxt=(t0,d,n)=>{const el=Date.now()-t0;return d>0&&d<n&&el>2500?' \u00b7 sisa \u00b1'+fmtDur(el*(n-d)/d):''};
/* jendela progres (persen + tahap + waktu berjalan) untuk proses panjang */
function progUI(title){
  const m=modal({title,lock:true,noFocus:true,body:'<div class="ip"><div class="ip-lb" id="ip-lb">Menyiapkan...</div><div class="ip-bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"><i id="ip-fill"></i></div><div class="ip-row"><b id="ip-pct">0%</b><span id="ip-dt"></span></div><div class="ip-el" id="ip-el"></div><div class="ip-hint">Jangan menutup atau memuat ulang halaman sampai proses selesai.</div></div>',actions:[]});
  const x=$('[data-x]',m.el);if(x)x.style.display='none';
  const t0=Date.now(),bar=$('.ip-bar',m.el),fill=$('#ip-fill',m.el),pc=$('#ip-pct',m.el),lb=$('#ip-lb',m.el),dt=$('#ip-dt',m.el),el=$('#ip-el',m.el);
  const tick=setInterval(()=>{el.textContent='Berjalan '+fmtDur(Date.now()-t0)},500),warn=e=>{e.preventDefault();e.returnValue=''};
  window.addEventListener('beforeunload',warn);
  return{set(pct,label,detail){const v=Math.max(0,Math.min(100,Math.round(pct)));fill.style.width=v+'%';pc.textContent=v+'%';bar.setAttribute('aria-valuenow',v);if(label!=null)lb.textContent=label;if(detail!=null)dt.textContent=detail},
    close(){clearInterval(tick);window.removeEventListener('beforeunload',warn);m.close()}}
}
async function busy(btn,fn){if(btn)btn.disabled=true;try{return await fn()}catch(e){fail(e)}finally{if(btn)btn.disabled=false}}
const loading=(msg='Memuat data...')=>'<div class="sk-load"><div class="sk-spin"></div>'+esc(msg)+'</div>';
function formData(f){const o={};$$('[name]',f).forEach(el=>{if(el.type==='checkbox')o[el.name]=el.checked;else if(el.type==='radio'){if(el.checked)o[el.name]=el.value}else if(el.type==='file')o[el.name]=el.files;else o[el.name]=el.value});return o}
const fld=(label,inner,hint)=>'<label class="sk-f"><span>'+esc(label)+'</span>'+inner+(hint?'<div class="sk-hint">'+esc(hint)+'</div>':'')+'</label>';
const inp=(name,val='',o={})=>'<input class="sk-in" name="'+name+'" type="'+(o.type||'text')+'" value="'+esc(val)+'"'+(o.ph?' placeholder="'+esc(o.ph)+'"':'')+(o.req?' required':'')+(o.dis?' disabled':'')+(o.min!=null?' min="'+o.min+'"':'')+(o.max!=null?' max="'+o.max+'"':'')+(o.step?' step="'+o.step+'"':'')+(o.attr?' '+o.attr:'')+'>';
const sel=(name,opts,val,o={})=>'<select class="sk-sel" name="'+name+'"'+(o.req?' required':'')+(o.dis?' disabled':'')+(o.attr?' '+o.attr:'')+'>'+(o.blank!=null?'<option value="">'+esc(o.blank)+'</option>':'')+opts.map(x=>{const v=Array.isArray(x)?x[0]:x,l=Array.isArray(x)?x[1]:x;return'<option value="'+esc(v)+'"'+(String(v)===String(val)?' selected':'')+'>'+esc(l)+'</option>'}).join('')+'</select>';
const chk=(name,label,val,x)=>'<label class="sk-chk"><input type="checkbox" name="'+name+'"'+(val?' checked':'')+(x?' '+x:'')+'><span>'+esc(label)+'</span></label>';
const TBL_PAGE=25,TBL_REG=new Map();let TBL_SEQ=0;
function tblRow(headers,r){return'<tr>'+r.map((c,i)=>{const h=headers[i];const cl=(h&&h.c)||'';return'<td class="'+cl+'" data-label="'+esc(h?String(typeof h==='string'?h:h.t).replace(/<[^>]+>/g,''):'')+'">'+(c==null?'':c)+'</td>'}).join('')+'</tr>'}
function tbl(headers,rows,o={}){
  if(!rows.length)return'<div class="sk-empty">'+(o.empty||'Belum ada data')+'</div>';
  const per=o.ef?efPer(o.ef,o.page||TBL_PAGE):(o.page||TBL_PAGE),more=rows.length>per;let id='';
  if(more){id='tp'+(++TBL_SEQ);TBL_REG.set(id,{h:headers,r:rows,n:per,per});if(TBL_REG.size>60)TBL_REG.delete(TBL_REG.keys().next().value)}
  return'<div class="sk-tw"'+(o.maxh?' style="max-height:'+o.maxh+'"':'')+'><table class="sk-t '+(o.cls||'')+'"><thead><tr>'+headers.map(h=>{const x=typeof h==='string'?{t:h}:h;return'<th class="'+(x.c||'')+'">'+x.t+'</th>'}).join('')+'</tr></thead><tbody>'+rows.slice(0,per).map(r=>tblRow(headers,r)).join('')+'</tbody></table>'
    +(more?'<div class="sk-more" data-tp="'+id+'"><span>'+per+' dari '+rows.length+' baris</span><button class="sk-b sec sm" data-act="tblMore" data-id="'+id+'">Tampilkan '+Math.min(per,rows.length-per)+' lagi</button></div>':'')+'</div>';
}
function tblMore(btn){
  const w=btn.closest('.sk-more'),st=TBL_REG.get(btn.dataset.id);if(!w||!st){if(w)w.remove();return}
  const tb=w.parentNode.querySelector('tbody'),to=Math.min(st.r.length,st.n+st.per);
  tb.insertAdjacentHTML('beforeend',st.r.slice(st.n,to).map(r=>tblRow(st.h,r)).join(''));st.n=to;
  if(to>=st.r.length){w.remove();TBL_REG.delete(btn.dataset.id)}
  else{w.firstChild.textContent=to+' dari '+st.r.length+' baris';btn.textContent='Tampilkan '+Math.min(st.per,st.r.length-to)+' lagi'}
}
const card=(title,body,actions='')=>'<section class="sk-card"><div class="hd"><h3>'+title+'</h3>'+actions+'</div>'+body+'</section>';
const kpi=(icon,color,val,label,sub,act)=>'<div class="sk-kpi'+(act?' link':'')+'"'+(act?' data-act="nav" data-k="'+act+'"':'')+'><div class="i" style="background:'+color+'">'+ic(icon)+'</div><div><b>'+val+'</b><small>'+esc(label)+'</small>'+(sub?'<span class="s2">'+esc(sub)+'</span>':'')+'</div></div>';
function skeleton(){return'<div class="sk-skel"><div class="r1"><i></i><i></i><i></i><i></i></div><i class="b1"></i></div>'}
const stabs=(items,cur,act)=>'<div class="sk-stab">'+items.map(([k,l])=>'<button data-act="'+act+'" data-k="'+k+'" class="'+(k===cur?'on':'')+'">'+esc(l)+'</button>').join('')+'</div>';
const dot=n=>n>0?'<span class="sk-dot">'+(n>99?'99+':n)+'</span>':'';

/* delegasi event */
const ACT={},FORM={},CHG={},INP={};
document.addEventListener('click',e=>{const t=e.target.closest('[data-act]');if(!t||!$('#sk-root').contains(t))return;e.preventDefault();const fn=ACT[t.dataset.act];if(fn)Promise.resolve().then(()=>fn(t,e)).catch(fail)});
document.addEventListener('submit',e=>{const f=e.target.closest('[data-form]');if(!f||!$('#sk-root').contains(f))return;e.preventDefault();const fn=FORM[f.dataset.form];if(fn){const b=$('button[type=submit]',f);Promise.resolve().then(()=>busy(b,()=>fn(f,formData(f),e)))}});
document.addEventListener('change',e=>{const t=e.target.closest('[data-chg]');if(!t||!$('#sk-root').contains(t))return;const fn=CHG[t.dataset.chg];if(fn)Promise.resolve().then(()=>fn(t,e)).catch(fail)});
document.addEventListener('input',e=>{const t=e.target.closest('[data-inp]');if(!t||!$('#sk-root').contains(t))return;const fn=INP[t.dataset.inp];if(fn)fn(t,e)});


/* =====================================================================
   5. MESIN PERHITUNGAN (jadwal efektif, rekap harian, cuti, validasi)
      Logika yang sama dengan fungsi SQL; dipakai untuk menghitung rekap
      kehadiran, tampilan jadwal, laporan, dan mode Preview.
   ===================================================================== */
const Engine={
  set(ctx,k,d){const v=ctx.settings&&ctx.settings[k];return v==null||v===''?d:v},
  ix(ctx){
    if(ctx._ix)return ctx._ix;
    const ix={emp:indexBy(ctx.employees||[]),sa:{},assign:{},wsd:{},ws:indexBy(ctx.work_schedules||[]),shift:indexBy(ctx.shifts||[]),hol:{},logs:{},req:{},rt:indexBy(ctx.request_types||[]),att:{}};
    for(const a of ctx.shift_assignments||[])ix.sa[a.employee_id+'|'+String(a.work_date).slice(0,10)]=a;
    for(const a of ctx.schedule_assignments||[])(ix.assign[a.employee_id]=ix.assign[a.employee_id]||[]).push(a);
    for(const k in ix.assign)ix.assign[k].sort((a,b)=>String(b.effective_from).localeCompare(String(a.effective_from)));
    for(const w of ctx.work_schedule_days||[])ix.wsd[w.work_schedule_id+'|'+w.weekday]=w;
    for(const h of ctx.holidays||[])(ix.hol[String(h.holiday_date).slice(0,10)]=ix.hol[String(h.holiday_date).slice(0,10)]||[]).push(h);
    for(const l of ctx.raw_logs||[]){if(!l.employee_id)continue;(ix.logs[l.employee_id]=ix.logs[l.employee_id]||[]).push(Date.parse(l.logged_at))}
    for(const k in ix.logs)ix.logs[k].sort((a,b)=>a-b);
    for(const r of ctx.requests||[])if(r.status==='approved')(ix.req[r.employee_id]=ix.req[r.employee_id]||[]).push(r);
    for(const a of ctx.attendance_records||[])ix.att[a.employee_id+'|'+String(a.work_date).slice(0,10)]=a;
    return ctx._ix=ix;
  },
  touch(ctx){ctx._ix=null},
  mk(src,ds,st,en,brk,tol,early,late,name,color,shiftId){
    st=D.hm(st);en=D.hm(en);
    const startMs=D.epoch(ds,st),endMs=D.epoch(en<=st?D.add(ds,1):ds,en);
    return{kind:'work',src,start:st,end:en,startMs,endMs,brk:brk||0,tol:tol==null?10:tol,early:early==null?120:early,late:late==null?240:late,name,color,shiftId};
  },
  sched(ctx,emp,ds,o={}){
    const ix=this.ix(ctx),e=ix.emp[emp],none={kind:'none',src:null};
    if(!e)return none;
    if(ds<String(e.join_date).slice(0,10)||(e.end_date&&ds>String(e.end_date).slice(0,10)))return none;
    const sa=ix.sa[emp+'|'+ds];
    if(sa&&(sa.status==='published'||o.draft)){
      if(!sa.shift_id)return{kind:'day_off',src:'shift',draft:sa.status!=='published'};
      const sh=ix.shift[sa.shift_id];if(!sh)return none;
      const r=this.mk('shift',ds,sh.start_time,sh.end_time,sh.break_minutes,sh.tolerance_minutes,sh.early_clock_in_minutes,sh.late_clock_out_minutes,sh.name,sh.color,sh.id);r.draft=sa.status!=='published';return r;
    }
    const as=(ix.assign[emp]||[]).find(a=>String(a.effective_from).slice(0,10)<=ds&&(!a.effective_to||String(a.effective_to).slice(0,10)>=ds));
    if(!as)return none;
    const ws=ix.ws[as.work_schedule_id];if(!ws)return none;
    const hol=(ix.hol[ds]||[]).find(h=>h.applies_to==='all'||h.applies_to===e.type);
    if(hol)return{kind:'holiday',src:'normal',name:hol.name};
    const row=ix.wsd[ws.id+'|'+D.iso(ds)];
    if(!row||!row.is_workday||!row.start_time)return{kind:'day_off',src:'normal'};
    return this.mk('normal',ds,row.start_time,row.end_time,row.break_minutes,ws.tolerance_minutes,ws.early_clock_in_minutes,ws.late_clock_out_minutes,ws.name);
  },
  isWorkday(ctx,emp,ds){return this.sched(ctx,emp,ds).kind==='work'},
  claim(ctx,emp,ds,s){
    const ix=this.ix(ctx),all=ix.logs[emp]||[];if(!all.length)return[];
    const nb=[D.add(ds,-1),D.add(ds,1)].map(d=>this.sched(ctx,emp,d)).filter(x=>x.kind==='work');
    const inWin=(t,n)=>t>=n.startMs-n.early*6e4&&t<=n.endMs+n.late*6e4;
    let mine;
    if(s.kind==='work'){
      mine=all.filter(t=>inWin(t,s));
      mine=mine.filter(t=>{const dist=Math.min(Math.abs(t-s.startMs),Math.abs(t-s.endMs));return nb.every(n=>!inWin(t,n)||dist<=Math.min(Math.abs(t-n.startMs),Math.abs(t-n.endMs)))});
    }else{
      const a=D.epoch(ds,'00:00'),b=a+864e5;
      mine=all.filter(t=>t>=a&&t<b&&!nb.some(n=>inWin(t,n)));
    }
    const out=[];for(const t of mine)if(!out.length||t-out[out.length-1]>=60000)out.push(t);
    return out;
  },
  /* hitung rekap satu pegawai satu tanggal. Mengembalikan objek siap simpan ke attendance_records */
  recalcDay(ctx,emp,ds,nowMs){
    nowMs=nowMs||NOW();
    const ix=this.ix(ctx),s=this.sched(ctx,emp,ds),today=D.local(nowMs).date;
    const rec={employee_id:emp,work_date:ds,schedule_type:s.src||'none',scheduled_start:null,scheduled_end:null,check_in_at:null,check_out_at:null,status:'pending',late_minutes:0,early_leave_minutes:0,worked_minutes:null,overtime_minutes:0,request_id:null,is_manual:false,note:null};
    const reqs=(ix.req[emp]||[]).filter(r=>String(r.start_date).slice(0,10)<=ds&&String(r.end_date).slice(0,10)>=ds);
    const cat=r=>(ix.rt[r.request_type_id]||{}).category;
    const minGap=+this.set(ctx,'min_in_out_gap_minutes',30);
    const logs=this.claim(ctx,emp,ds,s);
    let inMs=null,outMs=null;
    if(logs.length){
      inMs=logs[0];
      if(logs.length>1&&logs[logs.length-1]-logs[0]>=minGap*6e4)outMs=logs[logs.length-1];
      else if(s.kind==='work'&&Math.abs(logs[0]-s.startMs)>Math.abs(logs[0]-s.endMs)){outMs=logs[0];inMs=null}
    }
    const notes=[];
    const corr=reqs.find(r=>cat(r)==='correction'&&String(r.start_date).slice(0,10)===ds);
    if(corr&&corr.correction_payload){
      const cp=corr.correction_payload;
      if(cp.in){inMs=D.epoch(ds,cp.in)}
      if(cp.out){let od=ds;if(cp.in&&cp.out<=cp.in&&s.kind==='work'&&s.endMs>s.startMs&&D.hm(s.end)<=D.hm(s.start))od=D.add(ds,1);else if(!cp.in&&s.kind==='work'&&s.end<=s.start)od=D.add(ds,1);outMs=D.epoch(od,cp.out)}
      notes.push('Koreksi absen disetujui');rec.request_id=corr.id;
    }
    rec.check_in_at=inMs!=null?new Date(inMs).toISOString():null;
    rec.check_out_at=outMs!=null?new Date(outMs).toISOString():null;
    if(s.kind==='work'){
      rec.scheduled_start=new Date(s.startMs).toISOString();rec.scheduled_end=new Date(s.endMs).toISOString();
      const lateP=reqs.find(r=>cat(r)==='time_permit'&&!/pulang/i.test((ix.rt[r.request_type_id]||{}).code||'')&&r.start_time);
      const earlyP=reqs.find(r=>cat(r)==='time_permit'&&/pulang/i.test((ix.rt[r.request_type_id]||{}).code||'')&&r.start_time);
      const tol=s.tol*6e4,fromTol=this.set(ctx,'late_calc_from','start')==='tolerance_end';
      if(inMs!=null){
        let base=s.startMs;if(lateP)base=Math.max(base,D.epoch(ds,D.hm(lateP.start_time)));
        if(inMs>base+tol)rec.late_minutes=Math.floor((inMs-base-(fromTol?tol:0))/6e4);
      }
      if(outMs!=null){
        let base=s.endMs;if(earlyP)base=Math.min(base,D.epoch(ds,D.hm(earlyP.start_time)));
        if(outMs<base)rec.early_leave_minutes=Math.floor((base-outMs)/6e4);
      }
      if(inMs!=null&&outMs!=null)rec.worked_minutes=Math.max(0,Math.floor((outMs-inMs)/6e4)-s.brk);
      const winEnd=s.endMs+s.late*6e4;
      if(inMs!=null&&outMs!=null)rec.status=rec.late_minutes>0?'late':'present';
      else if(inMs!=null||outMs!=null){
        if(ds===today&&nowMs<winEnd){rec.status=inMs!=null?(rec.late_minutes>0?'late':'present'):'pending';notes.push(inMs!=null?'Belum absen pulang':'')}
        else if(ds>today)rec.status='pending';else rec.status='incomplete';
      }else rec.status=(ds<today||(ds===today&&nowMs>winEnd))?'absent':'pending';
      // lembur
      const ov=reqs.find(r=>cat(r)==='overtime'&&String(r.start_date).slice(0,10)===ds);
      if(ov&&outMs!=null&&outMs>s.endMs){
        const ap=ov.start_time&&ov.end_time?(D.epoch(ds,D.hm(ov.end_time))-D.epoch(ds,D.hm(ov.start_time)))/6e4:0;
        const act=Math.floor((outMs-s.endMs)/6e4),rnd=+this.set(ctx,'overtime_round_minutes',15);
        rec.overtime_minutes=Math.floor(Math.min(ap,act)/rnd)*rnd;rec.request_id=rec.request_id||ov.id;
      }
      // pengajuan yang menutup hari
      const cover=reqs.filter(r=>['leave','sick','permit','official_duty'].includes(cat(r)));
      if(cover.length){
        const r=cover[0],t=ix.rt[r.request_type_id]||{},map={leave:'leave',sick:'sick',permit:'permit',official_duty:'official_duty'};
        const half=r.day_part&&r.day_part!=='full';
        if(!half||(inMs==null&&outMs==null)){
          if(logs.length)notes.push('Ada absen pada hari '+(t.name||'izin').toLowerCase());
          rec.status=map[t.category];rec.late_minutes=0;rec.early_leave_minutes=0;rec.request_id=r.id;
          if(t.category==='official_duty'&&rec.worked_minutes==null)rec.worked_minutes=null;
        }
      }
    }else{
      rec.status=s.kind==='none'?'no_schedule':s.kind;
      if(logs.length){notes.push('Ada absen pada hari libur');if(inMs!=null&&outMs!=null)rec.worked_minutes=Math.floor((outMs-inMs)/6e4);
        const ov=reqs.find(r=>cat(r)==='overtime'&&String(r.start_date).slice(0,10)===ds);
        if(ov&&rec.worked_minutes){const ap=ov.start_time&&ov.end_time?(D.epoch(ds,D.hm(ov.end_time))-D.epoch(ds,D.hm(ov.start_time)))/6e4:0,rnd=+this.set(ctx,'overtime_round_minutes',15);rec.overtime_minutes=Math.floor(Math.min(ap,rec.worked_minutes)/rnd)*rnd;rec.request_id=ov.id}}
    }
    rec.note=notes.filter(Boolean).join('; ')||null;
    return rec;
  },
  recalcRange(ctx,empIds,from,to,nowMs){
    nowMs=nowMs||NOW();const ix=this.ix(ctx),today=D.local(nowMs).date,out=[];
    for(const emp of empIds){
      const e=ix.emp[emp];if(!e||(!e.is_active&&!e.end_date))continue;
      for(const ds of D.range(from,to)){
        if(ds>today)break;
        if(ds<String(e.join_date).slice(0,10))continue;if(e.end_date&&ds>String(e.end_date).slice(0,10))continue;
        const ex=ix.att[emp+'|'+ds];if(ex&&(ex.is_locked||ex.is_manual))continue;
        out.push(this.recalcDay(ctx,emp,ds,nowMs));
      }
    }
    return out;
  },
  /* ---------- cuti ---------- */
  /* siklus cuti: 26 Juni - 25 Juni (saldo hangus tiap 25 Juni). Kunci periode: '2026/2027' */
  leaveStart(ds){const p=D.parse(ds),y=(p.m>6||(p.m===6&&p.d>=26))?p.y:p.y-1;return y+'-06-26'},
  leaveKey(ds){const y=+this.leaveStart(ds).slice(0,4);return y+'/'+(y+1)},
  svcMonths(join,ds){const a=D.parse(ds),b=D.parse(String(join).slice(0,10));let mo=(a.y-b.y)*12+(a.m-b.m);if(a.d<b.d)mo--;return mo},
  periodKey(ctx,ds){
    const _m=this.set(ctx,'leave_period_mode','calendar');if(_m==='jun25')return this.leaveKey(ds);
    if(_m==='academic'){
      const ay=(ctx.academic_years||[]).filter(a=>String(a.start_date).slice(0,10)<=ds&&String(a.end_date).slice(0,10)>=ds).sort((a,b)=>String(b.start_date).localeCompare(String(a.start_date)))[0];
      if(ay)return ay.name;
    }
    return String(D.parse(ds).y);
  },
  countDays(ctx,emp,rt,start,end,part){
    let n=0;for(const ds of D.range(start,end)){if(rt.count_mode==='calendar_days')n++;else if(this.isWorkday(ctx,emp,ds))n++}
    if(start===end&&part&&part!=='full'&&rt.allow_half_day&&n>=1)n=0.5;return n;
  },
  ledger(ctx){return ctx.ledger||ctx.leave_balance_entries||[]},
  balance(ctx,emp,type,period){return sum(this.ledger(ctx).filter(l=>l.employee_id===emp&&l.request_type_id===type&&l.period_key===period),l=>+l.days)},
  reserved(ctx,emp,type,period,excl){return sum((ctx.requests||[]).filter(r=>r.employee_id===emp&&r.request_type_id===type&&['waiting_l1','waiting_l2'].includes(r.status)&&this.periodKey(ctx,String(r.start_date).slice(0,10))===period&&r.id!==excl),r=>+r.total_days)},
  /* cerminan fn_validate_request (dipakai mode Preview) */
  validateRequest(ctx,emp,p,excl){
    const ix=this.ix(ctx),e=ix.emp[emp],rt=(ctx.request_types||[]).find(t=>t.id===+p.type_id&&t.is_active!==false);
    if(!rt)return{ok:false,errors:['Jenis pengajuan tidak valid'],total_days:0};
    const err=[],today=TODAY();let start=p.start_date,end=p.end_date||p.start_date,part=p.day_part||'full',total=0,bal=0,res=0,avail=0,period=null;
    const reason=String(p.reason||'').trim(),att=String(p.attachment_url||'').trim();
    if(!start){err.push('Tanggal wajib diisi');start=end=today}
    if(end<start){err.push('Tanggal selesai tidak boleh sebelum tanggal mulai');end=start}
    if(!reason)err.push('Alasan wajib diisi');
    const active=(ctx.requests||[]).filter(r=>r.employee_id===emp&&['waiting_l1','waiting_l2','approved'].includes(r.status)&&r.id!==excl);
    const catOf=r=>(ix.rt[r.request_type_id]||{}).category;
    if(['leave','permit','sick','official_duty'].includes(rt.category)){
      total=this.countDays(ctx,emp,rt,start,end,part);
      if(total===0)err.push('Semua tanggal yang dipilih adalah hari libur atau bukan hari kerja Anda');
      if(active.some(r=>['leave','permit','sick','official_duty'].includes(catOf(r))&&String(r.start_date).slice(0,10)<=end&&String(r.end_date).slice(0,10)>=start&&!(r.start_date===r.end_date&&start===end&&r.day_part!=='full'&&part!=='full'&&r.day_part!==part)))err.push('Tanggal bertumpuk dengan pengajuan lain yang masih aktif');
      const jd=String(e.join_date).slice(0,10),a=D.parse(start),b=D.parse(jd);let months=(a.y-b.y)*12+(a.m-b.m);if(a.d<b.d)months--;
      if(rt.min_service_months>0&&months<rt.min_service_months)err.push('Masa kerja minimal '+rt.min_service_months+' bulan untuk jenis ini');
      if(rt.min_notice_days>0&&D.diff(today,start)<rt.min_notice_days)err.push('Pengajuan minimal '+rt.min_notice_days+' hari sebelum tanggal mulai');
      if(rt.max_consecutive_days!=null&&D.diff(start,end)+1>rt.max_consecutive_days)err.push('Maksimal '+rt.max_consecutive_days+' hari berturut-turut');
      if(rt.one_time_per_employment&&active.some(r=>r.request_type_id===rt.id))err.push('Jenis ini hanya dapat diambil satu kali selama masa kerja');
      if(rt.requires_attachment_after_days!=null&&total>rt.requires_attachment_after_days&&!att)err.push('Lampiran (tautan berkas) wajib bila lebih dari '+rt.requires_attachment_after_days+' hari');
      if(rt.deducts_balance){period=this.periodKey(ctx,start);bal=this.balance(ctx,emp,rt.id,period);res=this.reserved(ctx,emp,rt.id,period,excl);avail=bal-res;if(total>avail)err.push('Saldo tidak cukup. Tersedia '+avail+' hari, diminta '+total+' hari')}
    }else if(rt.category==='time_permit'){
      end=start;if(!p.start_time)err.push('Jam yang diminta wajib diisi');
      if(!this.isWorkday(ctx,emp,start))err.push('Tanggal yang dipilih bukan hari kerja Anda');
      if(active.some(r=>r.request_type_id===rt.id&&String(r.start_date).slice(0,10)===start))err.push('Sudah ada pengajuan sejenis pada tanggal ini');
    }else if(rt.category==='correction'){
      end=start;const c=p.correction||{};
      if(start>today)err.push('Koreksi hanya untuk tanggal yang sudah lewat atau hari ini');
      if(!c.in&&!c.out)err.push('Isi jam masuk dan/atau jam pulang yang benar');
      const lim=+this.set(ctx,'correction_days_limit',3);let cnt=0;for(let d=D.add(start,1);d<=today;d=D.add(d,1))if(this.isWorkday(ctx,emp,d))cnt++;
      if(cnt>lim)err.push('Batas pengajuan koreksi adalah '+lim+' hari kerja ke belakang');
      if((ctx.periods||[]).some(x=>x.status==='locked'&&x.year===D.parse(start).y&&x.month===D.parse(start).m))err.push('Periode tanggal tersebut sudah dikunci');
      if((ctx.requests||[]).some(r=>r.employee_id===emp&&r.request_type_id===rt.id&&String(r.start_date).slice(0,10)===start&&['waiting_l1','waiting_l2'].includes(r.status)&&r.id!==excl))err.push('Sudah ada pengajuan koreksi yang menunggu untuk tanggal ini');
    }else if(rt.category==='overtime'){
      end=start;const mx=+this.set(ctx,'overtime_max_hours_per_day',4);
      if(!p.start_time||!p.end_time)err.push('Jam mulai dan jam selesai lembur wajib diisi');
      else if(p.end_time<=p.start_time)err.push('Jam selesai lembur harus setelah jam mulai');
      else if((D.epoch('2000-01-01',p.end_time)-D.epoch('2000-01-01',p.start_time))/36e5>mx)err.push('Lembur maksimal '+mx+' jam per hari');
    }
    return{ok:!err.length,errors:err,total_days:total,balance:bal,reserved:res,available:avail,after:avail-total,deducts:!!rt.deducts_balance,period_key:period};
  },
  /* ringkasan saldo semua jenis cuti berpotong saldo untuk seorang pegawai pada periode sekarang */
  balances(ctx,emp,ds){
    ds=ds||TODAY();const period=this.periodKey(ctx,ds);
    return(ctx.request_types||[]).filter(t=>t.deducts_balance).map(t=>{
      const led=this.ledger(ctx).filter(l=>l.employee_id===emp&&l.request_type_id===t.id&&l.period_key===period);
      const total=sum(led.filter(l=>+l.days>0),l=>+l.days),used=0-sum(led.filter(l=>+l.days<0),l=>+l.days),res=this.reserved(ctx,emp,t.id,period);
      return{type:t,period,total,used,reserved:res,left:total-used-res,balance:total-used};
    });
  },
  /* peringatan penyusunan jadwal shift. plan: {"emp|date": shift_id|null} */
  shiftWarnings(ctx,plan,from,to){
    const ix=this.ix(ctx),w=[],byEmp={};
    for(const k in plan){const[emp,ds]=k.split('|');if(ds<from||ds>to)continue;(byEmp[emp]=byEmp[emp]||[]).push([ds,plan[k]])}
    const minGap=+this.set(ctx,'shift_min_gap_hours',8),maxRun=+this.set(ctx,'shift_max_consecutive_days',6);
    for(const emp in byEmp){
      const e=ix.emp[emp];if(!e)continue;const list=byEmp[emp].filter(x=>x[1]).sort((a,b)=>a[0].localeCompare(b[0]));
      let run=0,prev=null,flagged=false;
      for(const[ds,sid]of list){
        const sh=ix.shift[sid];if(!sh)continue;
        const st=this.mk('shift',ds,sh.start_time,sh.end_time,sh.break_minutes);
        if(prev&&D.diff(prev.ds,ds)===1){run++;const gap=(st.startMs-prev.endMs)/36e5;if(gap<minGap)w.push(e.full_name+': jeda antar shift '+prev.ds.slice(5)+' ke '+ds.slice(5)+' hanya '+nf(Math.max(gap,0))+' jam (minimal '+minGap+')')}
        else run=1;
        if(run>maxRun&&!flagged){w.push(e.full_name+': bekerja lebih dari '+maxRun+' hari berturut-turut (sampai '+ds.slice(5)+')');flagged=true}
        if(run<=maxRun)flagged=false;
        prev={ds,endMs:st.endMs};
        const lv=(ctx.requests||[]).find(r=>r.employee_id===+emp&&r.status==='approved'&&['leave','sick','permit'].includes((ix.rt[r.request_type_id]||{}).category)&&String(r.start_date).slice(0,10)<=ds&&String(r.end_date).slice(0,10)>=ds);
        if(lv)w.push(e.full_name+': dijadwalkan pada '+fmtD(ds)+' saat sedang '+((ix.rt[lv.request_type_id]||{}).name||'cuti').toLowerCase());
      }
    }
    const minStaff=this.set(ctx,'shift_min_staff',{});
    for(const sid in minStaff){const need=+minStaff[sid];if(!need)continue;
      for(const ds of D.range(from,to)){let n=0;for(const k in plan)if(k.endsWith('|'+ds)&&String(plan[k])===String(sid))n++;
        if(n<need){const sh=ix.shift[sid];w.push('Shift '+(sh?sh.name:sid)+' pada '+fmtD(ds)+' hanya '+n+' petugas (minimal '+need+')')}}}
    return w;
  }
};


/* =====================================================================
   6. DATA AWAL (master) dan DATA CONTOH MODE PREVIEW
   ===================================================================== */
const DEFAULT_SETTINGS={
  'school.name':'Fitrah Islamic World Academy','timezone':'Asia/Jakarta','leave_period_mode':'jun25',
  'correction_days_limit':3,'overtime_max_hours_per_day':4,'overtime_round_minutes':15,'min_in_out_gap_minutes':30,
  'late_calc_from':'start','clock.enabled':true,'clock.geofence':false,'school.radius_m':150,
  'approval.single_day_final_l1':false,'allow_cancel_waiting_l2':true,'session_hours':8,
  'shift_min_gap_hours':8,'shift_max_consecutive_days':6,'shift_min_staff':{}
};
const SETTING_DESC={
  'school.name':'Nama lembaga (tampil di sisi kiri dan cetakan)','timezone':'Zona waktu lembaga (Asia/Jakarta, Asia/Makassar, Asia/Jayapura)',
  'leave_period_mode':'Periode saldo cuti: jun25 (26 Juni-25 Juni, hangus tiap 25 Juni), calendar (tahun kalender) atau academic (tahun ajaran)','correction_days_limit':'Batas pengajuan koreksi absen (hari kerja ke belakang)',
  'overtime_max_hours_per_day':'Maksimal jam lembur per hari','overtime_round_minutes':'Pembulatan lembur ke bawah (menit)','min_in_out_gap_minutes':'Jeda minimal masuk-pulang agar dihitung dua kali absen (menit)',
  'late_calc_from':'Hitung menit terlambat dari: start (jam masuk) atau tolerance_end (akhir toleransi)','clock.enabled':'Izinkan absen mandiri dari web/ponsel','clock.geofence':'Batasi absen mandiri pada radius lembaga',
  'school.lat':'Latitude lembaga','school.lng':'Longitude lembaga','school.radius_m':'Radius lokasi lembaga (meter)','approval.single_day_final_l1':'Cuti 1 hari cukup disetujui Kepala Departemen',
  'allow_cancel_waiting_l2':'Pegawai boleh membatalkan pengajuan yang menunggu HRD','session_hours':'Lama sesi login pegawai (jam)','shift_min_gap_hours':'Jeda minimal antar shift (jam)',
  'shift_max_consecutive_days':'Maksimal hari kerja berturut-turut (shift)','shift_min_staff':'Jumlah petugas minimum per shift: {"ID_SHIFT": jumlah}','last_processed_date':'Tanggal terakhir diproses (otomatis)'
};
const SEED_TYPES=[
 {code:'cuti_tahunan',name:'Cuti Tahunan',category:'leave',deducts_balance:true,annual_quota_days:5,min_service_months:12,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:3,max_consecutive_days:5,allow_half_day:true,counts_as_present:false,carry_over_policy:'capped',carry_over_max_days:6,carry_over_valid_months:3,one_time_per_employment:false,is_active:true},
 {code:'cuti_melahirkan',name:'Cuti Melahirkan',category:'leave',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'calendar_days',requires_attachment_after_days:0,min_notice_days:0,max_consecutive_days:90,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true},
 {code:'cuti_menikah',name:'Cuti Menikah',category:'leave',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:7,max_consecutive_days:3,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:true,is_active:true},
 {code:'cuti_duka',name:'Cuti Duka',category:'leave',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:0,max_consecutive_days:3,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true},
 {code:'cuti_haji',name:'Cuti Ibadah Haji / Umrah',category:'leave',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'calendar_days',requires_attachment_after_days:0,min_notice_days:30,max_consecutive_days:40,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:true,is_active:true},
 {code:'cuti_besar',name:'Cuti Besar / Sabatikal',category:'leave',deducts_balance:false,annual_quota_days:0,min_service_months:72,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:30,max_consecutive_days:30,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:false},
 {code:'cuti_tanpa_gaji',name:'Cuti Tanpa Gaji',category:'leave',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:3,max_consecutive_days:null,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true},
 {code:'sakit',name:'Sakit',category:'sick',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:2,min_notice_days:0,max_consecutive_days:null,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true},
 {code:'izin_pribadi',name:'Izin Pribadi / Keluarga',category:'permit',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:0,max_consecutive_days:3,allow_half_day:true,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true},
 {code:'izin_terlambat',name:'Izin Terlambat',category:'time_permit',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:0,max_consecutive_days:null,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true},
 {code:'izin_pulang_cepat',name:'Izin Pulang Cepat',category:'time_permit',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:0,max_consecutive_days:null,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true},
 {code:'dinas_luar',name:'Dinas Luar / Tugas Lembaga',category:'official_duty',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:0,max_consecutive_days:null,allow_half_day:false,counts_as_present:true,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true},
 {code:'koreksi_absen',name:'Koreksi Absen',category:'correction',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:0,max_consecutive_days:null,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true},
 {code:'lembur',name:'Lembur',category:'overtime',deducts_balance:false,annual_quota_days:0,min_service_months:0,count_mode:'working_days',requires_attachment_after_days:null,min_notice_days:0,max_consecutive_days:null,allow_half_day:false,counts_as_present:false,carry_over_policy:'none',carry_over_max_days:0,carry_over_valid_months:0,one_time_per_employment:false,is_active:true}
];
const SEED_DEPTS=['Akademik','Tata Usaha','Keuangan','Sarana dan Prasarana','Keamanan','SDM / HRD'];
const SEED_SCHEDULES=[
  {name:'Guru',tolerance_minutes:10,days:{1:['07:00','15:00',30],2:['07:00','15:00',30],3:['07:00','15:00',30],4:['07:00','15:00',30],5:['07:00','11:30',0]}},
  {name:'Tata Usaha',tolerance_minutes:10,days:{1:['07:30','16:00',60],2:['07:30','16:00',60],3:['07:30','16:00',60],4:['07:30','16:00',60],5:['07:30','16:00',60],6:['07:30','12:00',0]}}
];
const SEED_SHIFTS=[
  {name:'Pagi',start_time:'06:00',end_time:'14:00',break_minutes:30,crosses_midnight:false,tolerance_minutes:10,early_clock_in_minutes:120,late_clock_out_minutes:240,color:'#f59e0b',is_active:true},
  {name:'Siang',start_time:'14:00',end_time:'22:00',break_minutes:30,crosses_midnight:false,tolerance_minutes:10,early_clock_in_minutes:120,late_clock_out_minutes:240,color:'#0ea5e9',is_active:true},
  {name:'Malam',start_time:'22:00',end_time:'06:00',break_minutes:30,crosses_midnight:true,tolerance_minutes:10,early_clock_in_minutes:120,late_clock_out_minutes:240,color:'#6366f1',is_active:true}
];
const FIXED_HOLIDAYS=[['01-01','Tahun Baru Masehi'],['05-01','Hari Buruh Internasional'],['06-01','Hari Lahir Pancasila'],['08-17','Hari Kemerdekaan RI'],['12-25','Hari Raya Natal']];
const EMP_COLS_BASE='id,employee_no,must_change_password,full_name,department_id,position,type,employment_status,roles,join_date,end_date,phone,email,photo_url,title_suffix,nip,division,is_active,failed_login_count,locked_until,last_login_at,created_at,updated_at';
const EMP_COLS=EMP_COLS_BASE.replace('division,','division,'+EMP_X.join(',')+',');
const LEGACY_MSG='Kolom data karyawan baru belum ada di database. Jalankan migrasi-karyawan.sql di Supabase > SQL Editor, lalu muat ulang halaman.';
const empMiss=t=>/(column|schema cache|does not exist)/i.test(t)&&/(gender|birth_place|birth_date|contract_end_date|education_|bpjs_|marital_status)/.test(t);

function mulberry(a){return function(){a|=0;a=a+0x6D2B79F5|0;let t=Math.imul(a^a>>>15,1|a);t=t+Math.imul(t^t>>>7,61|t)^t;return((t^t>>>14)>>>0)/4294967296}}

/* Foto demo: ilustrasi potret 3:4 (SVG) dibuat lokal - tanpa gambar dari internet */
const PH_BG=[['#dbeafe','#bfdbfe'],['#e0e7ff','#c7d2fe'],['#ccfbf1','#99f6e4'],['#fef3c7','#fde68a'],['#fce7f3','#fbcfe8'],['#e2e8f0','#cbd5e1'],['#dcfce7','#bbf7d0']],
 PH_SKIN=[['#f6d5b8','#e3b891'],['#ecbf94','#d6a073'],['#dca67b','#c58a5e'],['#c98d62','#ad7149'],['#a97250','#8f5b3c']],
 PH_HAIR=['#1f2937','#111827','#3b2a20','#4a3426','#6b7280'],
 PH_SHIRT=['#1e3a8a','#1d4ed8','#0f766e','#6d28d9','#9f1239','#334155','#92400e','#047857'],
 PH_HIJAB=['#1e293b','#7c3aed','#be185d','#0f766e','#b45309','#475569','#1d4ed8'];
function demoPhoto(o){
  const bg=PH_BG[o.bg%PH_BG.length],sk=PH_SKIN[o.skin%PH_SKIN.length][0],dk=PH_SKIN[o.skin%PH_SKIN.length][1],hr=PH_HAIR[(o.hair||0)%PH_HAIR.length],sh=PH_SHIRT[o.shirt%PH_SHIRT.length],hj=PH_HIJAB[(o.hc||0)%PH_HIJAB.length],st=o.style||'short',fem=o.g==='f',hij=st==='hijab';
  let s='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 400"><defs><linearGradient id="b" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="'+bg[0]+'"/><stop offset="1" stop-color="'+bg[1]+'"/></linearGradient><linearGradient id="s" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="'+sh+'"/><stop offset="1" stop-color="'+sh+'" stop-opacity=".82"/></linearGradient></defs>'
   +'<rect width="300" height="400" fill="url(#b)"/><circle cx="150" cy="182" r="128" fill="#fff" opacity=".3"/>';
  if(st==='long')s+='<path d="M90 178C80 102 116 74 150 74s70 28 60 104l10 104c-30 16-110 16-140 0z" fill="'+hr+'"/>';
  if(hij)s+='<path d="M78 186C72 98 112 64 150 64s78 34 72 122c-2 40 16 70 32 114H46c16-44 34-74 32-114z" fill="'+hj+'"/>';
  s+='<path d="M6 400c4-86 62-122 144-122s140 36 144 122z" fill="url(#s)"/>';
  s+='<rect x="125" y="224" width="50" height="70" rx="22" fill="'+dk+'"/>';
  if(hij)s+='<path d="M52 400c6-72 42-108 98-112 56 4 92 40 98 112z" fill="'+hj+'"/>';
  else if(fem)s+='<path d="M116 286q34 38 68 0l-8-10q-26 22-52 0z" fill="#fff" opacity=".92"/>';
  else s+='<path d="M118 282l32 44 32-44-14-8-18 18-18-18z" fill="#fff"/>'+(o.tie?'<path d="M142 306h16l7 44-15 14-15-14z" fill="'+o.tie+'"/>':'');
  if(!hij)s+='<ellipse cx="99" cy="184" rx="8" ry="12" fill="'+sk+'"/><ellipse cx="201" cy="184" rx="8" ry="12" fill="'+sk+'"/>';
  s+='<ellipse cx="150" cy="178" rx="'+(hij?45:50)+'" ry="'+(hij?55:60)+'" fill="'+sk+'"/>';
  if(st==='short')s+='<path d="M98 176C90 108 118 80 150 80s60 28 52 96c-4-26-14-40-28-48-26 10-46 10-72 0-14 8-24 22-28 48z" fill="'+hr+'"/>';
  if(st==='long')s+='<path d="M98 176C94 110 122 82 152 82c34 0 56 28 50 94-10-34-30-52-60-54-24 6-40 22-44 54z" fill="'+hr+'"/>';
  if(st==='bun')s+='<circle cx="150" cy="70" r="22" fill="'+hr+'"/><path d="M98 176C90 108 118 78 150 78s60 30 52 98c-4-26-14-40-28-48-26 10-46 10-72 0-14 8-24 22-28 48z" fill="'+hr+'"/>';
  if(hij)s+='<path d="M100 162C106 122 128 106 150 106s44 16 50 56c-14-20-32-28-50-28s-36 8-50 28z" fill="'+hj+'"/>';
  const ey=o.gl?'#1f2937':'#2b2118';
  s+='<g fill="'+ey+'"><ellipse cx="132" cy="182" rx="4" ry="4.6"/><ellipse cx="168" cy="182" rx="4" ry="4.6"/></g>'
   +'<g stroke="'+(st==='short'||st==='bun'||st==='long'?hr:'#3b2a20')+'" stroke-width="3.2" stroke-linecap="round" fill="none"><path d="M120 169q12-6 24-1"/><path d="M156 168q12-5 24 1"/></g>'
   +'<path d="M150 188q-5 14 0 20q4 1 6-1" stroke="'+dk+'" stroke-width="2.6" stroke-linecap="round" fill="none"/>'
   +'<path d="M136 212q14 11 28 0" stroke="#9f4a3a" stroke-width="3.2" stroke-linecap="round" fill="none"/>';
  if(fem)s+='<g fill="#f472b6" opacity=".22"><ellipse cx="120" cy="204" rx="10" ry="6"/><ellipse cx="180" cy="204" rx="10" ry="6"/></g>';
  if(o.gl)s+='<g fill="none" stroke="#1f2937" stroke-width="3"><rect x="114" y="170" width="38" height="26" rx="10"/><rect x="148" y="170" width="38" height="26" rx="10"/><path d="M152 182h-4"/></g>';
  return'data:image/svg+xml;charset=utf-8,'+encodeURIComponent(s+'</svg>')}

function buildPreviewDB(){
  const today=TODAY(),Y=D.parse(today).y,rnd=mulberry(20261011);
  const db={departments:[],employees:[],employee_fingerprints:[],admin_profiles:[],department_substitutes:[],academic_years:[],work_schedules:[],work_schedule_days:[],schedule_assignments:[],shifts:[],shift_patterns:[],shift_assignments:[],holidays:[],import_batches:[],raw_logs:[],request_types:[],requests:[],request_approvals:[],leave_balance_entries:[],attendance_records:[],attendance_periods:[],shift_swaps:[],notifications:[],settings:[],audit_logs:[],pw:{},sessions:{}};
  const now=new Date().toISOString();
  SEED_DEPTS.forEach((n,i)=>db.departments.push({id:i+1,name:n,head_employee_id:null,is_active:true}));
  const slug=s=>String(s).toLowerCase().replace(/[^a-z\s]/g,'').trim().split(/\s+/).slice(0,2).join('.');
  const genNip=(id,o)=>{const yr=1972+(id*3)%26,mo=String(1+(id*5)%12).padStart(2,'0'),dy=String(1+(id*7)%28).padStart(2,'0'),jd=String(o.join||((Y-4)+'-07-15')).slice(0,7).replace('-','');return''+yr+mo+dy+jd+((o.ph&&o.ph.g==='f')?'2':'1')+String(id).padStart(3,'0')};
  const PL=['Bogor','Jakarta','Depok','Bandung','Tangerang','Bekasi','Sukabumi','Cianjur'],UNI=['Universitas Pendidikan Indonesia','Universitas Indonesia','IPB University','UIN Syarif Hidayatullah Jakarta','Universitas Negeri Jakarta','Universitas Pakuan','Universitas Ibn Khaldun'],
    MJ=[[/matematika/i,'Pendidikan Matematika'],[/bahasa indonesia/i,'Pendidikan Bahasa Indonesia'],[/bahasa inggris/i,'Pendidikan Bahasa Inggris'],[/ipa/i,'Pendidikan IPA'],[/sejarah/i,'Pendidikan Sejarah'],[/informatika|\bit\b/i,'Teknik Informatika'],[/kimia/i,'Pendidikan Kimia'],[/fisika/i,'Pendidikan Fisika'],[/pjok/i,'Pendidikan Jasmani'],[/seni/i,'Pendidikan Seni Budaya'],[/keuangan|akuntansi/i,'Akuntansi'],[/hrd|kepegawaian/i,'Manajemen SDM'],[/administrasi|tata usaha|perpustakaan/i,'Administrasi Perkantoran']];
  const demoX=(id,nipv,sfx,pos,o)=>{
    const g=nipv[14]==='2'?'P':'L',lv=/^M\./i.test(sfx||'')?'S2':sfx?'S1':(id%2?'D3':'SMA/SMK'),mj=(MJ.find(x=>x[0].test(pos))||[0,'Manajemen'])[1];
    return{gender:g,birth_place:PL[id%PL.length],birth_date:nipv.slice(0,4)+'-'+nipv.slice(4,6)+'-'+nipv.slice(6,8),
      contract_end_date:o.st==='contract'?(o.end||D.add(today,[25,75,150,240,330][id%5])):null,
      education_level:lv,education_institution:lv==='SMA/SMK'?'SMA Negeri '+(1+id%5)+' '+PL[id%PL.length]:lv==='D3'?'Politeknik Negeri Jakarta':UNI[id%UNI.length],
      education_major:lv==='SMA/SMK'?['IPA','IPS'][id%2]:mj,
      bpjs_kesehatan:('0001'+String(100000000+id*7777777)).slice(0,13),bpjs_ketenagakerjaan:('2'+String(1000000000+id*13579)).slice(0,11),
      marital_status:['menikah','menikah','menikah','belum_menikah','menikah','cerai_mati'][id%6]}};
  const E=(no,name,sfx,dept,pos,type,o)=>{o=o||{};const id=db.employees.length+1,core=name.replace(/^((Prof|Dr|Dra|Drs|Ir|H|Hj)\.\s*)+/i,'');
    const nipv=o.nip||genNip(id,o);db.employees.push({id,_code:no,employee_no:nipv,must_change_password:false,full_name:name,title_suffix:sfx||null,department_id:dept,position:pos,type,employment_status:o.st||'permanent',roles:o.roles||['pegawai'],join_date:o.join||(Y-4)+'-07-15',end_date:o.end||null,phone:o.hp||('0812-'+(3400+id*37)+'-'+(5000+id*113)),email:slug(core)+'@fiwa.sch.id',photo_url:demoPhoto(o.ph),nip:nipv,division:o.div||null,...demoX(id,nipv,sfx,pos,o),is_active:o.act!==false,failed_login_count:0,locked_until:null,last_login_at:null,created_at:now,updated_at:now});db.pw[id]='demo123';return id};
  /* Akademik */
  E('K-001','Dra. Siti Rahmawati','M.Pd',1,'Kepala Departemen Akademik','teacher',{nip:'196809121994032004',div:'Kurikulum',join:(Y-19)+'-03-01',ph:{g:'f',style:'hijab',skin:1,hc:1,bg:0,shirt:2}});
  E('G-002','Budi Santoso','S.Pd',1,'Guru Matematika','teacher',{nip:'198203152008011012',div:'Kurikulum',join:(Y-12)+'-07-14',ph:{g:'m',style:'short',skin:2,hair:0,bg:3,shirt:0,tie:'#1e293b'}});
  E('G-003','Rina Wulandari','S.Pd',1,'Guru Bahasa Indonesia','teacher',{div:'Kurikulum',join:(Y-6)+'-07-15',ph:{g:'f',style:'long',skin:0,hair:2,bg:4,shirt:3}});
  E('G-004','Ahmad Fauzi','M.Pd',1,'Guru IPA','teacher',{nip:'198511042010011008',div:'Kurikulum',join:(Y-9)+'-07-12',ph:{g:'m',style:'short',skin:1,hair:1,bg:2,shirt:1,gl:1}});
  E('G-005','Dewi Lestari','S.Pd',1,'Guru Bahasa Inggris','teacher',{st:'contract',div:'Kurikulum',join:Y+'-01-05',ph:{g:'f',style:'bun',skin:0,hair:3,bg:6,shirt:4}});
  E('G-006','Fitri Handayani','S.Pd',1,'Guru Sejarah','teacher',{st:'honorary',div:'Kurikulum',join:(Y-2)+'-08-01',ph:{g:'f',style:'hijab',skin:2,hc:3,bg:5,shirt:7}});
  E('G-007','Yoga Pratama','S.Kom',1,'Guru Informatika','teacher',{st:'contract',div:'Kurikulum',join:(Y-1)+'-07-15',ph:{g:'m',style:'short',skin:3,hair:1,bg:1,shirt:5,gl:1}});
  E('G-008','Nia Kurniasih','M.Si',1,'Guru Kimia','teacher',{nip:'199002112015032001',div:'Kurikulum',join:(Y-11)+'-07-13',ph:{g:'f',style:'long',skin:1,hair:0,bg:0,shirt:2}});
  E('G-009','Rizky Ramadhan','S.Or',1,'Guru PJOK','teacher',{st:'honorary',div:'Kesiswaan',join:(Y-3)+'-01-10',ph:{g:'m',style:'short',skin:3,hair:2,bg:3,shirt:6}});
  E('G-010','Anita Putri','S.Pd',1,'Guru Seni Budaya','teacher',{st:'contract',div:'Kesiswaan',join:(Y-2)+'-07-15',ph:{g:'f',style:'long',skin:2,hair:1,bg:4,shirt:3}});
  E('G-011','Bambang Hermawan','M.Pd',1,'Guru Fisika','teacher',{nip:'197901262006041005',div:'Kurikulum',join:(Y-17)+'-07-17',ph:{g:'m',style:'short',skin:2,hair:4,bg:5,shirt:0,tie:'#7f1d1d',gl:1}});
  /* Tata Usaha */
  E('T-001','Hj. Nurhayati',null,2,'Kepala Tata Usaha','staff',{nip:'197306151998032002',div:'Administrasi',join:(Y-14)+'-02-02',ph:{g:'f',style:'hijab',skin:1,hc:2,bg:4,shirt:4}});
  E('T-002','Agus Setiawan',null,2,'Staf Administrasi','staff',{div:'Administrasi',join:(Y-5)+'-03-03',ph:{g:'m',style:'short',skin:2,hair:1,bg:6,shirt:1}});
  E('T-003','Maya Sari',null,2,'Staf Administrasi','staff',{div:'Administrasi',join:(Y-3)+'-09-01',ph:{g:'f',style:'bun',skin:1,hair:1,bg:1,shirt:5}});
  E('T-004','Wulan Safitri',null,2,'Staf Perpustakaan','staff',{st:'contract',div:'Perpustakaan',join:(Y-1)+'-08-15',ph:{g:'f',style:'hijab',skin:0,hc:4,bg:2,shirt:7}});
  E('T-005','Dedi Kurniawan','S.Kom',2,'Staf IT & Operator Lembaga','staff',{div:'Teknologi Informasi',join:(Y-4)+'-02-17',ph:{g:'m',style:'short',skin:3,hair:0,bg:1,shirt:2,gl:1}});
  /* Keuangan */
  E('F-001','Hendra Wijaya','S.E',3,'Kepala Keuangan','staff',{nip:'198001222005011007',div:'Akuntansi',join:(Y-13)+'-01-03',ph:{g:'m',style:'short',skin:1,hair:1,bg:0,shirt:0,tie:'#1d4ed8'}});
  E('F-002','Lina Marlina',null,3,'Staf Keuangan','staff',{div:'Kasir',join:(Y-5)+'-05-02',ph:{g:'f',style:'long',skin:2,hair:2,bg:4,shirt:4}});
  E('F-003','Anisa Rahma','S.Ak',3,'Staf Akuntansi','staff',{st:'contract',div:'Akuntansi',join:(Y-1)+'-10-01',ph:{g:'f',style:'hijab',skin:1,hc:6,bg:0,shirt:1}});
  /* Sarana dan Prasarana */
  E('R-001','Tono Sumarno',null,4,'Kepala Sarana dan Prasarana','staff',{nip:'197710182003121002',div:'Pemeliharaan',join:(Y-15)+'-06-16',ph:{g:'m',style:'short',skin:3,hair:4,bg:5,shirt:6}});
  E('R-002','Slamet Riyadi',null,4,'Teknisi Sarana','staff',{st:'honorary',div:'Pemeliharaan',join:(Y-3)+'-04-04',ph:{g:'m',style:'short',skin:4,hair:1,bg:3,shirt:5}});
  E('R-003','Sugeng Prayitno',null,4,'Petugas Kebersihan','staff',{st:'contract',div:'Kebersihan',join:(Y-2)+'-01-08',end:D.add(today,-60),act:false,ph:{g:'m',style:'short',skin:3,hair:4,bg:6,shirt:5}});
  /* Keamanan */
  E('S-001','Joko Susilo',null,5,'Kepala Keamanan','staff',{nip:'198104032007011009',div:'Keamanan Lingkungan',join:(Y-10)+'-02-01',ph:{g:'m',style:'short',skin:2,hair:0,bg:5,shirt:5}});
  E('S-002','Rudi Hartono',null,5,'Petugas Keamanan','staff',{div:'Keamanan Lingkungan',join:(Y-6)+'-03-15',ph:{g:'m',style:'short',skin:3,hair:1,bg:2,shirt:5}});
  E('S-003','Eko Prasetyo',null,5,'Petugas Keamanan','staff',{div:'Keamanan Lingkungan',join:(Y-4)+'-08-01',ph:{g:'m',style:'short',skin:2,hair:2,bg:6,shirt:5,gl:1}});
  /* SDM / HRD */
  E('H-001','Ir. Muhammad Yusuf',null,6,'HRD Manager','staff',{nip:'197508302001121003',div:'Kepegawaian',join:(Y-12)+'-01-15',ph:{g:'m',style:'short',skin:1,hair:4,bg:1,shirt:0,tie:'#0f766e',gl:1}});
  E('H-002','Putri Maharani','S.Psi',6,'HRD Staff','staff',{div:'Kepegawaian',join:(Y-2)+'-06-03',ph:{g:'f',style:'long',skin:0,hair:1,bg:4,shirt:3}});
  E('C-001','Siska Amelia',null,6,'Staf Penjadwalan','staff',{div:'Penjadwalan',join:(Y-5)+'-07-15',ph:{g:'f',style:'bun',skin:1,hair:0,bg:2,shirt:1}});
  E('P-001','Drs. H. Mulyadi',null,2,'Pimpinan Yayasan','staff',{nip:'196212051988031006',div:'Pimpinan',join:(Y-22)+'-01-04',ph:{g:'m',style:'short',skin:1,hair:4,bg:5,shirt:0,tie:'#7f1d1d',gl:1}});
  const eid=no=>db.employees.find(e=>e._code===no).id;
  const head=(d,no)=>{db.departments.find(x=>x.id===d).head_employee_id=eid(no)};
  head(1,'K-001');head(2,'T-001');head(3,'F-001');head(4,'R-001');head(5,'S-001');head(6,'H-001');
  db.admin_profiles.push({id:1,auth_user_id:'demo-admin',full_name:'Super Admin',role:'admin_sistem',employee_id:null,is_active:true,email:'admin@fiwa.sch.id',pw:'admin123'},
    {id:2,auth_user_id:'demo-hrd',full_name:'Ir. Muhammad Yusuf',role:'manajer_hrd',employee_id:eid('H-001'),is_active:true,email:'hrd@fiwa.sch.id',pw:'hrd123'},
    {id:3,auth_user_id:'demo-staf',full_name:'HRD Staff',role:'staf_hrd',employee_id:null,is_active:true,email:'staf@fiwa.sch.id',pw:'staf123'});
  db.academic_years.push({id:1,name:(today>=Y+'-07-01'?Y+'/'+(Y+1):(Y-1)+'/'+Y),start_date:(today>=Y+'-07-01'?Y:Y-1)+'-07-01',end_date:(today>=Y+'-07-01'?Y+1:Y)+'-06-30',is_current:true});
  db.settings=Object.keys(DEFAULT_SETTINGS).map(k=>({key:k,value:k==='school.name'?'Fitrah Islamic World Academy':DEFAULT_SETTINGS[k],description:SETTING_DESC[k]||null}));
  SEED_TYPES.forEach((t,i)=>db.request_types.push({id:i+1,...t}));
  SEED_SCHEDULES.forEach((s,i)=>{const id=i+1;db.work_schedules.push({id,name:s.name,tolerance_minutes:s.tolerance_minutes,early_clock_in_minutes:120,late_clock_out_minutes:240,is_active:true});
    for(let w=1;w<=7;w++){const d=s.days[w];db.work_schedule_days.push({id:db.work_schedule_days.length+1,work_schedule_id:id,weekday:w,is_workday:!!d,start_time:d?d[0]:null,end_time:d?d[1]:null,break_minutes:d?d[2]:0})}});
  SEED_SHIFTS.forEach((s,i)=>db.shifts.push({id:i+1,...s}));
  db.shift_patterns.push({id:1,name:'4 Pagi - 4 Siang - 4 Malam - Libur',cycle:['Pagi','Pagi','Pagi','Pagi',null,'Siang','Siang','Siang','Siang',null,'Malam','Malam','Malam','Malam',null]});
  FIXED_HOLIDAYS.forEach(([md,n])=>[Y-1,Y,Y+1].forEach(y=>db.holidays.push({id:db.holidays.length+1,holiday_date:y+'-'+md,name:n,type:'national',applies_to:'all'})));
  db.holidays.push({id:db.holidays.length+1,holiday_date:D.add(today,9),name:'Libur Peringatan Lembaga (contoh)',type:'school',applies_to:'teacher'});
  const secIds=[eid('S-001'),eid('S-002'),eid('S-003')];
  db.employees.forEach(e=>{if(!secIds.includes(e.id)){db.schedule_assignments.push({id:db.schedule_assignments.length+1,employee_id:e.id,work_schedule_id:e.type==='teacher'?1:2,effective_from:e.join_date,effective_to:null})}});
  // jadwal shift keamanan (terbit) -45 s.d. +40 hari
  const cyc=[1,1,1,1,null,2,2,2,2,null,3,3,3,3,null];
  secIds.forEach((id,k)=>{for(let i=-45;i<=40;i++){const ds=D.add(today,i);db.shift_assignments.push({id:db.shift_assignments.length+1,employee_id:id,work_date:ds,shift_id:cyc[(i+45+k*3)%cyc.length],status:'published',published_at:now})}});
  // saldo cuti
  const tt=db.request_types.find(t=>t.code==='cuti_tahunan');
  db.employees.filter(e=>Engine.svcMonths(e.join_date,TODAY())>=12).forEach(e=>db.leave_balance_entries.push({id:db.leave_balance_entries.length+1,employee_id:e.id,request_type_id:tt.id,period_key:Engine.leaveKey(TODAY()),entry_type:'accrual',days:5,request_id:null,note:'Jatah cuti tahunan '+Engine.leaveKey(TODAY()),created_by:'sistem',valid_until:null,created_at:now}));
  // pengajuan contoh
  const R=(emp,code,s,e,reason,status,extra)=>{const rt=db.request_types.find(t=>t.code===code);const id=db.requests.length+1;
    const r={id,employee_id:eid(emp),request_type_id:rt.id,start_date:s,end_date:e,start_time:null,end_time:null,day_part:'full',total_days:0,reason,attachment_url:null,status,current_stage:status==='waiting_l1'?1:2,correction_payload:null,overtime_payload:null,cancel_requested:false,cancel_reason:null,applied:false,submitted_at:new Date(Date.now()-86400000*2).toISOString(),cancelled_at:null,...extra};db.requests.push(r);return r};
  const ctx0={...db,settings:Object.fromEntries(db.settings.map(s=>[s.key,s.value])),ledger:db.leave_balance_entries,periods:[]};
  const days=(emp,code,s,e)=>Engine.countDays(ctx0,eid(emp),db.request_types.find(t=>t.code===code),s,e,'full');
  const apv=(r,stage,name)=>db.request_approvals.push({id:db.request_approvals.length+1,request_id:r.id,stage,approver_name:name,approver_employee_id:null,decision:'approved',note:null,decided_at:now});
  let nxt=D.add(today,3);while(D.iso(nxt)>5)nxt=D.add(nxt,1);let nxt2=D.add(nxt,1);while(D.iso(nxt2)>5)nxt2=D.add(nxt2,1);
  const r1=R('G-003','cuti_tahunan',nxt,nxt2,'Acara keluarga di luar kota','waiting_l1');r1.total_days=days('G-003','cuti_tahunan',nxt,nxt2);
  let sk=D.add(today,-3);while(D.iso(sk)>5)sk=D.add(sk,-1);const sk2=D.add(sk,1);
  const r2=R('G-004','sakit',sk,D.iso(sk2)>5?sk:sk2,'Demam dan flu','approved');r2.end_date=D.iso(sk2)>5?sk:sk2;r2.total_days=days('G-004','sakit',r2.start_date,r2.end_date);apv(r2,1,'Dra. Siti Rahmawati (Kepala Dept.)');apv(r2,2,'Ir. Muhammad Yusuf (manajer_hrd)');
  let pr=D.add(today,2);while(D.iso(pr)>5)pr=D.add(pr,1);
  const r3=R('T-002','izin_pribadi',pr,pr,'Mengurus dokumen keluarga','waiting_l2');r3.total_days=1;apv(r3,1,'Hj. Nurhayati (Kepala Dept.)');
  let cf=D.add(today,10);while(D.iso(cf)>3)cf=D.add(cf,1);const cf2=D.add(cf,2);
  const r4=R('F-002','cuti_tahunan',cf,cf2,'Liburan keluarga','approved');r4.total_days=days('F-002','cuti_tahunan',cf,cf2);apv(r4,1,'Hendra Wijaya, S.E (Kepala Dept.)');apv(r4,2,'Ir. Muhammad Yusuf (manajer_hrd)');
  db.leave_balance_entries.push({id:db.leave_balance_entries.length+1,employee_id:r4.employee_id,request_type_id:tt.id,period_key:String(Y),entry_type:'usage',days:-r4.total_days,request_id:r4.id,note:'Cuti Tahunan disetujui',created_by:'sistem',valid_until:null,created_at:now});
  let cd=D.add(today,-1);while(D.iso(cd)>5)cd=D.add(cd,-1);
  const r5=R('G-005','koreksi_absen',cd,cd,'Lupa absen pulang karena rapat','waiting_l1');r5.correction_payload={in:'07:00',out:'15:05'};
  const r6=R('S-002','lembur',today,today,'Menjaga kegiatan malam lembaga','waiting_l1');r6.start_time='18:00';r6.end_time='20:00';r6.overtime_payload={note:'Kegiatan malam'};
  db.shift_swaps.push({id:1,requester_employee_id:eid('S-002'),partner_employee_id:eid('S-003'),requester_date:D.add(today,6),partner_date:D.add(today,6),reason:'Ada keperluan keluarga',status:'waiting_coordinator',decided_by:null,decided_at:null,created_at:now});
  // log fingerprint contoh 30 hari
  const ctx1={...db,settings:Object.fromEntries(db.settings.map(s=>[s.key,s.value])),ledger:db.leave_balance_entries,periods:[],raw_logs:[],attendance_records:[]};
  const nowMs=NOW();
  for(const e of db.employees){
    db.employee_fingerprints.push({id:db.employee_fingerprints.length+1,employee_id:e.id,fingerprint_id:e.employee_no});
    for(let i=-35;i<=0;i++){
      const ds=D.add(today,i);Engine.touch(ctx1);const s=Engine.sched(ctx1,e.id,ds);if(s.kind!=='work')continue;
      const cover=db.requests.some(r=>r.employee_id===e.id&&r.status==='approved'&&r.start_date<=ds&&r.end_date>=ds);if(cover)continue;
      const x=rnd();if(x<0.04)continue;
      let inMs=s.startMs+(rnd()<0.12?(10+Math.floor(rnd()*35))*6e4:(-Math.floor(rnd()*18)+3)*6e4);
      let outMs=s.endMs+Math.floor(rnd()*30)*6e4-(rnd()<0.05?70*6e4:0);
      if(ds===today){if(inMs>nowMs)continue;outMs=null}
      const push=(ms)=>db.raw_logs.push({id:db.raw_logs.length+1,fingerprint_id:e.employee_no,employee_id:e.id,logged_at:new Date(ms).toISOString(),source:'fingerprint_import',match_status:'matched',import_batch_id:null,latitude:null,longitude:null});
      if(rnd()<0.02&&ds!==today){/* lupa absen pulang */push(inMs)}else{push(inMs);if(outMs!=null)push(outMs)}
    }
  }
  const ctx2={...db,settings:ctx1.settings,ledger:db.leave_balance_entries,periods:[]};
  const recs=Engine.recalcRange(ctx2,db.employees.map(e=>e.id),D.add(today,-35),D.add(today,-1),nowMs);
  recs.forEach((r,i)=>db.attendance_records.push({id:i+1,...r,updated_at:now}));
  db.settings.push({key:'last_processed_date',value:D.add(today,-1),description:SETTING_DESC.last_processed_date});
  const nt=(emp,title,body,link)=>db.notifications.push({id:db.notifications.length+1,employee_id:emp,audience:emp?'employee':'hrd',title,body,link,is_read:false,created_at:now});
  nt(null,'Pengajuan menunggu persetujuan akhir','Izin Pribadi T-002 menunggu persetujuan HRD.','persetujuan');
  nt(eid('K-001'),'Pengajuan baru menunggu persetujuan','Rina Wulandari mengajukan Cuti Tahunan.','persetujuan');
  nt(eid('G-003'),'Selamat datang','Ini data contoh mode Preview. Silakan coba semua menu.','dashboard');
  return db;
}


/* =====================================================================
   7. LAYANAN DATA: Preview (memori) dan Supabase
   ===================================================================== */
const PV={
  db:null,adm:null,
  init(){if(!this.db)this.db=buildPreviewDB()},
  settingsObj(){return Object.fromEntries(this.db.settings.map(s=>[s.key,s.value]))},
  ctxAll(){return{...this.db,settings:this.settingsObj(),ledger:this.db.leave_balance_entries,periods:this.db.attendance_periods}},
  nid(t){const a=this.db[t];return a.length?Math.max(...a.map(x=>x.id||0))+1:1},
  emp(id){return this.db.employees.find(e=>e.id===id)},
  now(){return new Date().toISOString()},
  notify(emp,title,body,link){this.db.notifications.push({id:this.nid('notifications'),employee_id:emp,audience:'employee',title,body,link,is_read:false,created_at:this.now()})},
  notifyHrd(title,body,link){this.db.notifications.push({id:this.nid('notifications'),employee_id:null,audience:'hrd',title,body,link,is_read:false,created_at:this.now()})},
  audit(actor,action,entity,id,o,n){this.db.audit_logs.push({id:this.nid('audit_logs'),occurred_at:this.now(),actor,action,entity_type:entity,entity_id:String(id||''),old_value:o||null,new_value:n||null})},
  headDepts(emp){const t=TODAY(),d=this.db;return[...d.departments.filter(x=>x.head_employee_id===emp).map(x=>x.id),...d.department_substitutes.filter(s=>s.substitute_employee_id===emp&&s.start_date<=t&&s.end_date>=t).map(s=>s.department_id)]},
  canL1(a,r){const e=this.emp(r);return a!==r&&this.headDepts(a).includes(e.department_id)},
  sess(tok){const id=this.db.sessions[tok];if(!id)throw new Error('SESI_TIDAK_VALID: Sesi berakhir, silakan login kembali');return id},
  me(id){const e=this.emp(id);return{id:e.id,employee_no:e.employee_no,full_name:e.full_name,department_id:e.department_id,position:e.position,type:e.type,employment_status:e.employment_status,roles:e.roles,join_date:e.join_date,phone:e.phone,email:e.email,photo_url:e.photo_url,title_suffix:e.title_suffix,nip:e.nip,division:e.division,must_change_password:e.must_change_password,head_dept_ids:this.headDepts(id)}},
  adminName(){return this.adm?this.adm.full_name+' ('+this.adm.role+')':'sistem'},
  isHr(){return this.adm&&['manajer_hrd','admin_sistem'].includes(this.adm.role)},
  periodLocked(ds){const p=D.parse(ds);return this.db.attendance_periods.some(x=>x.year===p.y&&x.month===p.m&&x.status==='locked')},
  finalize(id,decision,note,actor,actorEmp,stage){
    const d=this.db,r=d.requests.find(x=>x.id===id);if(!r)throw new Error('Pengajuan tidak ditemukan');
    if(!['waiting_l1','waiting_l2'].includes(r.status))throw new Error('Status pengajuan sudah berubah, muat ulang halaman');
    const rt=d.request_types.find(t=>t.id===r.request_type_id),ctx=this.ctxAll();
    if(decision==='approved'){
      if(rt.deducts_balance&&r.total_days>0){const per=Engine.periodKey(ctx,r.start_date);if(Engine.balance(ctx,r.employee_id,rt.id,per)<r.total_days)throw new Error('Saldo cuti pegawai tidak cukup saat persetujuan (periode '+per+')');
        d.leave_balance_entries.push({id:this.nid('leave_balance_entries'),employee_id:r.employee_id,request_type_id:rt.id,period_key:per,entry_type:'usage',days:-r.total_days,request_id:r.id,note:rt.name+' disetujui',created_by:actor,valid_until:null,created_at:this.now()})}
      r.status='approved';r.current_stage=2;r.applied=false;
      d.request_approvals.push({id:this.nid('request_approvals'),request_id:id,stage,approver_name:actor,approver_employee_id:actorEmp,decision:'approved',note:(note||'').trim()||null,decided_at:this.now()});
      this.notify(r.employee_id,'Pengajuan disetujui',rt.name+' tanggal '+fmtD(r.start_date)+' telah disetujui.','pengajuan');
    }else{
      if(!String(note||'').trim())throw new Error('Catatan wajib diisi saat menolak');
      r.status='rejected';r.current_stage=2;
      d.request_approvals.push({id:this.nid('request_approvals'),request_id:id,stage,approver_name:actor,approver_employee_id:actorEmp,decision:'rejected',note,decided_at:this.now()});
      this.notify(r.employee_id,'Pengajuan ditolak',rt.name+' tanggal '+fmtD(r.start_date)+' ditolak: '+note,'pengajuan');
    }
    this.audit(actor,'request.'+decision,'requests',id,null,{stage,note});
    return{ok:true,id,status:r.status,employee_id:r.employee_id,start_date:r.start_date,end_date:r.end_date};
  },
  swapApply(id,decision,actor){
    const d=this.db,s=d.shift_swaps.find(x=>x.id===id);if(!s)throw new Error('Permintaan tukar shift tidak ditemukan');
    if(s.status!=='waiting_coordinator')throw new Error('Status permintaan sudah berubah');
    if(decision!=='approved'){s.status='rejected';s.decided_by=actor;s.decided_at=this.now();this.notify(s.requester_employee_id,'Tukar shift ditolak','Permintaan tukar shift ditolak.','jadwal');this.notify(s.partner_employee_id,'Tukar shift ditolak','Permintaan tukar shift ditolak HRD.','jadwal');return{ok:true,status:'rejected'}}
    if(this.periodLocked(s.requester_date)||this.periodLocked(s.partner_date))throw new Error('Periode tanggal tersebut sudah dikunci');
    const A=(e,ds)=>d.shift_assignments.find(a=>a.employee_id===e&&a.work_date===ds);
    const a1=A(s.requester_employee_id,s.requester_date),a2=A(s.partner_employee_id,s.partner_date);
    if(!a1||!a2||!a1.shift_id||!a2.shift_id||a1.status!=='published'||a2.status!=='published')throw new Error('Salah satu jadwal sudah berubah atau bukan shift kerja');
    const s1=a1.shift_id,s2=a2.shift_id,up=(e,ds,sh)=>{const a=A(e,ds);if(a){a.shift_id=sh;a.status='published'}else d.shift_assignments.push({id:this.nid('shift_assignments'),employee_id:e,work_date:ds,shift_id:sh,status:'published',published_at:this.now()})};
    if(s.requester_date===s.partner_date){a1.shift_id=s2;a2.shift_id=s1}
    else{
      const b=A(s.partner_employee_id,s.requester_date),c=A(s.requester_employee_id,s.partner_date);
      if(b&&b.shift_id)throw new Error('Pegawai pengganti sudah punya shift pada tanggal pengaju');
      if(c&&c.shift_id)throw new Error('Pengaju sudah punya shift pada tanggal pengganti');
      a1.shift_id=null;a2.shift_id=null;up(s.partner_employee_id,s.requester_date,s1);up(s.requester_employee_id,s.partner_date,s2);
    }
    s.status='approved';s.decided_by=actor;s.decided_at=this.now();
    this.notify(s.requester_employee_id,'Tukar shift disetujui','Jadwal Anda telah diperbarui.','jadwal');this.notify(s.partner_employee_id,'Tukar shift disetujui','Jadwal Anda telah diperbarui.','jadwal');
    this.audit(actor,'swap.approved','shift_swaps',id,null,s);
    return{ok:true,status:'approved',dates:[s.requester_date,s.partner_date],employees:[s.requester_employee_id,s.partner_employee_id]};
  },
  async rpc(name,p){
    this.init();await sleep(40);const d=this.db,RP=this.RP;if(!RP[name])throw new Error('Fungsi tidak dikenal: '+name);
    return clone(await RP[name].call(this,p||{}));
  },
  RP:{
    u_login(p){
      const mail=String(p.p_email||'').trim().toLowerCase(),e=this.db.employees.find(x=>String(x.email||'').toLowerCase()===mail);
      if(!e||!e.is_active||this.db.pw[e.id]!==p.p_password)return{ok:false,message:'Email atau kata sandi salah'};
      const tok='pv-'+uid()+uid();this.db.sessions[tok]=e.id;e.last_login_at=this.now();
      return{ok:true,token:tok,me:this.me(e.id)};
    },
    u_logout(p){delete this.db.sessions[p.p_token]},
    u_change_password(p){const id=this.sess(p.p_token),e=this.emp(id);
      if(this.db.pw[e.id]!==p.p_old)return{ok:false,message:'Kata sandi lama salah'};
      if(String(p.p_new||'').length<6)return{ok:false,message:'Kata sandi baru minimal 6 karakter'};
      if(p.p_new.toLowerCase()===String(e.email||'').toLowerCase()||p.p_new.toLowerCase()===e.employee_no.toLowerCase())return{ok:false,message:'Kata sandi baru tidak boleh sama dengan email atau NIP'};
      this.db.pw[e.id]=p.p_new;e.must_change_password=false;return{ok:true}},
    u_update_profile(p){const e=this.emp(this.sess(p.p_token));e.phone=p.p_phone||null;e.email=p.p_email||null;return{ok:true}},
    u_mark_read(p){const id=this.sess(p.p_token);this.db.notifications.filter(n=>n.employee_id===id&&(p.p_id==null||n.id===p.p_id)).forEach(n=>n.is_read=true)},
    u_context(p){
      const id=this.sess(p.p_token),me=this.emp(id),d=this.db,c=this.ctxAll(),heads=this.headDepts(id);
      const inD=e=>heads.includes(e.department_id);
      const att=d.employees.filter(e=>e.id===id||inD(e)).map(e=>e.id),list=d.employees.filter(e=>e.id===id||inD(e)).map(e=>e.id),req=d.employees.filter(e=>e.id===id||inD(e)).map(e=>e.id);
      const from=p.p_from,to=p.p_to,f1=D.add(from,-1),t1=D.add(to,1),fMs=D.epoch(f1,'00:00'),tMs=D.epoch(D.add(to,2),'00:00');
      const reqs=d.requests.filter(r=>req.includes(r.employee_id)&&((r.start_date<=to&&r.end_date>=from)||['waiting_l1','waiting_l2'].includes(r.status)||r.cancel_requested));
      return{me:this.me(id),server_now:this.now(),today:TODAY(),settings:c.settings,
        departments:d.departments.map(x=>({id:x.id,name:x.name,head_employee_id:x.head_employee_id,is_active:x.is_active})),
        employees:d.employees.filter(e=>list.includes(e.id)).map(e=>({id:e.id,employee_no:e.employee_no,full_name:e.full_name,department_id:e.department_id,position:e.position,type:e.type,employment_status:e.employment_status,roles:e.roles,join_date:e.join_date,end_date:e.end_date,is_active:e.is_active})),
        work_schedules:d.work_schedules,work_schedule_days:d.work_schedule_days,schedule_assignments:d.schedule_assignments.filter(a=>list.includes(a.employee_id)),shifts:d.shifts,
        shift_assignments:d.shift_assignments.filter(a=>list.includes(a.employee_id)&&a.work_date>=f1&&a.work_date<=t1&&a.status==='published'),
        holidays:d.holidays.filter(h=>h.holiday_date>=f1&&h.holiday_date<=t1),
        raw_logs:d.raw_logs.filter(l=>att.includes(l.employee_id)&&Date.parse(l.logged_at)>=fMs&&Date.parse(l.logged_at)<tMs).map(l=>({id:l.id,employee_id:l.employee_id,logged_at:l.logged_at,source:l.source})),
        attendance_records:d.attendance_records.filter(a=>att.includes(a.employee_id)&&a.work_date>=from&&a.work_date<=to),
        request_types:d.request_types.filter(t=>t.is_active),
        requests:reqs.map(r=>({...r,reason:req.includes(r.employee_id)?r.reason:null,attachment_url:req.includes(r.employee_id)?r.attachment_url:null})).sort((a,b)=>b.id-a.id),
        approvals:d.request_approvals.filter(a=>reqs.some(r=>r.id===a.request_id&&req.includes(r.employee_id))),
        ledger:d.leave_balance_entries.filter(l=>req.includes(l.employee_id)),
        swaps:d.shift_swaps.filter(s=>s.requester_employee_id===id||s.partner_employee_id===id).sort((a,b)=>b.id-a.id),
        periods:d.attendance_periods.map(x=>({year:x.year,month:x.month,status:x.status})),academic_years:d.academic_years,
        notifications:d.notifications.filter(n=>n.employee_id===id).sort((a,b)=>b.id-a.id).slice(0,50)};
    },
    u_clock(p){
      const id=this.sess(p.p_token),st=this.settingsObj(),d=this.db;
      if(st['clock.enabled']===false)return{ok:false,message:'Absen mandiri tidak diaktifkan oleh lembaga. Gunakan mesin fingerprint.'};
      if(d.raw_logs.some(l=>l.employee_id===id&&Date.now()-Date.parse(l.logged_at)<60000))return{ok:false,message:'Absen baru saja tercatat. Tunggu satu menit.'};
      if(st['clock.geofence']){
        const duty=d.requests.some(r=>r.employee_id===id&&r.status==='approved'&&(d.request_types.find(t=>t.id===r.request_type_id)||{}).category==='official_duty'&&r.start_date<=TODAY()&&r.end_date>=TODAY());
        if(!duty){if(p.p_lat==null||p.p_lng==null)return{ok:false,message:'Lokasi tidak terbaca. Izinkan akses lokasi pada browser Anda.'};
          if(st['school.lat']!=null&&st['school.lng']!=null){const R=6371000,rad=x=>x*Math.PI/180,dist=2*R*Math.asin(Math.sqrt(Math.sin(rad(p.p_lat-st['school.lat'])/2)**2+Math.cos(rad(st['school.lat']))*Math.cos(rad(p.p_lat))*Math.sin(rad(p.p_lng-st['school.lng'])/2)**2));
            if(dist>(st['school.radius_m']||150))return{ok:false,message:'Anda berada di luar area lembaga ('+Math.round(dist)+' m dari lokasi, batas '+(st['school.radius_m']||150)+' m).'}}}
      }
      const l={id:this.nid('raw_logs'),fingerprint_id:null,employee_id:id,logged_at:this.now(),source:'web',match_status:'matched',import_batch_id:null,latitude:p.p_lat||null,longitude:p.p_lng||null};d.raw_logs.push(l);
      return{ok:true,id:l.id,logged_at:l.logged_at};
    },
    u_preview_request(p){return Engine.validateRequest(this.ctxAll(),this.sess(p.p_token),p.p_payload)},
    u_submit_request(p){
      const id=this.sess(p.p_token),d=this.db,pl=p.p_payload,v=Engine.validateRequest(this.ctxAll(),id,pl);if(!v.ok)return v;
      const rt=d.request_types.find(t=>t.id===+pl.type_id),e=this.emp(id),start=pl.start_date;
      const end=['time_permit','correction','overtime'].includes(rt.category)?start:(pl.end_date||start);
      const dept=d.departments.find(x=>x.id===e.department_id),toL2=!dept||!dept.head_employee_id||dept.head_employee_id===id;
      const r={id:this.nid('requests'),employee_id:id,request_type_id:rt.id,start_date:start,end_date:end,start_time:pl.start_time||null,end_time:pl.end_time||null,day_part:pl.day_part||'full',total_days:v.total_days,reason:String(pl.reason).trim(),attachment_url:(pl.attachment_url||'').trim()||null,status:toL2?'waiting_l2':'waiting_l1',current_stage:toL2?2:1,correction_payload:rt.category==='correction'?pl.correction:null,overtime_payload:rt.category==='overtime'?pl.overtime||{}:null,cancel_requested:false,cancel_reason:null,applied:false,submitted_at:this.now(),cancelled_at:null};
      d.requests.push(r);
      if(!toL2){const ids=new Set([dept.head_employee_id,...d.department_substitutes.filter(s=>s.department_id===dept.id&&s.start_date<=TODAY()&&s.end_date>=TODAY()).map(s=>s.substitute_employee_id)]);ids.forEach(x=>this.notify(x,'Pengajuan baru menunggu persetujuan',e.full_name+' mengajukan '+rt.name+' tanggal '+fmtD(start)+'.','persetujuan'))}
      else this.notifyHrd('Pengajuan baru menunggu persetujuan akhir',e.full_name+' mengajukan '+rt.name+' tanggal '+fmtD(start)+'.','persetujuan');
      this.audit(e.full_name,'request.submit','requests',r.id,null,pl);
      return{ok:true,id:r.id,status:r.status,total_days:v.total_days};
    },
    u_cancel_request(p){
      const id=this.sess(p.p_token),r=this.db.requests.find(x=>x.id===p.p_id&&x.employee_id===id);if(!r)throw new Error('Pengajuan tidak ditemukan');
      const st=this.settingsObj();
      if(r.status==='waiting_l1'||(r.status==='waiting_l2'&&st.allow_cancel_waiting_l2!==false)){r.status='cancelled';r.cancelled_at=this.now();r.cancel_reason=p.p_reason||null;return{ok:true,status:'cancelled'}}
      if(r.status==='approved'){if(r.end_date<TODAY())throw new Error('Pengajuan yang sudah lewat tidak dapat dibatalkan');if(!String(p.p_reason||'').trim())throw new Error('Alasan pembatalan wajib diisi');
        r.cancel_requested=true;r.cancel_reason=p.p_reason;this.notifyHrd('Permintaan pembatalan pengajuan',this.emp(id).full_name+' meminta pembatalan pengajuan tanggal '+fmtD(r.start_date)+'.','persetujuan');return{ok:true,status:'cancel_requested'}}
      throw new Error('Pengajuan ini tidak dapat dibatalkan');
    },
    u_decide_l1(p){
      const id=this.sess(p.p_token),r=this.db.requests.find(x=>x.id===p.p_id);if(!r)throw new Error('Pengajuan tidak ditemukan');
      if(r.status!=='waiting_l1')throw new Error('Pengajuan ini sudah diproses');
      if(!this.canL1(id,r.employee_id))throw new Error('Anda bukan penyetuju tahap 1 untuk pegawai ini');
      const me=this.emp(id),rt=this.db.request_types.find(t=>t.id===r.request_type_id),actor=me.full_name+' (Kepala Dept.)';
      if(p.p_decision==='rejected')return this.finalize(r.id,'rejected',p.p_note,actor,id,1);
      if(this.settingsObj()['approval.single_day_final_l1']&&rt.category==='leave'&&r.total_days<=1)return this.finalize(r.id,'approved',p.p_note,actor,id,1);
      r.status='waiting_l2';r.current_stage=2;
      this.db.request_approvals.push({id:this.nid('request_approvals'),request_id:r.id,stage:1,approver_name:actor,approver_employee_id:id,decision:'approved',note:(p.p_note||'').trim()||null,decided_at:this.now()});
      this.notify(r.employee_id,'Pengajuan disetujui tahap 1',rt.name+' tanggal '+fmtD(r.start_date)+' menunggu persetujuan HRD.','pengajuan');
      this.notifyHrd('Pengajuan menunggu persetujuan akhir',rt.name+' tanggal '+fmtD(r.start_date)+' telah disetujui Kepala Departemen.','persetujuan');
      return{ok:true,id:r.id,status:'waiting_l2'};
    },
    u_save_shifts(p){
      const id=this.sess(p.p_token),me=this.emp(id),d=this.db;throw new Error('Jadwal shift hanya dapat diubah HRD melalui panel admin');
      const ids=new Set();let n=0;
      for(const r of p.p_rows){if(this.periodLocked(r.work_date))throw new Error('Periode '+r.work_date.slice(5,7)+'-'+r.work_date.slice(0,4)+' sudah dikunci');
        const i=d.shift_assignments.findIndex(a=>a.employee_id===r.employee_id&&a.work_date===r.work_date);
        if(r.remove){if(i>=0)d.shift_assignments.splice(i,1)}
        else{const row={employee_id:r.employee_id,work_date:r.work_date,shift_id:r.shift_id||null,status:p.p_publish?'published':'draft',published_at:p.p_publish?this.now():null};if(i>=0)Object.assign(d.shift_assignments[i],row);else d.shift_assignments.push({id:this.nid('shift_assignments'),...row})}
        n++;ids.add(r.employee_id)}
      if(p.p_publish)ids.forEach(x=>this.notify(x,'Jadwal shift diperbarui','Jadwal shift Anda telah diterbitkan atau diubah. Silakan cek menu Jadwal Saya.','jadwal'));
      return{ok:true,saved:n};
    },
    u_publish_shifts(p){
      const id=this.sess(p.p_token),me=this.emp(id);throw new Error('Jadwal shift hanya dapat diterbitkan HRD melalui panel admin');
      const rows=this.db.shift_assignments.filter(a=>a.work_date>=p.p_from&&a.work_date<=p.p_to&&a.status==='draft');
      new Set(rows.map(r=>r.employee_id)).forEach(x=>this.notify(x,'Jadwal shift diterbitkan','Jadwal shift '+fmtD(p.p_from)+' s.d. '+fmtD(p.p_to)+' sudah terbit.','jadwal'));
      rows.forEach(r=>{r.status='published';r.published_at=this.now()});return{ok:true,published:rows.length};
    },
    u_swap_options(p){
      const id=this.sess(p.p_token),d=this.db,me=this.emp(id),t=TODAY(),e=D.add(t,60);
      const sh=emp=>d.shift_assignments.filter(a=>a.employee_id===emp&&a.status==='published'&&a.shift_id&&a.work_date>=t&&a.work_date<=e).sort((a,b)=>a.work_date.localeCompare(b.work_date)).map(a=>({work_date:a.work_date,shift_id:a.shift_id}));
      const partners=d.employees.filter(x=>x.is_active&&x.id!==id&&(!me.department_id||x.department_id===me.department_id)).map(x=>({id:x.id,full_name:x.full_name,shifts:sh(x.id)})).filter(x=>x.shifts.length);
      const ids=new Set();d.shift_swaps.forEach(s=>{if(s.partner_employee_id===id)ids.add(s.requester_employee_id);if(s.requester_employee_id===id)ids.add(s.partner_employee_id)});
      return{mine:sh(id),partners,people:[...ids].map(i=>({id:i,full_name:this.emp(i).full_name}))};
    },
    u_swap_create(p){
      const id=this.sess(p.p_token),d=this.db,me=this.emp(id),t=TODAY();
      if(p.p_partner===id)throw new Error('Pegawai pengganti tidak boleh diri sendiri');
      if(p.p_my_date<t||p.p_partner_date<t)throw new Error('Tanggal tukar harus hari ini atau setelahnya');
      const has=(e,ds)=>d.shift_assignments.some(a=>a.employee_id===e&&a.work_date===ds&&a.status==='published'&&a.shift_id);
      if(!has(id,p.p_my_date))throw new Error('Anda tidak punya shift terbit pada tanggal '+fmtD(p.p_my_date));
      if(!has(p.p_partner,p.p_partner_date))throw new Error('Pegawai pengganti tidak punya shift terbit pada tanggal '+fmtD(p.p_partner_date));
      if(d.shift_swaps.some(s=>['waiting_partner','waiting_coordinator'].includes(s.status)&&((s.requester_employee_id===id&&s.requester_date===p.p_my_date)||(s.partner_employee_id===p.p_partner&&s.partner_date===p.p_partner_date))))throw new Error('Sudah ada permintaan tukar yang masih berjalan untuk shift tersebut');
      const s={id:this.nid('shift_swaps'),requester_employee_id:id,partner_employee_id:p.p_partner,requester_date:p.p_my_date,partner_date:p.p_partner_date,reason:(p.p_reason||'').trim()||null,status:'waiting_partner',decided_by:null,decided_at:null,created_at:this.now()};
      d.shift_swaps.push(s);this.notify(p.p_partner,'Permintaan tukar shift',me.full_name+' mengajak tukar shift tanggal '+fmtD(p.p_my_date)+'.','jadwal');return{ok:true,id:s.id};
    },
    u_swap_respond(p){
      const id=this.sess(p.p_token),s=this.db.shift_swaps.find(x=>x.id===p.p_id);if(!s||s.partner_employee_id!==id)throw new Error('Permintaan tidak ditemukan');
      if(s.status!=='waiting_partner')throw new Error('Permintaan sudah diproses');
      if(p.p_accept){s.status='waiting_coordinator';this.notifyHrd('Tukar shift menunggu persetujuan','Ada permintaan tukar shift yang sudah disetujui kedua pegawai.','jadwal-shift')}
      else{s.status='rejected';s.decided_by='Pegawai pengganti';s.decided_at=this.now();this.notify(s.requester_employee_id,'Tukar shift ditolak','Pegawai pengganti menolak permintaan tukar shift.','jadwal')}
      return{ok:true};
    },
    u_swap_decide(p){throw new Error('Tukar shift diputuskan HRD melalui panel admin')},
    adm_whoami(){return this.adm?{id:this.adm.id,full_name:this.adm.full_name,role:this.adm.role,employee_id:this.adm.employee_id}:null},
    adm_claim_first_admin(){throw new Error('Akun ini belum terdaftar sebagai admin. Minta Super Admin menambahkan.')},
    adm_decide_final(p){
      if(!this.isHr())throw new Error('Hanya HRD Manager atau Super Admin yang dapat memberi persetujuan akhir');
      const r=this.db.requests.find(x=>x.id===p.p_id);if(!r)throw new Error('Pengajuan tidak ditemukan');
      if(r.status==='waiting_l1'&&!p.p_override)throw new Error('Pengajuan masih menunggu Kepala Departemen');
      if(this.adm.employee_id&&this.adm.employee_id===r.employee_id)throw new Error('Tidak dapat memutuskan pengajuan milik sendiri');
      if(this.adm.role==='manajer_hrd'&&this.db.admin_profiles.some(a=>a.employee_id===r.employee_id&&a.role==='manajer_hrd'))throw new Error('Pengajuan HRD Manager harus diputuskan oleh Super Admin');
      return this.finalize(p.p_id,p.p_decision,r.status==='waiting_l1'?(p.p_note||'')+' [override HRD tanpa tahap 1]':p.p_note,this.adminName(),this.adm.employee_id,2);
    },
    adm_decide_cancel(p){
      if(!this.isHr())throw new Error('Hanya HRD Manager atau Super Admin yang dapat memutuskan pembatalan');
      const d=this.db,r=d.requests.find(x=>x.id===p.p_id);if(!r||!r.cancel_requested||r.status!=='approved')throw new Error('Tidak ada permintaan pembatalan untuk pengajuan ini');
      const rt=d.request_types.find(t=>t.id===r.request_type_id),actor=this.adminName();
      if(p.p_decision==='approved'){r.status='cancelled';r.cancelled_at=this.now();r.cancel_requested=false;r.applied=false;
        if(rt.deducts_balance&&r.total_days>0)d.leave_balance_entries.push({id:this.nid('leave_balance_entries'),employee_id:r.employee_id,request_type_id:rt.id,period_key:Engine.periodKey(this.ctxAll(),r.start_date),entry_type:'refund',days:r.total_days,request_id:r.id,note:'Pembatalan '+rt.name,created_by:actor,valid_until:null,created_at:this.now()});
        d.request_approvals.push({id:this.nid('request_approvals'),request_id:r.id,stage:2,approver_name:actor,approver_employee_id:null,decision:'cancel_approved',note:p.p_note||null,decided_at:this.now()});
        this.notify(r.employee_id,'Pembatalan disetujui','Pembatalan '+rt.name+' tanggal '+fmtD(r.start_date)+' disetujui.','pengajuan')}
      else{r.cancel_requested=false;d.request_approvals.push({id:this.nid('request_approvals'),request_id:r.id,stage:2,approver_name:actor,approver_employee_id:null,decision:'cancel_rejected',note:p.p_note||null,decided_at:this.now()});this.notify(r.employee_id,'Pembatalan ditolak','Permintaan pembatalan '+rt.name+' ditolak. '+(p.p_note||''),'pengajuan')}
      this.audit(actor,'request.cancel_'+p.p_decision,'requests',r.id,null,{note:p.p_note});
      return{ok:true,employee_id:r.employee_id,start_date:r.start_date,end_date:r.end_date};
    },
    adm_create_employees(p){
      const d=this.db;let ok=0;const skip=[];
      for(const r of p.p_rows){
        const no=String(r.employee_no||'').trim();if(!no||!String(r.full_name||'').trim()){skip.push({employee_no:no,reason:'NIP dan nama wajib diisi'});continue}
        if(d.employees.some(e=>e.employee_no===no)){skip.push({employee_no:no,reason:'NIP sudah ada'});continue}
        if(!String(r.email||'').trim()){skip.push({employee_no:no,reason:'Email wajib diisi (dipakai untuk login)'});continue}
        {const v=empValidate(r);if(v){skip.push({employee_no:no,reason:v});continue}
         if(r.email&&d.employees.some(e=>(e.email||'').toLowerCase()===String(r.email).trim().toLowerCase())){skip.push({employee_no:no,reason:'Email sudah dipakai pegawai lain'});continue}}
        let dept=r.department_id||null;const dn=String(r.department_name||'').trim();
        if(!dept&&dn){let x=d.departments.find(y=>y.name.toLowerCase()===dn.toLowerCase());if(!x){x={id:this.nid('departments'),name:dn,head_employee_id:null,is_active:true};d.departments.push(x)}dept=x.id}
        const id=this.nid('employees');
        d.employees.push({id,employee_no:no,must_change_password:true,full_name:String(r.full_name).trim(),department_id:dept,position:r.position||null,type:r.type||'staff',employment_status:r.employment_status||'permanent',roles:(r.roles&&r.roles.length?r.roles:['pegawai']),join_date:r.join_date||TODAY(),end_date:null,phone:r.phone||null,email:String(r.email).trim().toLowerCase(),photo_url:r.photo_url||null,title_suffix:r.title_suffix||null,nip:no,division:r.division||null,...empX(r),is_active:true,failed_login_count:0,locked_until:null,last_login_at:null,created_at:this.now(),updated_at:this.now()});
        d.pw[id]=r.password||no;if(!d.employee_fingerprints.some(f=>f.fingerprint_id===no))d.employee_fingerprints.push({id:this.nid('employee_fingerprints'),employee_id:id,fingerprint_id:no});ok++;
      }
      return{ok:true,inserted:ok,skipped:skip};
    },
    adm_set_employee_password(p){const e=this.emp(p.p_id);if(String(p.p_password||'').length<6)throw new Error('Kata sandi minimal 6 karakter');this.db.pw[e.id]=p.p_password;e.must_change_password=true;Object.keys(this.db.sessions).forEach(k=>{if(this.db.sessions[k]===e.id)delete this.db.sessions[k]});return{ok:true}},
    adm_mark_applied(p){this.db.requests.filter(r=>(p.p_ids||[]).includes(r.id)&&r.status==='approved').forEach(r=>r.applied=true);return{ok:true}},
    adm_swap_decide(p){return this.swapApply(p.p_id,p.p_decision,this.adminName())},
    adm_lock_period(p){
      if(!this.isHr())throw new Error('Hanya HRD Manager atau Super Admin yang dapat mengunci periode');
      const d=this.db;let x=d.attendance_periods.find(a=>a.year===p.p_year&&a.month===p.p_month);
      if(!x){x={id:this.nid('attendance_periods'),year:p.p_year,month:p.p_month,status:'open'};d.attendance_periods.push(x)}
      Object.assign(x,{status:'locked',locked_by:this.adminName(),locked_at:this.now()});
      const a=p.p_year+'-'+pad(p.p_month)+'-01',b=D.mEnd(a);d.attendance_records.filter(r=>r.work_date>=a&&r.work_date<=b).forEach(r=>r.is_locked=true);
      this.audit(this.adminName(),'period.lock','attendance_periods',p.p_year+'-'+p.p_month);return{ok:true};
    },
    adm_reopen_period(p){
      if(!this.isHr())throw new Error('Hanya HRD Manager atau Super Admin yang dapat membuka periode');
      if(!String(p.p_reason||'').trim())throw new Error('Alasan pembukaan kembali wajib diisi');
      const d=this.db,x=d.attendance_periods.find(a=>a.year===p.p_year&&a.month===p.p_month);if(x){x.status='open';x.reopened_by=this.adminName();x.reopened_at=this.now()}
      const a=p.p_year+'-'+pad(p.p_month)+'-01',b=D.mEnd(a);d.attendance_records.filter(r=>r.work_date>=a&&r.work_date<=b).forEach(r=>r.is_locked=false);
      this.audit(this.adminName(),'period.reopen','attendance_periods',p.p_year+'-'+p.p_month,null,{reason:p.p_reason});return{ok:true};
    }
  }
};

/* ---------- akses tabel (admin) ---------- */
const PVDB={
  rows(t){PV.init();if(!PV.db[t])throw new Error('Tabel tidak dikenal: '+t);return PV.db[t]},
  async select(t,o={}){
    await sleep(15);let r=this.rows(t).slice();
    if(o.eq)for(const k in o.eq)r=r.filter(x=>x[k]===o.eq[k]);
    if(o.in)for(const k in o.in)r=r.filter(x=>o.in[k].includes(x[k]));
    if(o.gte)for(const k in o.gte)r=r.filter(x=>x[k]>=o.gte[k]);
    if(o.lte)for(const k in o.lte)r=r.filter(x=>x[k]<=o.lte[k]);
    if(t==='employees')r=r.map(({...e})=>e);
    const ord=o.order||null;if(ord){const k=ord[0],asc=ord[1]!==false;r.sort((a,b)=>(a[k]>b[k]?1:a[k]<b[k]?-1:0)*(asc?1:-1))}
    if(o.limit)r=r.slice(0,o.limit);
    return clone(r);
  },
  lockCheck(t,row){if(t==='attendance_records'&&(row.is_locked||PV.periodLocked(row.work_date)))throw new Error('PERIOD_LOCKED: Periode sudah dikunci')},
  async insert(t,rows){const a=this.rows(t),out=[];for(const r of rows){const o={...r};if(t!=='settings'&&o.id==null)o.id=PV.nid(t);if(o.created_at===undefined&&['notifications','audit_logs','leave_balance_entries'].includes(t))o.created_at=PV.now();if(t==='audit_logs'&&!o.occurred_at)o.occurred_at=PV.now();this.lockCheck(t,o);a.push(o);out.push(o)}return clone(out)},
  async upsert(t,rows,onConflict,ignoreDup){
    const a=this.rows(t),keys=(onConflict||(t==='settings'?'key':'id')).split(','),out=[];
    for(const r of rows){const i=a.findIndex(x=>keys.every(k=>x[k]===r[k]));
      if(i>=0){if(ignoreDup)continue;this.lockCheck(t,a[i]);Object.assign(a[i],r);out.push(a[i])}
      else{const o={...r};if(t!=='settings'&&o.id==null)o.id=PV.nid(t);this.lockCheck(t,o);a.push(o);out.push(o)}}
    return clone(out);
  },
  async update(t,patch,match){const a=this.rows(t).filter(x=>Object.keys(match).every(k=>x[k]===match[k]));a.forEach(x=>{this.lockCheck(t,x);Object.assign(x,patch)});return clone(a)},
  async remove(t,match){const a=this.rows(t);const rm=a.filter(x=>Object.keys(match).every(k=>x[k]===match[k]));rm.forEach(x=>this.lockCheck(t,x));PV.db[t]=a.filter(x=>!rm.includes(x));
    if(t==='departments')PV.db.employees.forEach(e=>{if(rm.some(d=>d.id===e.department_id))e.department_id=null});
    if(t==='employees'){const ids=rm.map(x=>x.id);['schedule_assignments','shift_assignments','raw_logs','requests','attendance_records','employee_fingerprints','leave_balance_entries'].forEach(k=>PV.db[k]=PV.db[k].filter(x=>!ids.includes(x.employee_id)));PV.db.departments.forEach(d=>{if(ids.includes(d.head_employee_id))d.head_employee_id=null})}
    if(t==='work_schedules')PV.db.work_schedule_days=PV.db.work_schedule_days.filter(x=>!rm.some(w=>w.id===x.work_schedule_id));
    return clone(rm)}
};
async function ensureSB(){
  if(S.sb)return S.sb;
  if(!/^https?:\/\//.test(SUPABASE_URL)||SUPABASE_ANON_KEY==='ANON_KEY_ANDA')throw new Error('URL dan Anon Key Supabase belum diisi di bagian KONFIGURASI pada kode aplikasi.');
  const m=await import(SUPABASE_ESM);S.sb=m.createClient(SUPABASE_URL,SUPABASE_ANON_KEY,{auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:false}});return S.sb;
}
const SBDB={
  err(e){const m=new Error(e.message||String(e));if(/PERIOD_LOCKED/.test(e.message||''))m.message='PERIOD_LOCKED: Periode sudah dikunci';else if(empMiss(e.message||''))m.message=LEGACY_MSG;return m},
  cols(t){return t==='employees'?(S.legacyEmp?EMP_COLS_BASE:EMP_COLS):'*'},
  async select(t,o={}){
    const out=[],page=1000;let from=0;const ord=o.order||[t==='settings'?'key':'id',true];
    while(true){
      const mk=()=>{let q=S.sb.from(t).select(o.cols||this.cols(t));
        for(const k in(o.eq||{}))q=q.eq(k,o.eq[k]);
        for(const k in(o.in||{}))q=q.in(k,o.in[k]);
        for(const k in(o.gte||{}))q=q.gte(k,o.gte[k]);
        for(const k in(o.lte||{}))q=q.lte(k,o.lte[k]);
        return q.order(ord[0],{ascending:ord[1]!==false})};
      let q=mk();
      if(o.limit){const{data,error}=await q.limit(o.limit);if(error)throw this.err(error);return data}
      const{data,error}=await q.range(from,from+page-1);if(error){if(t==='employees'&&!o.cols&&!S.legacyEmp&&empMiss(error.message||'')){S.legacyEmp=true;return this.select(t,o)}throw this.err(error)}
      out.push(...data);if(data.length<page)break;from+=page;
      if(o.par>1){/* tabel besar: halaman berikutnya diambil paralel */
        let more=true;
        while(more){const bt=await Promise.all(Array.from({length:o.par},(_,i)=>mk().range(from+i*page,from+(i+1)*page-1)));
          for(const r of bt){if(r.error)throw this.err(r.error);out.push(...r.data);if(r.data.length<page)more=false}
          from+=o.par*page}
        break}
    }
    return out;
  },
  async insert(t,rows){const out=[];for(let i=0;i<rows.length;i+=500){const{data,error}=await S.sb.from(t).insert(rows.slice(i,i+500)).select(this.cols(t));if(error)throw this.err(error);out.push(...data)}return out},
  async upsert(t,rows,onConflict,ignoreDup){const out=[];for(let i=0;i<rows.length;i+=500){const{data,error}=await S.sb.from(t).upsert(rows.slice(i,i+500),{onConflict,ignoreDuplicates:!!ignoreDup}).select(this.cols(t));if(error)throw this.err(error);out.push(...(data||[]))}return out},
  /* upsert bertahap: kepingan paralel + callback progres. o.cols=false -> tanpa mengembalikan baris */
  async upsertP(t,rows,onConflict,ignoreDup,o={}){
    const parts=chunk(rows,o.size||1000),out=[];let done=0,idx=0,err=null;
    const worker=async()=>{while(idx<parts.length&&!err){const part=parts[idx++];
      try{let q=S.sb.from(t).upsert(part,{onConflict,ignoreDuplicates:!!ignoreDup});if(o.cols!==false)q=q.select(o.cols||this.cols(t));
        const{data,error}=await q;if(error)throw this.err(error);if(data)out.push(...data)}
      catch(e){err=e;return}
      done+=part.length;if(o.onProgress)o.onProgress(done,rows.length)}};
    await Promise.all(Array.from({length:Math.min(o.conc||4,parts.length)},worker));
    if(err)throw err;return out},
  async update(t,patch,match){let q=S.sb.from(t).update(patch);for(const k in match)q=q.eq(k,match[k]);const{data,error}=await q.select(this.cols(t));if(error)throw this.err(error);return data},
  async remove(t,match){let q=S.sb.from(t).delete();for(const k in match)q=q.eq(k,match[k]);const{data,error}=await q.select(this.cols(t));if(error)throw this.err(error);return data}
};
PVDB.upsertP=async function(t,rows,k,i,o){const r=await this.upsert(t,rows,k,i);if(o&&o.onProgress)o.onProgress(rows.length,rows.length);return r};
const DB=IS_PREVIEW?PVDB:SBDB;
async function rpc(name,params){
  if(IS_PREVIEW)return PV.rpc(name,params);
  const sb=await ensureSB();const{data,error}=await sb.rpc(name,params||{});if(error)throw new Error(error.message);return data;
}
const actor=()=>S.mode==='admin'?S.admin.full_name+' ('+S.admin.role+')':(S.me?S.me.full_name:'sistem');
async function audit(action,entity,id,oldv,newv){try{await DB.insert('audit_logs',[{actor:actor(),action,entity_type:entity,entity_id:String(id==null?'':id),old_value:oldv||null,new_value:newv||null}])}catch(e){console.warn('audit gagal',e)}}
async function notifyEmp(emp,title,body,link){try{await DB.insert('notifications',[{employee_id:emp,audience:'employee',title,body,link:link||null}])}catch(e){console.warn(e)}}

/* ---------- konteks data ---------- */
const chunk=(a,n)=>{const o=[];for(let i=0;i<a.length;i+=n)o.push(a.slice(i,i+n));return o};
async function loadAdminCtx(from,to){
  const f1=D.add(from,-1),t1=D.add(to,1),fromTs=new Date(D.epoch(f1,'00:00')).toISOString(),toTs=new Date(D.epoch(D.add(to,2),'00:00')).toISOString();
  const[employees,departments,work_schedules,work_schedule_days,schedule_assignments,shifts,shift_assignments,holidays,raw_logs,attendance_records,request_types,rr,w1,w2,cr,ledger,settingsRows,periods,swaps,subs,academic_years,fps,notifs]=await Promise.all([
    DB.select('employees'),DB.select('departments'),DB.select('work_schedules'),DB.select('work_schedule_days'),DB.select('schedule_assignments'),DB.select('shifts'),
    DB.select('shift_assignments',{gte:{work_date:f1},lte:{work_date:t1}}),DB.select('holidays',{gte:{holiday_date:f1},lte:{holiday_date:t1}}),
    DB.select('raw_logs',{cols:'id,employee_id,fingerprint_id,logged_at,source,match_status',gte:{logged_at:fromTs},lte:{logged_at:toTs},par:4}),
    DB.select('attendance_records',{gte:{work_date:from},lte:{work_date:to},par:4}),DB.select('request_types'),
    DB.select('requests',{lte:{start_date:to},gte:{end_date:from}}),DB.select('requests',{eq:{status:'waiting_l1'}}),DB.select('requests',{eq:{status:'waiting_l2'}}),DB.select('requests',{eq:{cancel_requested:true}}),
    DB.select('leave_balance_entries'),DB.select('settings'),DB.select('attendance_periods'),DB.select('shift_swaps'),DB.select('department_substitutes'),DB.select('academic_years'),DB.select('employee_fingerprints'),
    DB.select('notifications',{eq:{audience:'hrd'},order:['id',false],limit:60})
  ]);
  const rm={};[...rr,...w1,...w2,...cr].forEach(r=>rm[r.id]=r);const requests=Object.values(rm).sort((a,b)=>b.id-a.id);
  let approvals=[];for(const c of chunk(requests.map(r=>r.id),120))approvals.push(...await DB.select('request_approvals',{in:{request_id:c}}));
  const settings=Object.fromEntries(settingsRows.map(s=>[s.key,s.value]));
  return{me:null,today:TODAY(),server_now:new Date().toISOString(),settings,settingsRows,departments,employees,work_schedules,work_schedule_days,schedule_assignments,shifts,shift_assignments,holidays,raw_logs,attendance_records,request_types,requests,approvals,ledger,leave_balance_entries:ledger,swaps,substitutes:subs,academic_years,fingerprints:fps,notifications:notifs,periods};
}
async function loadUserCtx(from,to){const c=await rpc('u_context',{p_token:S.token,p_from:from,p_to:to});c.leave_balance_entries=c.ledger;return c}
const Ctx={
  async load(from,to,force){
    const key=S.mode+'|'+from+'|'+to;if(!force&&S.cache.key===key&&S.ctx)return S.ctx;
    const c=S.mode==='admin'?await loadAdminCtx(from,to):await loadUserCtx(from,to);
    S.ctx=c;S.cache.key=key;S.range={from,to};S.settings=c.settings||{};TZH=TZMAP[S.settings.timezone]||7;
    S.skew=IS_PREVIEW?0:(Date.parse(c.server_now)-Date.now());if(Math.abs(S.skew)<3000)S.skew=0;
    if(c.me)S.me=c.me;
    const sn=$('#sk-school');if(sn)sn.textContent=APP_NAME;
    return c;
  },
  reset(){S.cache.key=null},
  async month(ds,force){return this.load(D.mStart(ds),D.mEnd(ds),force)}
};


/* =====================================================================
   8. PENGOLAHAN BERSAMA (rekap, penyimpanan, pembantu tampilan)
   ===================================================================== */
const VIEWS={},MENUS=[];
let RT=0;
const reg=(k,def)=>{VIEWS[k]=def};
const empIx=()=>Engine.ix(S.ctx).emp;
const nameOf=id=>(empIx()[id]||(S.people&&S.people[id])||{}).full_name||'-';
const deptOf=id=>{const e=empIx()[id];if(!e)return'-';const d=(S.ctx.departments||[]).find(x=>x.id===e.department_id);return d?d.name:'-'};
const deptName=id=>{const d=(S.ctx.departments||[]).find(x=>x.id===id);return d?d.name:'-'};
const rtOf=id=>(S.ctx.request_types||[]).find(t=>t.id===id)||{name:'-',category:''};
const isMgr=()=>S.mode==='admin'&&['manajer_hrd','admin_sistem'].includes(S.admin.role);
const isSysAdmin=()=>S.mode==='admin'&&S.admin.role==='admin_sistem';
const isHead=()=>S.mode==='user'&&S.me&&(S.me.head_dept_ids||[]).length>0;
const hasRole=r=>S.mode==='user'&&S.me&&(S.me.roles||[]).includes(r);
const ymStart=()=>S.ym||(S.ym=D.mStart(TODAY()));
/* Periode payroll: tanggal 26 bulan sebelumnya s.d. 25 bulan berjalan. Label periode = bulan tempat periode berakhir. */
const PAY_DAY=26;
const payLabel=ds=>D.parse(ds).d>=PAY_DAY?D.mStart(D.add(D.mStart(ds),32)):D.mStart(ds);
const payRange=lbl=>{const pv=D.mStart(D.add(lbl,-1));return[pv.slice(0,8)+pad(PAY_DAY),lbl.slice(0,8)+pad(PAY_DAY-1)]};
const pmLabel=()=>S.pm||(S.pm=payLabel(TODAY()));
const payTitle=()=>{const r=payRange(pmLabel());return'Payroll '+fmtM(pmLabel())+' ('+fmtD(r[0])+' s.d. '+fmtD(r[1])+')'};
const payBar=()=>{const l=pmLabel(),r=payRange(l);return'<div class="sk-row" style="gap:6px"><button class="sk-ibtn" data-act="pnav" data-d="-1" title="Periode sebelumnya" aria-label="Periode sebelumnya">'+ic('left')+'</button><div class="sk-pbar"><b>Payroll '+esc(fmtM(l))+'</b><small>'+esc(fmtD(r[0]))+' s.d. '+esc(fmtD(r[1]))+'</small></div><button class="sk-ibtn" data-act="pnav" data-d="1" title="Periode berikutnya" aria-label="Periode berikutnya">'+ic('right')+'</button></div>'};
ACT.pnav=async t=>{S.pm=D.mStart(D.add(pmLabel(),+t.dataset.d>0?32:-1));await render()};
function recFor(emp,ds){const ix=Engine.ix(S.ctx),ex=ix.att[emp+'|'+ds];if(ex&&(ex.is_locked||ex.is_manual))return ex;return Engine.recalcDay(S.ctx,emp,ds)}
async function persistRecs(recs,onP){
  if(!recs.length)return 0;
  const parts=chunk(recs,500);let saved=0,done=0,idx=0,err=null;
  /* periode terkunci: belah dua berulang agar hanya baris terkunci yang dilewati (bukan satu per satu) */
  const put=async part=>{
    try{await DB.upsertP('attendance_records',part,'employee_id,work_date',false,{cols:false,size:part.length,conc:1});return part.length}
    catch(e){if(!/PERIOD_LOCKED|dikunci/i.test(e.message||''))throw e;if(part.length<2)return 0;const h=part.length>>1;return(await put(part.slice(0,h)))+(await put(part.slice(h)))}};
  const worker=async()=>{while(idx<parts.length&&!err){const p=parts[idx++];
    try{const n=await put(p);saved+=n}catch(e){err=e;return}
    done+=p.length;if(onP)onP(done,recs.length)}};
  await Promise.all(Array.from({length:Math.min(4,parts.length)},worker));
  if(err)throw err;return saved
}
async function recalcAndSave(c,empIds,from,to){Engine.touch(c);const recs=Engine.recalcRange(c,empIds,from,to);return persistRecs(recs)}
/* hitung ulang rekap untuk pegawai + rentang tertentu (setelah persetujuan, pembatalan, tukar shift) */
async function applyRange(emp,start,end){
  const t=TODAY();if(start>t)return 0;end=end>t?t:end;
  const c=await loadAdminCtx(start,end);return recalcAndSave(c,[emp],start,end);
}
async function reprocess(from,to,ids,pg){const t=TODAY();if(to>t)to=t;if(from>to)return 0;
  if(pg)pg('load');const c=await loadAdminCtx(from,to);
  if(!pg)return recalcAndSave(c,ids||c.employees.filter(e=>e.is_active).map(e=>e.id),from,to);
  pg('calc');await new Promise(r=>setTimeout(r,40));/* beri kesempatan layar menggambar sebelum hitung berat */
  Engine.touch(c);const recs=Engine.recalcRange(c,ids||c.employees.filter(e=>e.is_active).map(e=>e.id),from,to);
  return persistRecs(recs,(d,n)=>pg('save',d,n))}
/* jatah otomatis (masa kerja >= minimal) + hangus tiap 25 Juni */
async function leaveAuto(c){
  try{
    if(!(isMgr()||(S.admin&&S.admin.role==='staf_hrd')))return 0;
    const today=TODAY(),rts=(c.request_types||[]).filter(t=>t.code==='cuti_tahunan'&&t.deducts_balance&&+t.annual_quota_days>0);
    if(!rts.length)return 0;
    const mode=Engine.set(c,'leave_period_mode','calendar'),led=Engine.ledger(c),rows=[];
    for(const rt of rts){
      const per=Engine.periodKey(c,today),minM=+rt.min_service_months||0;
      const have=new Set(led.filter(l=>l.request_type_id===rt.id&&l.period_key===per&&l.entry_type==='accrual').map(l=>l.employee_id));
      for(const e of c.employees.filter(x=>x.is_active&&x.join_date)){
        if(have.has(e.id)||Engine.svcMonths(e.join_date,today)<minM)continue;
        rows.push({employee_id:e.id,request_type_id:rt.id,period_key:per,entry_type:'accrual',days:+rt.annual_quota_days,note:'Jatah otomatis '+rt.name+' '+per+' (masa kerja '+minM+' bulan terpenuhi)',created_by:'sistem'});
      }
      if(mode==='jun25'){
        const prev=Engine.leaveKey(D.add(Engine.leaveStart(today),-1));
        const done=new Set(led.filter(l=>l.request_type_id===rt.id&&l.period_key===prev&&l.entry_type==='expire').map(l=>l.employee_id));
        for(const e of c.employees){if(done.has(e.id))continue;const bal=Engine.balance(c,e.id,rt.id,prev);if(bal>0)rows.push({employee_id:e.id,request_type_id:rt.id,period_key:prev,entry_type:'expire',days:-bal,note:'Hangus 25 Juni (periode '+prev+')',created_by:'sistem'})}
      }
    }
    if(rows.length){await DB.insert('leave_balance_entries',rows);if(Ctx.reset)Ctx.reset()}
    return rows.length;
  }catch(e){console.warn('leaveAuto',e);return 0}
}
async function processPending(){
  try{
    const today=TODAY(),sRows=await DB.select('settings');const last=(sRows.find(s=>s.key==='last_processed_date')||{}).value;
    const un=await DB.select('requests',{eq:{status:'approved',applied:false}});
    let from=last?D.add(last,-1):D.add(today,-14);const floor=D.add(today,-120);if(from<floor)from=floor;
    for(const r of un){const s=String(r.start_date).slice(0,10);if(s<from&&s>=floor)from=s}
    const c=await loadAdminCtx(from,today);S.settings=c.settings;
    await leaveAuto(c);
    const n=await recalcAndSave(c,c.employees.filter(e=>e.is_active).map(e=>e.id),from,today);
    if(un.length)await rpc('adm_mark_applied',{p_ids:un.map(r=>r.id)});
    if(isMgr()||S.admin.role==='staf_hrd')await DB.upsert('settings',[{key:'last_processed_date',value:today,description:SETTING_DESC['last_processed_date']}],'key').catch(()=>{});
    return n;
  }catch(e){console.warn('processPending',e);return 0}
}
const monthBar=(act='mnav')=>{const y=ymStart();return'<div class="sk-row" style="gap:6px"><button class="sk-ibtn" data-act="'+act+'" data-d="-1" title="Bulan sebelumnya">'+ic('left')+'</button><b style="min-width:130px;text-align:center">'+fmtM(y)+'</b><button class="sk-ibtn" data-act="'+act+'" data-d="1" title="Bulan berikutnya">'+ic('right')+'</button></div>'};
ACT.mnav=async(t)=>{S.ym=D.mStart(D.add(ymStart(),+t.dataset.d>0?32:-1));await render()};
const sideEffectsReload=async()=>{Ctx.reset();await render(true)};
const empOptions=(list,o={})=>(list||S.ctx.employees).filter(e=>o.all||e.is_active).slice().sort((a,b)=>a.full_name.localeCompare(b.full_name)).map(e=>[e.id,e.full_name+' ('+e.employee_no+')']);
const deptOptions=()=>(S.ctx.departments||[]).filter(d=>d.is_active).map(d=>[d.id,d.name]);

/* =====================================================================
   FILTER DEPARTEMEN & DIVISI
   Pilihan diambil dari data karyawan (department_id & kolom division),
   bukan dari daftar tetap sistem. Divisi menyesuaikan departemen terpilih.
   ===================================================================== */
S.ef=S.ef||{};
const EF_NONE='__none',EF_CB={};
const efSt=k=>S.ef[k]||(S.ef[k]={dept:'',div:'',gender:'',lim:'25'});
const EF_LIMS=[['10','Tampil 10'],['25','Tampil 25'],['50','Tampil 50'],['100','Tampil 100'],['all','Tampil semua']];
const efPer=(k,d)=>{const l=efSt(k).lim;return l==='all'?1e9:(+l>0?+l:d)};
const efOn=k=>{const s=efSt(k);return!!(s.dept||s.div||s.gender)};
const efDiv=e=>String(e&&e.division!=null?e.division:'').replace(/\s+/g,' ').trim();
const efHasDept=e=>e.department_id!=null&&e.department_id!==''&&deptName(e.department_id)!=='-';
const efDeptOk=(e,v)=>!v||(v===EF_NONE?!efHasDept(e):efHasDept(e)&&String(e.department_id)===String(v));
const efDivOk=(e,v)=>!v||(v===EF_NONE?!efDiv(e):efDiv(e).toLowerCase()===String(v).toLowerCase());
const efPass=(k,e)=>{const s=efSt(k);return(!s.dept&&!s.div&&!s.gender)||(!!e&&efDeptOk(e,s.dept)&&efDivOk(e,s.div)&&(!s.gender||e.gender===s.gender))};
const efPassId=(k,id)=>efPass(k,empIx()[id]);
function efOpts(pool,dept){
  const all=pool||[],dm=new Map();let noD=false;
  all.forEach(e=>{if(efHasDept(e))dm.set(String(e.department_id),deptName(e.department_id));else noD=true});
  const depts=[...dm.entries()].sort((a,b)=>a[1].localeCompare(b[1],'id'));if(noD&&depts.length)depts.push([EF_NONE,'Tanpa departemen']);
  const vm=new Map();let noV=false;
  all.filter(e=>efDeptOk(e,dept)).forEach(e=>{const v=efDiv(e);if(!v)noV=true;else if(!vm.has(v.toLowerCase()))vm.set(v.toLowerCase(),v)});
  const divs=[...vm.values()].sort((a,b)=>a.localeCompare(b,'id')).map(v=>[v,v]);if(noV&&divs.length)divs.push([EF_NONE,'Tanpa divisi']);
  return{depts,divs}
}
/* pastikan pilihan tersimpan masih ada di data (mis. setelah data karyawan berubah) */
function efValid(s,pool){
  if(s.dept&&!efOpts(pool,'').depts.some(x=>x[0]===String(s.dept)))s.dept='';
  if(s.div&&!efOpts(pool,s.dept).divs.some(x=>x[0].toLowerCase()===String(s.div).toLowerCase()))s.div='';
  return s
}
const efFix=(k,pool)=>efValid(efSt(k),pool);
function efBar(k,pool,o={}){
  const s=efValid(efSt(k),pool),od=efOpts(pool,'').depts,ov=efOpts(pool,s.dept).divs;
  if(!(pool||[]).length)return'';
  const dv=(ov.find(x=>x[0].toLowerCase()===String(s.div).toLowerCase())||[s.div])[0],on=!!(s.dept||s.div||s.gender),at=n=>'data-chg="efChg" data-ef="'+k+'" aria-label="'+n+'"';
  return'<div class="ef-bar'+(o.inl?' ef-inl':'')+'">'
    +(od.length?sel('dept',[['','Semua departemen'],...od],s.dept,{attr:at('Filter departemen')}):'')
    +(ov.length?sel('div',[['','Semua divisi'],...ov],dv,{attr:at('Filter divisi')}):'')
    +sel('gender',[['','Semua jenis kelamin'],...Object.entries(GENDER)],s.gender||'',{attr:at('Filter jenis kelamin')})
    +(o.nolim?'':sel('lim',EF_LIMS,String(s.lim||'25'),{attr:at('Jumlah data ditampilkan')}))
    +(on?'<button type="button" class="sk-b ghost sm" data-act="efReset" data-ef="'+k+'">'+ic('x','sm')+'Reset</button>':'')+'</div>'
}
CHG.efChg=async t=>{const k=t.dataset.ef;efSt(k)[t.name]=t.value;const cb=EF_CB[k];if(cb)cb();else await render()};
ACT.efReset=async t=>{const k=t.dataset.ef,lm=efSt(k).lim;S.ef[k]={dept:'',div:'',gender:'',lim:lm};const cb=EF_CB[k];if(cb)cb();else await render()};
/* versi untuk daftar centang di dalam jendela (modal): menyaring baris di tempat, centang tetap tersimpan */
const efTr=e=>{const dp=efHasDept(e)?e.department_id:'',dv=efDiv(e);return' data-g="'+esc(e.gender||'')+'" data-dp="'+esc(dp)+'" data-dv="'+esc(dv.toLowerCase())+'" data-dvt="'+esc(dv)+'"'};
function efModBar(pool){
  const o=efOpts(pool,'');if(!(pool||[]).length)return'';
  return'<div class="ef-bar ef-mod" data-efm="1">'+(o.depts.length?sel('_efd',[['','Semua departemen'],...o.depts],'',{attr:'data-chg="efModal" aria-label="Filter departemen"'}):'')+(o.divs.length?sel('_efv',[['','Semua divisi'],...o.divs],'',{attr:'data-chg="efModal" aria-label="Filter divisi"'}):'')+sel('_efg',[['','Semua jenis kelamin'],...Object.entries(GENDER)],'',{attr:'data-chg="efModal" aria-label="Filter jenis kelamin"'})+'<span class="sk-sub" data-efm-n></span></div>'
}
CHG.efModal=t=>{
  const w=t.closest('[data-efm]'),root=t.closest('.sk-ov')||document,dS=$('[name=_efd]',w),vS=$('[name=_efv]',w),gS=$('[name=_efg]',w),gv=gS?gS.value:'',rows=$$('tr[data-dp]',root);
  const dOk=r=>!dS||!dS.value||(dS.value===EF_NONE?r.dataset.dp==='':r.dataset.dp===dS.value);
  if(t===dS&&vS){const seen=new Map();let none=false;rows.filter(dOk).forEach(r=>{const v=r.dataset.dvt;if(!v)none=true;else if(!seen.has(v.toLowerCase()))seen.set(v.toLowerCase(),v)});
    const cur=vS.value,list=[...seen.values()].sort((a,b)=>a.localeCompare(b,'id')).map(v=>[v,v]);if(none&&list.length)list.push([EF_NONE,'Tanpa divisi']);
    vS.innerHTML=[['','Semua divisi'],...list].map(([v,l])=>'<option value="'+esc(v)+'"'+(v.toLowerCase()===String(cur).toLowerCase()?' selected':'')+'>'+esc(l)+'</option>').join('')}
  const v=vS?vS.value:'';let n=0;
  rows.forEach(r=>{const ok=dOk(r)&&(!v||(v===EF_NONE?r.dataset.dv==='':r.dataset.dv===v.toLowerCase()))&&(!gv||r.dataset.g===gv);r.style.display=ok?'':'none';if(ok)n++});
  const m=$('[data-efm-n]',w);if(m)m.textContent=((dS&&dS.value)||v||gv)?n+' pegawai ditampilkan. Pegawai yang sudah dicentang tetap ikut diproses walau tersembunyi.':''
};
const rtOptions=(cats)=>(S.ctx.request_types||[]).filter(t=>t.is_active!==false&&(!cats||cats.includes(t.category))).map(t=>[t.id,t.name]);
const fileRead=f=>new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=()=>rej(new Error('Gagal membaca berkas'));r.readAsArrayBuffer(f)});
const fmtMon=(y,m)=>MON[m-1]+' '+y;
function periodStatus(y,m){const p=(S.ctx.periods||[]).find(x=>x.year===y&&x.month===m);return p?p.status:'open'}
const isLockedDate=ds=>{const p=D.parse(ds);return periodStatus(p.y,p.m)==='locked'};

/* =====================================================================
   9. PWA (manifest data URI di <head>, Service Worker dari Blob)
   ===================================================================== */
/* Ikon aplikasi diambil dari logo lembaga. Untuk hasil paling tajam ganti dengan PNG persegi >= 512x512 (harus bisa diakses lintas domain/CORS). */
const PWA_ICON_URL=LOGO_URL;
/* Opsional: bila Anda punya berkas sw.js di domain yang sama, isi alamatnya di sini agar Service Worker aktif. Blogger tidak bisa menyimpan berkas sw.js. */
const PWA_SW_URL='/sw.js';
const pwaStandalone=()=>{try{return matchMedia('(display-mode: standalone)').matches||navigator.standalone===true}catch(e){return false}};
const pwaIOS=()=>/iphone|ipad|ipod/i.test(navigator.userAgent)||(navigator.platform==='MacIntel'&&navigator.maxTouchPoints>1);
/* gambar logo di atas kanvas putih menjadi ikon PNG (192, 512, maskable 512, dan 180 untuk iOS) */
async function pwaIcons(){
  const img=await new Promise((res,rej)=>{const i=new Image();i.crossOrigin='anonymous';i.onload=()=>res(i);i.onerror=()=>rej(new Error('Logo tidak dapat dimuat untuk ikon'));i.src=PWA_ICON_URL});
  const mk=(size,pad)=>{const c=document.createElement('canvas');c.width=c.height=size;const x=c.getContext('2d');x.fillStyle='#ffffff';x.fillRect(0,0,size,size);
    const box=size*(1-2*pad),r=Math.min(box/img.naturalWidth,box/img.naturalHeight),w=img.naturalWidth*r,h=img.naturalHeight*r;
    x.imageSmoothingEnabled=true;x.imageSmoothingQuality='high';x.drawImage(img,(size-w)/2,(size-h)/2,w,h);return c.toDataURL('image/png')};
  return{i192:mk(192,.1),i512:mk(512,.1),m512:mk(512,.2),a180:mk(180,.08)}
}
function pwaApply(ic){
  const origin=location.origin,start=origin+location.pathname;
  const man={id:start,name:APP_NAME+' - Sistem Presensi Pegawai',short_name:'FIWA Presence',description:'Sistem presensi pegawai Fitrah Islamic World Academy',lang:'id',dir:'ltr',start_url:start,scope:origin+'/',display:'standalone',orientation:'portrait',background_color:'#ffffff',theme_color:'#2563eb',
    icons:[{src:ic.i192,sizes:'192x192',type:'image/png',purpose:'any'},{src:ic.i512,sizes:'512x512',type:'image/png',purpose:'any'},{src:ic.m512,sizes:'512x512',type:'image/png',purpose:'maskable'}]};
  document.querySelectorAll('link[rel=manifest]').forEach(x=>x.remove());
  const l=document.createElement('link');l.rel='manifest';l.href=URL.createObjectURL(new Blob([JSON.stringify(man)],{type:'application/manifest+json'}));document.head.appendChild(l);
  document.querySelectorAll('link[rel=apple-touch-icon]').forEach(x=>x.remove());
  const a=document.createElement('link');a.rel='apple-touch-icon';a.setAttribute('sizes','180x180');a.href=ic.a180;document.head.appendChild(a);
}
function setupPWA(){
  if(!document.querySelector('link[rel=manifest][data-static]'))pwaIcons().then(pwaApply).catch(e=>console.info('Ikon PWA tidak dapat dibuat dari logo:',e&&e.message));
  try{if(PWA_SW_URL&&'serviceWorker' in navigator&&/^https?:$/.test(location.protocol))navigator.serviceWorker.register(PWA_SW_URL).catch(err=>console.info('Service Worker gagal didaftarkan:',err&&err.message))}catch(e){}
  window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();S.installEvt=e});
  window.addEventListener('appinstalled',()=>{S.installEvt=null;$$('[data-act=install]').forEach(b=>b.classList.add('sk-hide'));try{toast('Aplikasi FIWA Presence terpasang','ok')}catch(x){}});
}
ACT.install=async()=>{
  if(S.installEvt){S.installEvt.prompt();try{await S.installEvt.userChoice}catch(e){}S.installEvt=null;return}
  if(pwaStandalone()){toast('Aplikasi sudah terpasang dan sedang dibuka sebagai aplikasi','ok');return}
  const steps=pwaIOS()
    ?'<ol class="sk-hint" style="padding-left:18px;line-height:1.7"><li>Buka halaman ini di <b>Safari</b>.</li><li>Ketuk tombol <b>Bagikan</b> (kotak dengan panah ke atas).</li><li>Pilih <b>Tambah ke Layar Utama</b>, lalu ketuk <b>Tambah</b>.</li></ol>'
    :'<ol class="sk-hint" style="padding-left:18px;line-height:1.7"><li>Buka halaman ini di <b>Chrome</b> (bukan di dalam aplikasi lain).</li><li>Ketuk menu <b>&#8942;</b> di pojok kanan atas.</li><li>Pilih <b>Instal aplikasi</b> atau <b>Tambahkan ke layar utama</b>.</li></ol>';
  modal({title:'Pasang FIWA Presence di HP',body:'<div class="sk-sub" style="margin-bottom:8px">Ikon FIWA akan muncul di layar utama dan aplikasi terbuka layar penuh.</div>'+steps,actions:[{t:'Mengerti',c:'sec',fn:()=>{}}]});
};

/* =====================================================================
   10. AUTENTIKASI
   ===================================================================== */
const ssGet=k=>{try{return sessionStorage.getItem(k)}catch(e){return null}};
const ssSet=(k,v)=>{try{v==null?sessionStorage.removeItem(k):sessionStorage.setItem(k,v)}catch(e){}};
async function adminEnter(email,pw){
  if(IS_PREVIEW){PV.init();const p=PV.db.admin_profiles.find(a=>a.email===String(email).trim().toLowerCase()&&a.pw===pw&&a.is_active);if(!p)throw new Error('Email atau kata sandi salah');PV.adm=p;return{id:p.id,full_name:p.full_name,role:p.role,employee_id:p.employee_id}}
  const sb=await ensureSB();const{error}=await sb.auth.signInWithPassword({email:String(email).trim(),password:pw});
  if(error)throw new Error(/invalid/i.test(error.message)?'Email atau kata sandi salah':error.message);
  return adminProfile();
}
async function adminProfile(){
  let w=await rpc('adm_whoami');
  if(!w){
    const name=await promptBox('Pendaftaran Admin Pertama','Akun ini belum terdaftar sebagai admin. Jika ini adalah akun pertama (belum ada admin sama sekali), isi nama lengkap Anda untuk menjadi Super Admin.',{multi:false,ok:'Daftarkan'});
    if(!name){await S.sb.auth.signOut();throw new Error('Akun belum terdaftar sebagai admin. Minta Super Admin menambahkan.')}
    try{w=await rpc('adm_claim_first_admin',{p_name:name})}catch(e){await S.sb.auth.signOut();throw e}
  }
  if(!w)throw new Error('Akun belum terdaftar sebagai admin');return w;
}
async function startAdmin(w){
  S.mode='admin';S.admin=w;S.token=null;S.me=null;ssSet('sk_tok',null);
  const t=TODAY();S.pm=null;S.geo=null;S.ym=D.mStart(t);
  await Ctx.month(S.ym,true);
  if(isMgr())await autoSetupDatabase().catch(e=>console.warn('autoSetup',e));
  processPending().then(n=>{if(n&&S.mode==='admin'&&S.route==='dashboard'){Ctx.reset();render()}});
  enterApp();
}
async function startUser(r){
  S.mode='user';S.admin=null;S.token=r.token;S.me=r.me;ssSet('sk_tok',r.token);S.pm=null;S.geo=null;S.ym=D.mStart(TODAY());
  await Ctx.month(S.ym,true);enterApp();
  if(S.me.must_change_password)openChangePassword(true);
}
function showLogin(){
  $('#sk-app').classList.add('sk-hide');const L=$('#sk-login');L.classList.remove('sk-hide');
  const dr=(k,badge,cls,lbl,u,p)=>'<button type="button" class="dr" data-tab="'+k+'" data-u="'+u+'" data-p="'+p+'"><span class="rb '+cls+'">'+badge+'</span><span class="dm"><b>'+lbl+'</b><code>'+u+'</code> &middot; <code>'+p+'</code></span></button>';
  const demo=IS_PREVIEW?'<details class="sk-demo"><summary><span class="di">'+ic('key','sm')+'</span><span class="dt"><b>Akun Pengujian Demo</b><small>Klik salah satu akun untuk mengisi otomatis</small></span></summary><div class="db">'
    +dr('admin','ADMIN','a','Super Admin','admin@fiwa.sch.id','admin123')+dr('admin','HRD','h','HRD Manager','hrd@fiwa.sch.id','hrd123')+dr('admin','STAF','h','HRD Staff','staf@fiwa.sch.id','staf123')
    +dr('user','GURU','p','Guru','budi.santoso@fiwa.sch.id','demo123')+dr('user','KADEPT','p','Manager / Kepala Dept.','siti.rahmawati@fiwa.sch.id','demo123')+dr('user','SHIFT','p','Petugas Keamanan','rudi.hartono@fiwa.sch.id','demo123')
    +'<div class="dn">Data contoh lokal (mode Preview). Klik untuk uji coba instan.</div></div></details>':'';
  const feat=(i,t,d)=>'<li><i>'+ic(i,'sm')+'</i><div><b>'+t+'</b><span>'+d+'</span></div></li>';
  L.innerHTML='<aside class="sk-lbrand"><div class="bn"><div class="sk-logo">'+logoImg()+'</div><div class="bn-text"><b>'+esc(APP_NAME)+'</b><small>'+esc(APP_TAGLINE)+'</small></div></div>'
   +'<div class="mid"><div class="sk-badge-pill"><i></i> Bismillah, Mulai Hari dengan Amanah</div><h2>Hadir dengan Amanah. <span>Berkhidmat dengan Ikhlas.</span></h2><p>Portal kehadiran keluarga besar Fitrah Islamic World Academy &mdash; Islamic International Boarding School. Membantu para ustadz, ustadzah, dan seluruh tim mencatat presensi, mengelola cuti, serta menata jadwal dengan tertib, transparan, dan penuh keberkahan.</p>'
   +'<ul>'+feat('clock','Presensi Tepat Waktu, Amanah Terjaga','Fingerprint &amp; check-in mandiri dengan validasi lokasi &mdash; karena disiplin adalah bagian dari akhlak mulia')+feat('file','Cuti &amp; Izin yang Transparan','Persetujuan bertingkat dengan kuota cuti terhitung otomatis, agar setiap urusan tertata dengan adil')+feat('calendar','Jadwal &amp; Piket Boarding','Atur jadwal reguler, piket shift, dan tukar shift mandiri dengan mudah dan tertib')+'</ul></div>'
   +'<div class="ft">'+ic('shield','sm')+'<span>Data terlindungi enkripsi database Supabase &mdash; amanah dijaga dengan sepenuh hati</span></div></aside>'
   +'<div class="sk-mhero"><div class="mh-logo">'+logoImg()+'</div><b>'+esc(APP_NAME)+'</b><small>'+esc(APP_TAGLINE)+'</small><div class="mh-pill">Bismillah, Mulai Hari dengan Amanah</div><button type="button" class="mh-install'+(pwaStandalone()?' sk-hide':'')+'" data-act="install" aria-label="Pasang aplikasi FIWA Presence di HP">'+ic('download','sm')+'<span>Pasang App</span></button></div>'
   +'<div class="sk-lside"><div class="sk-login-card"><div class="sk-login-card-head"><h1>Assalamu&rsquo;alaikum, Selamat Datang</h1><div class="sub" id="sk-lsub">Silakan pilih peran dan masukkan akun Anda untuk memulai hari dengan basmalah.</div></div>'
   +'<div class="sk-tabs" role="tablist"><button type="button" class="on" data-act="ltab" data-k="user">'+ic('user','sm')+'Pegawai</button><button type="button" data-act="ltab" data-k="admin">'+ic('shield','sm')+'Admin / HRD</button></div>'
   +'<form data-form="login-user" id="lf-user">'
   +fld('Email Pegawai',inp('email','',{type:'email',ph:'contoh: nama@fiwa.sch.id',req:1,attr:'autocomplete="username" autocapitalize="none" inputmode="email"'}))
   +fld('Kata Sandi',inp('password','',{type:'password',req:1,attr:'autocomplete="current-password"'}))
   +'<div class="sk-auth-options">'
   +'<label class="sk-rem-label"><input type="checkbox" name="remember" id="rem-user"><span>Simpan akses login</span></label>'
   +'<button type="button" class="sk-forgot-btn" data-act="openForgot" data-type="user">Lupa kata sandi?</button>'
   +'</div>'
   +'<button class="sk-b" type="submit">'+ic('check','sm')+'Masuk ke Akun Pegawai</button></form>'
   +'<form data-form="login-admin" id="lf-admin" class="sk-hide">'
   +fld('Email Akun Admin',inp('email','',{type:'email',ph:'contoh: admin@fiwa.sch.id',req:1,attr:'autocomplete="username"'}))
   +fld('Kata Sandi Admin',inp('password','',{type:'password',req:1,attr:'autocomplete="current-password"'}))
   +'<div class="sk-auth-options">'
   +'<label class="sk-rem-label"><input type="checkbox" name="remember" id="rem-admin"><span>Simpan akses login</span></label>'
   +'<button type="button" class="sk-forgot-btn" data-act="openForgot" data-type="admin">Lupa kata sandi?</button>'
   +'</div>'
   +'<button class="sk-b" type="submit">'+ic('shield','sm')+'Masuk Portal Admin</button></form>'
   +'<div id="lf-err" role="alert" aria-live="polite"></div>'+demo+'</div></div>';
  const MAIL='<svg class="sk-ic" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/></svg>';
  const EYE='<svg class="sk-ic" viewBox="0 0 24 24" aria-hidden="true"><path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>';
  const EYEOFF='<svg class="sk-ic" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 3l18 18"/><path d="M6.5 6.9C3.8 8.7 2 12 2 12s3.6 7 10 7c1.6 0 3-.4 4.3-1M10 5.2c.6-.1 1.3-.2 2-.2 6.4 0 10 7 10 7a17 17 0 0 1-3.2 4"/><path d="M9.9 10a3 3 0 0 0 4.2 4.2"/></svg>';
  $$('#sk-login .sk-f').forEach(lab=>{
    const i=$('input',lab);if(!i)return;
    const w=document.createElement('div');w.className='sk-iw';i.parentNode.insertBefore(w,i);
    const lead=document.createElement('span');lead.className='lead';lead.innerHTML=i.type==='password'?ic('lock'):(i.type==='email'?MAIL:ic('user'));w.appendChild(lead);w.appendChild(i);
    const fe=document.createElement('div');fe.className='sk-fe';fe.setAttribute('role','alert');lab.appendChild(fe);
    i.addEventListener('invalid',ev=>{ev.preventDefault();const vm=i.validity.valueMissing;fe.textContent=vm?(i.type==='password'?'Kata sandi wajib diisi':'Email wajib diisi'):'Masukkan email yang valid, contoh: nama@fiwa.sch.id';w.classList.add('bad');i.setAttribute('aria-invalid','true');const f=i.form;if(f&&!f._fi){f._fi=1;i.focus();setTimeout(()=>{f._fi=0},80)}});
    i.addEventListener('input',()=>{fe.textContent='';w.classList.remove('bad');i.removeAttribute('aria-invalid')});
    if(i.type==='password'){
      i.classList.add('hs');
      const tg=document.createElement('button');tg.type='button';tg.className='sk-pwb';tg.innerHTML=EYE;tg.setAttribute('aria-label','Tampilkan kata sandi');tg.setAttribute('aria-pressed','false');
      tg.addEventListener('click',()=>{const h=i.type==='password';i.type=h?'text':'password';tg.innerHTML=h?EYEOFF:EYE;tg.setAttribute('aria-label',h?'Sembunyikan kata sandi':'Tampilkan kata sandi');tg.setAttribute('aria-pressed',h?'true':'false')});
      w.appendChild(tg);
    }
  });

  // Pulihkan akun tersimpan jika sebelumnya memilih "Simpan akses login"
  try{
    const savedUser=localStorage.getItem('sk_saved_user');
    if(savedUser){
      const parsed=JSON.parse(savedUser);
      const fu=$('#lf-user');
      if(fu&&parsed.e){
        $('input[name=email]',fu).value=parsed.e;
        if(parsed.p)$('input[name=password]',fu).value=parsed.p;
        const cb=$('#rem-user');if(cb)cb.checked=true;
      }
    }
    const savedAdmin=localStorage.getItem('sk_saved_admin');
    if(savedAdmin){
      const parsed=JSON.parse(savedAdmin);
      const fa=$('#lf-admin');
      if(fa&&parsed.e){
        $('input[name=email]',fa).value=parsed.e;
        if(parsed.p)$('input[name=password]',fa).value=parsed.p;
        const cb=$('#rem-admin');if(cb)cb.checked=true;
      }
    }
  }catch(e){console.warn('load saved auth',e)}
  $$('#sk-login .dr').forEach(b=>b.addEventListener('click',()=>{
    const k=b.dataset.tab,tb=$('.sk-tabs button[data-k="'+k+'"]');if(tb&&!tb.classList.contains('on'))ACT.ltab(tb);
    const f=$(k==='admin'?'#lf-admin':'#lf-user'),u=$('input[name=email]',f),pw=$('input[name=password]',f);
    u.value=b.dataset.u;pw.value=b.dataset.p;[u,pw].forEach(x=>x.dispatchEvent(new Event('input')));$('button[type=submit]',f).focus();
  }));
}
ACT.tblMore=t=>tblMore(t);
ACT.ltab=t=>{$$('.sk-tabs button').forEach(b=>b.classList.toggle('on',b===t));$('#lf-user').classList.toggle('sk-hide',t.dataset.k!=='user');$('#lf-admin').classList.toggle('sk-hide',t.dataset.k!=='admin');$('#lf-err').innerHTML=''};
const loginErr=m=>{const e=$('#lf-err');if(e)e.innerHTML='<div class="sk-err">'+esc(m)+'</div>'};
ACT.openForgot=t=>{
  const isAdmin=t.dataset.type==='admin';
  modal({
    title:'Bantuan Lupa Kata Sandi',
    body:'<div style="display:flex;flex-direction:column;gap:14px;line-height:1.6">'
      +'<div class="sk-info" style="margin:0"><b>Keamanan Akun:</b> Kata sandi dienkripsi dengan standar database untuk perlindungan data presensi lembaga.</div>'
      +(isAdmin
        ?'<p>Untuk mereset kata sandi akun Admin atau HRD, silakan hubungi <b>Super Admin lembaga</b>. Super Admin dapat mengatur ulang kata sandi atau mengelola profil admin melalui menu Akun Pengguna.</p>'
        :'<p>Untuk mereset kata sandi akun Pegawai, silakan hubungi <b>HRD Staff atau Tata Usaha lembaga</b>. Pihak HRD/Admin dapat mereset kata sandi Anda dengan opsi "Reset Sandi" sementara, dan Anda akan diminta membuat kata sandi baru saat pertama kali masuk kembali.</p>')
      +'<div style="background:#f8fafc;padding:12px 14px;border-radius:12px;border:1px solid #e2e8f0;font-size:calc(13px*var(--fs,1));color:#475569">'
      +'<b style="color:#0f172a;display:block;margin-bottom:4px">Tips:</b>'
      +'Pastikan email yang dimasukkan sama dengan email yang terdaftar di HRD, dan Caps Lock sedang mati saat mengetik kata sandi.'
      +'</div>'
      +'</div>',
    actions:[{t:'Tutup',c:'sec',fn:()=>{}}]
  });
};

FORM['login-user']=async(f,d)=>{
  try{
    const r=await rpc('u_login',{p_email:String(d.email||'').trim(),p_password:d.password});
    if(!r.ok){loginErr(r.message);return}
    try{
      if(d.remember){localStorage.setItem('sk_saved_user',JSON.stringify({e:String(d.email||'').trim(),p:d.password}))}
      else{localStorage.removeItem('sk_saved_user')}
    }catch(e){}
    if(IS_PREVIEW)try{sessionStorage.setItem('sk_pv',JSON.stringify({m:'user',e:String(d.email||'').trim(),p:d.password}))}catch(e){}
    await startUser(r);
  }catch(e){loginErr(errMsg(e))}
};

FORM['login-admin']=async(f,d)=>{
  try{
    const w=await adminEnter(d.email,d.password);
    try{
      if(d.remember){localStorage.setItem('sk_saved_admin',JSON.stringify({e:d.email,p:d.password}))}
      else{localStorage.removeItem('sk_saved_admin')}
    }catch(e){}
    if(IS_PREVIEW)try{sessionStorage.setItem('sk_pv',JSON.stringify({m:'admin',e:String(d.email||'').trim(),p:d.password}))}catch(e){}
    await startAdmin(w);
  }catch(e){loginErr(errMsg(e))}
};
async function doLogout(silent){
  try{if(S.mode==='user'&&S.token)await rpc('u_logout',{p_token:S.token}).catch(()=>{});if(S.mode==='admin'){PV.adm=null;if(S.sb)await S.sb.auth.signOut().catch(()=>{})}}catch(e){}
  ssSet('sk_tok',null);try{sessionStorage.removeItem('sk_pv')}catch(e){}clearTimers();Object.assign(S,{mode:null,admin:null,me:null,token:null,ctx:null,route:'',cache:{}});
  if(location.hash)history.replaceState(null,'',location.pathname+location.search);
  showLogin();
}
ACT.logout=async()=>{if(await confirmBox('Keluar dari aplikasi?',{ok:'Keluar'}))await doLogout()};

async function tryRestore(){
  if(IS_PREVIEW){
    let pv=null;try{pv=JSON.parse(sessionStorage.getItem('sk_pv')||'null')}catch(e){}
    if(!pv)return false;
    try{
      if(pv.m==='admin'){await startAdmin(await adminEnter(pv.e,pv.p));return true}
      const r=await rpc('u_login',{p_email:pv.e,p_password:pv.p});if(r&&r.ok){await startUser(r);return true}
    }catch(e){}
    try{sessionStorage.removeItem('sk_pv')}catch(e){}
    return false;
  }
  const tok=ssGet('sk_tok');
  if(tok){
    try{
      S.token=tok;S.mode='user';
      const c=await rpc('u_context',{p_token:tok,p_from:D.mStart(TODAY()),p_to:D.mEnd(TODAY())});
      S.me=c.me;await startUserRestore();return true;
    }catch(e){
      S.token=null;S.mode=null;
      /* token dihapus hanya bila server menyatakan sesi tidak valid; gangguan jaringan tidak memutus login */
      if(/SESI_TIDAK_VALID/.test((e&&e.message)||''))ssSet('sk_tok',null);else console.warn('Pemulihan sesi pengguna gagal',e);
    }
  }
  try{
    const sb=await ensureSB(),{data}=await sb.auth.getSession();
    if(data&&data.session){const w=await rpc('adm_whoami');if(w){await startAdmin(w);return true}}
  }catch(e){console.info('restore admin',e&&e.message)}
  return false;
}
async function startUserRestore(){S.pm=null;S.geo=null;S.ym=D.mStart(TODAY());await Ctx.month(S.ym,true);enterApp()}

/* ---------- ganti sandi ---------- */
function openChangePassword(force){
  modal({title:force?'Ganti Kata Sandi Awal':'Ganti Kata Sandi',lock:!!force,body:(force?'<div class="sk-warn">Demi keamanan, ganti kata sandi awal Anda sebelum melanjutkan.</div>':'')
    +'<form data-form="chpw" id="chpw">'+fld('Kata sandi lama',inp('old','',{type:'password',req:1}))+fld('Kata sandi baru',inp('nw','',{type:'password',req:1}),'Minimal 6 karakter, tidak boleh sama dengan email atau NIP')+fld('Ulangi kata sandi baru',inp('nw2','',{type:'password',req:1}))+'<button type="submit" class="sk-hide"></button></form>',
    actions:[...(force?[]:[{t:'Batal',c:'sec',fn:()=>{}}]),{t:'Simpan',i:'check',keep:true,fn:async m=>{
      const f=$('#chpw',m.el),d=formData(f);if(d.nw!==d.nw2){toast('Kata sandi baru dan pengulangannya tidak sama','warn');return false}
      const r=await rpc('u_change_password',{p_token:S.token,p_old:d.old,p_new:d.nw});if(!r.ok){toast(r.message,'bad');return false}
      toast('Kata sandi berhasil diganti','ok');S.me.must_change_password=false;m.close()}}]});
}
FORM.chpw=()=>{};

/* =====================================================================
   11. SHELL, MENU, ROUTER
   ===================================================================== */
const clearTimers=()=>{(S.timers||[]).forEach(clearInterval);S.timers=[]};
function adminMenu(){return[
  ['Utama',[['dashboard','Dashboard','home'],['rekap','Rekap Kehadiran','list'],['persetujuan','Persetujuan','check']]],
  ['Data Master',[['pegawai','Pegawai & Departemen','users'],['akun','Akun Pengguna','key'],['jadwal','Jadwal Normal','clock'],['shift','Jadwal Shift','swap'],['libur','Kalender Libur','calendar']]],
  ['Kehadiran',[['impor','Impor Fingerprint','finger'],['saldo','Saldo Cuti','brief'],['laporan','Laporan','chart'],['periode','Penutupan Periode','lock']]],
  ['Sistem',[['pengaturan','Pengaturan','settings'],...(isMgr()?[['audit','Log Audit','history']]:[])]]
]}
function userMenu(){
  const g=[['Saya',[['beranda','Beranda & Absen','home'],['riwayat','Riwayat Kehadiran','list'],['jadwal-saya','Jadwal Saya','calendar'],['ajukan','Ajukan','plus'],['pengajuan','Pengajuan Saya','file'],['tukar','Tukar Shift','swap']]]];
  const k=[];
  if(isHead())k.push(['persetujuan-l1','Kotak Persetujuan','inbox'],['dept','Kehadiran Departemen','users']);
  if(isHead())k.push(['laporan-u','Laporan Departemen','chart']);
  if(k.length)g.push(['Kelola',k]);
  g.push(['Akun',[['notifikasi','Notifikasi','bell'],['profil','Profil & Sandi','user']]]);return g;
}
function badges(){
  const c=S.ctx||{},b={};
  if(S.mode==='admin'){b.persetujuan=(c.requests||[]).filter(r=>r.status==='waiting_l2'||(r.status==='approved'&&r.cancel_requested)).length;b.shift=(c.swaps||[]).filter(s=>s.status==='waiting_coordinator').length}
  else if(S.me){
    const hd=S.me.head_dept_ids||[];
    b['persetujuan-l1']=(c.requests||[]).filter(r=>r.status==='waiting_l1'&&r.employee_id!==S.me.id&&hd.includes((empIx()[r.employee_id]||{}).department_id)).length;
    b.tukar=(c.swaps||[]).filter(s=>s.partner_employee_id===S.me.id&&s.status==='waiting_partner').length;
  }
  b.notif=(c.notifications||[]).filter(n=>!n.is_read).length;b.notifikasi=b.notif;return b;
}
function buildNav(){
  const groups=S.mode==='admin'?adminMenu():userMenu(),bg=badges();S.menus=groups.flatMap(g=>g[1]);S.groups=groups;
  $('#sk-nav').innerHTML=groups.map(g=>'<div class="sec">'+esc(g[0])+'</div>'+g[1].map(m=>'<a data-act="nav" data-k="'+m[0]+'" title="'+esc(m[1])+'" class="'+(S.route===m[0]?'on':'')+'"'+(S.route===m[0]?' aria-current="page"':'')+'>'+ic(m[2])+'<span style="flex:1">'+esc(m[1])+'</span>'+dot(bg[m[0]]||0)+'</a>').join('')).join('');
  const main=S.mode==='admin'?['dashboard','pegawai','persetujuan','laporan']:['beranda','riwayat','ajukan','pengajuan'];
  $('#sk-bnav').innerHTML=main.map(k=>{const m=S.menus.find(x=>x[0]===k);return m?'<a data-act="nav" data-k="'+k+'" class="'+(S.route===k?'on':'')+'">'+ic(m[2])+'<span>'+esc(m[1].split(' ')[0])+'</span>'+dot(bg[k]||0)+'</a>':''}).join('')+'<a data-act="more" class="'+(main.includes(S.route)?'':'on')+'">'+ic('grid')+'<span>Menu</span>'+dot(Object.entries(bg).filter(([k])=>!main.includes(k)&&k!=='notif'&&!['dashboard'].includes(k)).reduce((s,[k,v])=>s+v,0))+'</a>';
  const bell=$('#sk-bell');bell.innerHTML=ic('bell')+dot(bg.notif);
}
ACT.sideToggle=()=>{const a=$('#sk-app');a.classList.toggle('sk-coll');try{localStorage.setItem('sk.coll',a.classList.contains('sk-coll')?'1':'0')}catch(e){}};
ACT.nav=t=>{const m=$('.sk-sheet');if(m)m.remove();go(t.dataset.k)};
ACT.more=()=>{
  const old=$('#sk-sheet');if(old){old.remove();return}
  const bg=badges(),s=document.createElement('div');s.id='sk-sheet';s.className='sk-ov';
  const PAL=['#0f766e','#2563eb','#7c3aed','#d97706','#be123c','#0369a1'];
  s.innerHTML='<div class="sk-sheet"><div class="sk-row sp" style="margin-bottom:6px"><b>Semua Menu</b><button class="sk-ibtn" data-act="closeSheet" aria-label="Tutup">'+ic('x')+'</button></div>'+(S.groups||[['',S.menus]]).map((g,i)=>'<div class="sk-sg">'+esc(g[0])+'</div><div class="gr" style="--c:'+PAL[i%PAL.length]+'">'+g[1].map(m=>'<a data-act="navs" data-k="'+m[0]+'" class="'+(S.route===m[0]?'cur':'')+'">'+ic(m[2])+'<span>'+esc(m[1])+'</span>'+dot(bg[m[0]]||0)+'</a>').join('')+'</div>').join('')+'<div class="sk-sg">Akun</div><div class="gr"><a data-act="logout">'+ic('logout')+'<span>Keluar</span></a></div></div>';
  s.addEventListener('mousedown',e=>{if(e.target===s)s.remove()});$('#sk-root').appendChild(s);
};
ACT.closeSheet=()=>{const s=$('#sk-sheet');if(s)s.remove()};
ACT.navs=t=>{const s=$('#sk-sheet');if(s)s.remove();go(t.dataset.k)};
ACT.goNotif=()=>{if(S.mode==='user')go('notifikasi');else openAdminNotifs()};
const go=k=>{if(location.hash==='#/'+k)render();else location.hash='#/'+k};
function enterApp(){
  $('#sk-login').classList.add('sk-hide');$('#sk-app').classList.remove('sk-hide');
  const who=S.mode==='admin'?S.admin:S.me;$('#sk-un').textContent=who.full_name;$('#sk-av').textContent=initials(who.full_name);$('#sk-tav').textContent=initials(who.full_name);$('#sk-tun').textContent=who.full_name;
  $('#sk-ur').textContent=S.mode==='admin'?ROLE[S.admin.role]:(S.me.roles.filter(r=>r!=='pegawai').map(r=>ROLE[r]).join(', ')||'Pegawai');
  $('#sk-tur').textContent=$('#sk-ur').textContent;$$('[data-act=logout]').forEach(b=>b.innerHTML=ic('logout'));
  try{if(localStorage.getItem('sk.coll')==='1')$('#sk-app').classList.add('sk-coll')}catch(e){}
  const cb=$('#sk-coll');if(cb)cb.innerHTML=ic('menu');
  const def=S.mode==='admin'?'dashboard':'beranda';const h=(location.hash||'').replace(/^#\//,'');
  S.route='';if(h&&(S.mode==='admin'?adminMenu():userMenu()).some(g=>g[1].some(m=>m[0]===h))&&VIEWS[h])render();else go(def);
  const lg=$('#sk-logo');if(lg&&!lg.innerHTML)lg.innerHTML=logoImg();
}
async function render(force){
  if(!S.mode)return;
  const h=(location.hash||'').replace(/^#\//,'');const all=(S.mode==='admin'?adminMenu():userMenu()).flatMap(g=>g[1]);
  let k=all.some(m=>m[0]===h)?h:(S.mode==='admin'?'dashboard':'beranda');
  const my=++RT;clearTimers();S.route=k;const v=VIEWS[k],el=$('#sk-view');
  $('#sk-title').textContent=(all.find(m=>m[0]===k)||[])[1]||'';
  {const g=(S.mode==='admin'?adminMenu():userMenu()).find(x=>x[1].some(m=>m[0]===k));const cr=$('#sk-crumb');if(cr)cr.textContent=g?g[0]:'';const dt=$('#sk-date');if(dt)dt.textContent=fmtDL(TODAY())}
  if(!v){el.innerHTML='<div class="sk-err">Halaman tidak tersedia</div>';return}
  el.innerHTML=skeleton();$('.sk-main').scrollTop=0;el.scrollTop=0;window.scrollTo&&0;
  try{
    if(v.needs!==false){const c=v.range?v.range():[ymStart(),D.mEnd(ymStart())];await Ctx.load(c[0],c[1],force)}
    if(my!==RT)return;
    const html=await v.render(el);if(my!==RT)return;
    if(typeof html==='string')el.innerHTML=html;
    if(v.after)v.after(el);
    buildNav();
  }catch(e){if(my!==RT)return;console.error(e);if(/SESI_TIDAK_VALID/.test(e.message||'')){fail(e);return}el.innerHTML='<div class="sk-err">'+esc(errMsg(e))+'</div><button class="sk-b sec" data-act="retry">'+ic('refresh','sm')+'Coba lagi</button>'}
}
ACT.retry=()=>render(true);
/* tarik layar ke bawah untuk memuat ulang data (HP) */
(function(){
  let y0=0,d=0,on=false,busy=false;const TH=55;
  const ind=()=>{let e=$('#sk-ptr');if(!e){e=document.createElement('div');e.id='sk-ptr';e.setAttribute('aria-hidden','true');e.innerHTML=ic('refresh');$('#sk-root').appendChild(e)}return e};
  const rest=e=>{e.classList.remove('spin');e.style.transform='translate(-50%,-48px)';e.style.opacity='0'};
  document.addEventListener('touchstart',ev=>{
    on=false;if(!S.mode||busy||ev.touches.length!==1||window.innerWidth>900)return;
    const v=$('#sk-view');if(!v||!v.contains(ev.target)||$('.sk-ov')||ev.target.closest('input,textarea,select,.sk-tw'))return;
    if(v.scrollTop>0)return;y0=ev.touches[0].clientY;d=0;on=true;
  },{passive:true});
  document.addEventListener('touchmove',ev=>{
    if(!on)return;const v=$('#sk-view'),e=ind();
    const dy=ev.touches[0].clientY-y0;if(dy<=0||v.scrollTop>0){on=false;rest(e);return}
    d=Math.min(dy*.5,90);e.style.transition='none';e.style.opacity=String(Math.min(1,d/40));
    e.style.transform='translate(-50%,'+(d-48)+'px) rotate('+(d*5)+'deg)';e.classList.toggle('rdy',d>=TH);
  },{passive:true});
  document.addEventListener('touchend',async()=>{
    if(!on)return;on=false;const e=ind();e.style.transition='';e.classList.remove('rdy');
    if(d<TH){rest(e);return}
    busy=true;e.classList.add('spin');e.style.opacity='1';e.style.transform='translate(-50%,'+(TH-24)+'px)';
    try{await render(true);toast('Data diperbarui')}catch(x){}
    busy=false;rest(e);
  },{passive:true});
  document.addEventListener('touchcancel',()=>{if(on&&!busy){on=false;rest(ind())}},{passive:true});
})();
window.addEventListener('hashchange',()=>{if(S.mode)render()});

/* ---------- notifikasi ---------- */
function notifList(list){
  return list.length?'<div class="sk-list">'+list.map(n=>'<div class="sk-li '+(n.is_read?'':'unread')+'"><div style="flex:1;min-width:0"><div class="t">'+esc(n.title)+'</div><div>'+esc(n.body||'')+'</div><div class="m">'+esc(fmtDT(n.created_at))+'</div></div>'+(n.is_read?'':'<button class="sk-b sec sm" data-act="readOne" data-id="'+n.id+'">Tandai dibaca</button>')+'</div>').join('')+'</div>':'<div class="sk-empty">Belum ada notifikasi</div>';
}
async function markRead(id){
  if(S.mode==='user')await rpc('u_mark_read',{p_token:S.token,p_id:id==null?null:id});
  else{const ids=(S.ctx.notifications||[]).filter(n=>!n.is_read&&(id==null||n.id===id)).map(n=>n.id);for(const i of ids)await DB.update('notifications',{is_read:true},{id:i})}
  (S.ctx.notifications||[]).forEach(n=>{if(id==null||n.id===id)n.is_read=true});buildNav();
}
ACT.readOne=async t=>{await markRead(+t.dataset.id);if(S.route==='notifikasi')$('#sk-view').innerHTML=VIEWS.notifikasi.render();else{const b=$('.sk-ov .b');if(b)b.innerHTML=notifList(S.ctx.notifications)}};
ACT.readAll=async()=>{await markRead(null);if(S.route==='notifikasi')$('#sk-view').innerHTML=VIEWS.notifikasi.render();else{const b=$('.sk-ov .b');if(b)b.innerHTML=notifList(S.ctx.notifications)}};
function openAdminNotifs(){modal({title:'Notifikasi HRD',size:'wide',body:notifList(S.ctx.notifications||[]),actions:[{t:'Tandai semua dibaca',c:'sec',i:'check',keep:true,fn:()=>{ACT.readAll();return false}},{t:'Tutup',c:'sec',fn:()=>{}}]})}
reg('notifikasi',{needs:true,render(){return card('Notifikasi',notifList(S.ctx.notifications||[]),'<button class="sk-b sec sm" data-act="readAll">'+ic('check','sm')+'Tandai semua dibaca</button>')}});

/* ---------- boot ---------- */
async function boot(){
  if(S.booted)return;S.booted=true;
  try{document.body.appendChild($('#sk-root'));document.documentElement.style.overflow='hidden';document.body.style.overflow='hidden'}catch(e){}
  setupPWA();
  $('#sk-logo').innerHTML=logoImg();$$('[data-act=logout]').forEach(b=>b.innerHTML=ic('logout'));
  $('#sk-bell').innerHTML=ic('bell');
  if(IS_PREVIEW)PV.init();
  let ok=false;try{ok=await tryRestore()}catch(e){}
  if(!ok)showLogin();
}


/* =====================================================================
   12. UTILITAS TABEL / IMPOR
   ===================================================================== */
function parseCSV(text){
  text=String(text).replace(/^﻿/,'');const first=text.split(/\r?\n/)[0]||'';
  const delim=[',',';','\t','|'].map(d=>[d,first.split(d).length]).sort((a,b)=>b[1]-a[1])[0][0];
  const rows=[];let row=[],f='',q=false;
  for(let i=0;i<text.length;i++){const c=text[i];
    if(q){if(c==='"'){if(text[i+1]==='"'){f+='"';i++}else q=false}else f+=c}
    else if(c==='"')q=true;
    else if(c===delim){row.push(f);f=''}
    else if(c==='\n'||c==='\r'){if(c==='\r'&&text[i+1]==='\n')i++;row.push(f);f='';if(row.some(x=>String(x).trim()!==''))rows.push(row);row=[]}
    else f+=c}
  row.push(f);if(row.some(x=>String(x).trim()!==''))rows.push(row);return rows;
}
async function readTable(file){
  const nm=file.name.toLowerCase();
  if(/\.(xlsx|xls)$/.test(nm)){
    await loadScript(XLSX_URL);const wb=XLSX.read(await fileRead(file),{type:'array',cellDates:false});
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:false,defval:''}).filter(r=>r.some(x=>String(x).trim()!==''));
  }
  return parseCSV(await file.text());
}
const normKey=s=>String(s||'').toLowerCase().replace(/[^a-z0-9]/g,'');
function summarize(emp,from,to){
  const o={sched:0,present:0,late:0,incomplete:0,leave:0,sick:0,permit:0,official_duty:0,absent:0,late_min:0,early_min:0,worked:0,ot:0};
  for(const ds of D.range(from,to)){
    const r=recFor(emp,ds),s=r.status;
    if(['present','late','incomplete','absent','leave','sick','permit','official_duty'].includes(s))o.sched++;
    if(o[s]!==undefined)o[s]++;
    o.late_min+=r.late_minutes||0;o.early_min+=r.early_leave_minutes||0;o.worked+=r.worked_minutes||0;o.ot+=r.overtime_minutes||0;
  }
  o.hadir=o.present+o.late;o.pct=o.sched?Math.round((o.hadir+o.official_duty)*100/o.sched):0;return o;
}
const tm=iso=>iso?fmtT(Date.parse(iso)):'-';
const EMPTY_STAT={present:0,late:0,incomplete:0,absent:0,leave:0,sick:0,permit:0,official_duty:0};
function barChart(items,max){
  max=max||Math.max(1,...items.map(i=>i.v));
  return'<div class="sk-bars">'+items.map(i=>'<div class="c"><small>'+esc(i.l2||'')+'</small><div class="bar" style="height:'+Math.max(2,Math.round(i.v*100/max))+'%;'+(i.c?'background:'+i.c:'')+'" title="'+esc(i.t||i.v)+'"></div><small>'+esc(i.l)+'</small></div>').join('')+'</div>';
}
function attachLink(u){return u?'<a href="'+esc(u)+'" target="_blank" rel="noopener noreferrer">Buka lampiran</a>':'-'}
function reqPeriod(r){return r.start_date===r.end_date?fmtD(r.start_date):fmtD(r.start_date)+' s.d. '+fmtD(r.end_date)}
function reqExtra(r){
  const rt=rtOf(r.request_type_id);let x='';
  if(r.start_time)x+=' '+D.hm(r.start_time)+(r.end_time?'-'+D.hm(r.end_time):'');
  if(r.day_part&&r.day_part!=='full')x+=' ('+(r.day_part==='morning'?'pagi':'siang')+')';
  if(rt.category==='correction'&&r.correction_payload){const c=r.correction_payload;x+=' masuk '+(c.in||'-')+' / pulang '+(c.out||'-')}
  return x;
}
function approvalTrail(id){
  const l=(S.ctx.approvals||[]).filter(a=>a.request_id===id).sort((a,b)=>a.id-b.id);
  if(!l.length)return'<div class="sk-sub">Belum ada keputusan.</div>';
  const lab={approved:'Disetujui',rejected:'Ditolak',cancel_approved:'Pembatalan disetujui',cancel_rejected:'Pembatalan ditolak'};
  return'<div class="sk-list">'+l.map(a=>'<div class="sk-li"><div><div class="t">Tahap '+a.stage+': '+esc(lab[a.decision]||a.decision)+'</div><div class="m">'+esc(a.approver_name||'')+' &middot; '+esc(fmtDT(a.decided_at))+'</div>'+(a.note?'<div>'+esc(a.note)+'</div>':'')+'</div></div>').join('')+'</div>';
}
function requestDetail(r,withBalance){
  const rt=rtOf(r.request_type_id);let b='';
  if(withBalance&&rt.deducts_balance){const bl=Engine.balances(S.ctx,r.employee_id,r.start_date).find(x=>x.type.id===rt.id);if(bl)b='<div class="sk-info">Saldo '+esc(rt.name)+' periode '+esc(bl.period)+': jatah '+nf(bl.total)+', terpakai '+nf(bl.used)+', menunggu '+nf(bl.reserved)+', sisa tersedia '+nf(bl.left)+' hari.</div>'}
  return'<div class="sk-grid sk-g2"><div><div class="sk-sub">Pegawai</div><b>'+esc(nameOf(r.employee_id))+'</b><div class="sk-sub">'+esc(deptOf(r.employee_id))+'</div></div><div><div class="sk-sub">Jenis</div><b>'+esc(rt.name)+'</b> '+rstBadge(r.status)+(r.cancel_requested?' '+bd('Minta batal','o'):'')+'</div></div>'
   +'<div class="sk-grid sk-g2" style="margin-top:10px"><div><div class="sk-sub">Tanggal</div>'+esc(reqPeriod(r))+esc(reqExtra(r))+'</div><div><div class="sk-sub">Jumlah hari</div>'+nf(r.total_days)+'</div></div>'
   +'<div style="margin-top:10px"><div class="sk-sub">Alasan</div>'+(r.reason?esc(r.reason):'<i>(disembunyikan)</i>')+'</div><div style="margin-top:6px"><div class="sk-sub">Lampiran</div>'+attachLink(r.attachment_url)+'</div>'
   +(r.cancel_requested?'<div class="sk-warn" style="margin-top:10px">Alasan pembatalan: '+esc(r.cancel_reason||'-')+'</div>':'')+b
   +'<h4 style="margin:14px 0 8px">Riwayat keputusan</h4>'+approvalTrail(r.id);
}

/* =====================================================================
   13. ADMIN: DASHBOARD
   ===================================================================== */
const dbRing=(p)=>{const R=38,C=2*Math.PI*R,v=p==null?0:Math.max(0,Math.min(100,p));return`<div class="db-ring" role="img" aria-label="Tingkat hadir hari ini ${p==null?'belum tersedia':v+' persen'}"><svg viewBox="0 0 100 100" aria-hidden="true"><circle cx="50" cy="50" r="${R}" fill="none" stroke="rgba(255,255,255,.2)" stroke-width="10"/><circle cx="50" cy="50" r="${R}" fill="none" stroke="#fff" stroke-width="10" stroke-linecap="round" stroke-dasharray="${(C*v/100).toFixed(1)} ${C.toFixed(1)}" transform="rotate(-90 50 50)"/></svg><div class="db-rv"><b>${p==null?'-':v+'%'}</b><small>Hadir hari ini</small></div></div>`};
const dbTrend=wd=>`<div class="db-tr">${wd.map((x,i)=>{const p=Math.round(x.v),lv=x.l2==='-'?'n':rkLv(p);return`<div class="db-tc${i===wd.length-1?' now':''}" title="${esc(x.t)}"><span class="db-tp">${esc(x.l2)}</span><div class="db-tbw"><div class="db-tb ${lv}" style="height:${Math.max(3,p)}%"></div></div><span class="db-td">${esc(x.l)}</span></div>`}).join('')}</div><div class="db-lg"><span><i class="g"></i>95% ke atas</span><span><i class="a"></i>85-94%</span><span><i class="r"></i>di bawah 85%</span><span class="sp"></span><span>Pegawai terjadwal yang hadir (termasuk dinas luar)</span></div>`;
reg('dashboard',{
  range:()=>[D.add(TODAY(),-7),TODAY()],
  render(){
    const c=S.ctx,t=TODAY(),act=c.employees.filter(e=>e.is_active),st={};
    const SCH=['present','late','incomplete','absent','leave','sick','permit','official_duty'],HDR=['present','late','official_duty'];
    let lateMin=0;
    act.forEach(e=>{const r=recFor(e.id,t);st[r.status]=(st[r.status]||0)+1;lateMin+=r.late_minutes||0});
    const wd=D.range(D.add(t,-6),t).map(ds=>{let s=0,h=0;act.forEach(e=>{const r=recFor(e.id,ds);if(SCH.includes(r.status)){s++;if(HDR.includes(r.status))h++}});return{l:dayName(ds),l2:s?Math.round(h*100/s)+'%':'-',v:s?h*100/s:0,t:fmtD(ds)+': '+h+' dari '+s}});
    const wl2=c.requests.filter(r=>r.status==='waiting_l2'),wl1=c.requests.filter(r=>r.status==='waiting_l1'),cr=c.requests.filter(r=>r.status==='approved'&&r.cancel_requested);
    const absent=act.filter(e=>{const s=recFor(e.id,t).status;return s==='absent'||s==='incomplete'}),onLeave=act.filter(e=>['leave','sick','permit','official_duty'].includes(recFor(e.id,t).status));
    const lastP=c.settings.last_processed_date,nSw=(c.swaps||[]).filter(s=>s.status==='waiting_coordinator').length;
    const nS=act.filter(e=>SCH.includes(recFor(e.id,t).status)).length,nH=act.filter(e=>HDR.includes(recFor(e.id,t).status)).length,pct=nS?Math.round(nH*100/nS):null;
    const soon=act.filter(e=>e.employment_status==='contract'&&e.contract_end_date&&daysTo(e.contract_end_date)<=30).length;
    const qlink=[['pegawai','users','Data pegawai','Pegawai & departemen'],['rekap','list','Rekap kehadiran','Data absensi'],['shift','swap','Jadwal & shift','Jadwal kerja'],['persetujuan','check','Persetujuan','Cuti, izin, sakit',wl2.length+cr.length],['laporan','chart','Laporan','Rekap periodik'],['pengaturan','settings','Pengaturan','Sistem']];
    S.dash=Object.assign({q:'',st:'',p:1},{per:S.dash&&S.dash.per!=null?S.dash.per:5});
    S._dashRows=act.map(e=>{const r=recFor(e.id,t);return{e,n:e.full_name,no:e.employee_no||'',dp:deptName(e.department_id),st:r.status,c:[rkWho(e),esc(deptName(e.department_id)),esc(schedText(e.id,t)),tm(r.check_in_at)+(r.late_minutes?` <span class="rk-n a">+${r.late_minutes}m</span>`:''),tm(r.check_out_at),stBadge(r.status)]}});
    const od=(n,h)=>'<div style="order:'+n+';min-width:0">'+h+'</div>';
    const perSel='<select class="sk-sel" aria-label="Jumlah baris ditampilkan" data-chg="dashPer">'+[[5,'Tampilkan 5'],[10,'Tampilkan 10'],[25,'Tampilkan 25'],[0,'Tampilkan semua']].map(o=>'<option value="'+o[0]+'"'+(S.dash.per===o[0]?' selected':'')+'>'+o[1]+'</option>').join('')+'</select>';
    const nm=String(S.admin&&S.admin.full_name||'').replace(/^((Prof|Dr|Dra|Drs|Ir|H|Hj)\.\s*)+/i,'').split(' ')[0];
    const pend=wl2.length+cr.length;
    const hero=`<div class="rk-top db-hero"><div class="rk-tl"><span class="rk-eye">${ic('home','sm')}Dashboard</span><div class="lp-h">${esc(greet())}${nm?', '+esc(nm):''}</div><div class="lp-sub">${esc(fmtDL(t))} · ${nS?nH+' dari '+nS+' pegawai terjadwal sudah hadir':'tidak ada pegawai terjadwal hari ini'}</div><div class="rk-meta"><span class="rk-mc">${ic('users','sm')}${act.length} pegawai aktif</span><span class="rk-mc">${ic('refresh','sm')}Rekap terakhir: ${lastP?esc(fmtD(lastP)):'belum pernah'}</span></div><div class="sk-rkact" style="margin-top:6px"><button class="sk-b ghost sm rk-pr" data-act="nav" data-k="persetujuan">${ic('check','sm')}<span class="lbl">Persetujuan</span>${pend?`<span class="db-bc">${pend}</span>`:''}</button><button class="sk-b sec sm" data-act="nav" data-k="rekap">${ic('list','sm')}<span class="lbl">Rekap Kehadiran</span></button><button class="sk-b sec sm" data-act="procNow">${ic('refresh','sm')}<span class="lbl">Proses rekap</span></button></div></div></div>`;
    const kp='<div class="sk-kpis rk-kpis">'+kpi('check','#059669',(st.present||0)+(st.late||0),'Hadir hari ini',(st.present||0)+' tepat waktu')+kpi('clock','#d97706',st.late||0,'Terlambat hari ini',fmtMin(lateMin)+' total')+kpi('alert','#dc2626',(st.absent||0)+(st.incomplete||0),'Alpha / belum lengkap',(st.incomplete||0)+' belum lengkap')
      +kpi('calendar','#7c3aed',onLeave.length,'Cuti, sakit, izin, dinas','Cuti '+(st.leave||0)+' · Sakit '+(st.sick||0)+' · Izin '+(st.permit||0)+' · Dinas '+(st.official_duty||0))+kpi('user','#64748b',st.pending||0,'Belum absen','Terjadwal, belum tercatat')+kpi('inbox','#0ea5e9',pend,'Menunggu keputusan HRD','','persetujuan')+'</div>';
    const mix=rkMix({present:st.present||0,late:st.late||0,official_duty:st.official_duty||0,leave:st.leave||0,sick:st.sick||0,permit:st.permit||0,incomplete:st.incomplete||0,absent:st.absent||0},'pegawai terjadwal hari ini');
    const todo=[['inbox','#2563eb','Menunggu persetujuan HRD',wl2.length+' pengajuan, '+cr.length+' permintaan pembatalan',pend,'persetujuan'],['users','#d97706','Menunggu Kepala Departemen',wl1.length+' pengajuan (HRD dapat mengambil alih)',wl1.length,'persetujuan'],['swap','#0ea5e9','Tukar shift menunggu',nSw+' permintaan',nSw,'shift'],['brief','#dc2626','Kontrak segera berakhir',soon+' pegawai dalam 30 hari',soon,'pegawai']];
    const tot=todo.reduce((a,x)=>a+x[4],0);
    const todoHtml=todo.map(x=>`<button type="button" class="db-td2${x[4]?' hot':''}" data-act="nav" data-k="${x[5]}"><span class="pn-ci" style="background:${x[1]}">${ic(x[0])}</span><span class="db-tx"><b>${esc(x[2])}</b><small>${esc(x[3])}</small></span><span class="db-n">${x[4]}</span>${ic('right','sm')}</button>`).join('');
    const avg=wd.filter(x=>x.l2!=='-'),avgP=avg.length?Math.round(avg.reduce((a,x)=>a+x.v,0)/avg.length):null;
    const byD={};act.forEach(e=>{const s=recFor(e.id,t).status;if(!SCH.includes(s))return;const k=e.department_id||0,o=byD[k]||(byD[k]={s:0,h:0});o.s++;if(HDR.includes(s))o.h++});
    const dRows=Object.keys(byD).map(k=>({n:+k?deptName(+k):'Tanpa departemen',s:byD[k].s,p:Math.round(byD[k].h*100/byD[k].s)})).sort((a,b)=>b.p-a.p);
    const dept=`<section class="rk-panel"><h4>${ic('building','sm')}Kehadiran per departemen hari ini</h4>${dRows.length?dRows.map(r=>`<div class="rk-dr"><div class="rk-drh"><span>${esc(r.n)} <small>${r.s} terjadwal</small></span><b class="${rkLv(r.p)}">${r.p}%</b></div><div class="rk-bar"><i class="${rkLv(r.p)}" style="width:${Math.min(100,r.p)}%"></i></div></div>`).join(''):'<div class="sk-empty">Belum ada pegawai terjadwal hari ini</div>'}</section>`;
    const abHtml=absent.length?absent.slice(0,8).map(e=>`<div class="db-ar">${rkWho(e)}<span class="db-ad">${esc(deptName(e.department_id))}</span>${stBadge(recFor(e.id,t).status)}</div>`).join('')+(absent.length>8?`<div class="db-more"><span class="sk-sub">dan ${absent.length-8} lainnya</span><button class="sk-b ghost sm" data-act="nav" data-k="rekap">Lihat di Rekap</button></div>`:''):`<div class="rk-ok">${ic('check','sm')}Tidak ada. Semua pegawai terjadwal sudah hadir atau belum waktunya.</div>`;
    const dAgo=lastP?D.diff(String(lastP).slice(0,10),t):null,fresh=lastP&&dAgo<=1;
    const prc=`<div class="db-pr"><span class="db-pi ${fresh?'ok':'warn'}">${ic(fresh?'check':'clock')}</span><div><b>${lastP?esc(fmtD(lastP)):'Belum pernah'}</b><small>${lastP?(fresh?'Rekap sudah terbaru':dAgo+' hari lalu, sebaiknya diproses ulang'):'Belum ada rekap yang diproses'}</small></div></div><div class="sk-sub" style="margin:10px 0">Rekap diproses otomatis saat admin masuk dan setiap keputusan persetujuan.</div><button class="sk-b" data-act="procNow">${ic('refresh','sm')}Proses rekap 7 hari terakhir</button>`;
    return hero+kp+mix
      +'<div class="sk-dg"><div class="sk-stack">'
      +od(2,card('Kehadiran hari ini &middot; '+esc(fmtD(t)),'<div id="dash-tbl"></div>','<div class="sk-tool"><span id="dash-ef" class="ef-inl">'+efBar('dash',act,{inl:1,nolim:1})+'</span><input class="sk-in" type="search" placeholder="Cari nama, NIP, departemen" aria-label="Cari pegawai" data-inp="dashQ"><select class="sk-sel" aria-label="Filter status" data-chg="dashSt"><option value="">Semua status</option>'+Object.keys(ST).map(k=>'<option value="'+k+'">'+esc(ST[k][0])+'</option>').join('')+'</select>'+perSel+'</div>'))
      +od(3,card('Tingkat kehadiran 7 hari terakhir',dbTrend(wd),avgP==null?'':'<span class="rk-cnt">Rata-rata '+avgP+'%</span>'))
      +od(4,card('Tidak hadir / belum lengkap hari ini <span class="rk-cnt">'+absent.length+'</span>',abHtml))
      +od(5,dept)
      +'</div><div class="sk-stack">'
      +od(1,card('Perlu tindakan'+(tot?' <span class="rk-cnt">'+tot+'</span>':''),(tot?'':`<div class="rk-ok" style="margin-bottom:10px">${ic('check','sm')}Semua beres, tidak ada yang menunggu tindakan.</div>`)+todoHtml))
      +od(6,card('Pengelolaan operasional','<div class="db-ql">'+qlink.map(q=>'<a class="db-q" data-act="nav" data-k="'+q[0]+'" role="button" tabindex="0"><span class="db-qi">'+ic(q[1])+'</span><b>'+q[2]+'</b><small>'+q[3]+'</small>'+dot(q[4]||0)+'</a>').join('')+'</div>'))
      +od(7,card('Pemrosesan rekap',prc))
      +od(8,card('Notifikasi HRD',notifList((c.notifications||[]).slice(0,5)),'<button class="sk-b sec sm" data-act="readAll">Tandai dibaca</button>'))
      +'</div></div>';
  },
  after(){drawDash()}
});
function drawDash(){
  const el=$('#dash-tbl');if(!el)return;const d=S.dash,q=d.q.trim().toLowerCase();
  const rows=(S._dashRows||[]).filter(r=>(!d.st||r.st===d.st)&&efPass('dash',r.e)&&(!q||(r.n+' '+r.no+' '+r.dp).toLowerCase().includes(q)));
  const all=!d.per,per=all?Math.max(rows.length,1):d.per,pages=Math.max(1,Math.ceil(rows.length/per));if(d.p>pages)d.p=pages;
  const from=(d.p-1)*per,part=rows.slice(from,from+per);
  el.innerHTML=tbl(['Pegawai','Departemen','Jadwal','Masuk','Pulang','Status'],part.map(r=>r.c),{empty:'Tidak ada pegawai yang sesuai dengan pencarian atau filter',maxh:all&&rows.length>10?'440px':'',cls:'sk-fit'})
    +(rows.length?'<div class="sk-pager"><span class="sk-sub">'+(all?'Menampilkan semua '+rows.length+' pegawai':'Menampilkan '+(from+1)+'-'+(from+part.length)+' dari '+rows.length+' pegawai')+'</span><div class="sk-row">'+(rows.length>per?'<button class="sk-b sec sm" data-act="dashPg" data-p="'+(d.p-1)+'"'+(d.p<=1?' disabled':'')+'>Sebelumnya</button><button class="sk-b sec sm" data-act="dashPg" data-p="'+(d.p+1)+'"'+(d.p>=pages?' disabled':'')+'>Berikutnya</button>':'')+'<button class="sk-b ghost sm" data-act="nav" data-k="rekap">Buka Rekap Kehadiran</button></div></div>':'');
}
INP.dashQ=t=>{S.dash.q=t.value;S.dash.p=1;drawDash()};
EF_CB.dash=()=>{const el=$('#dash-ef');if(el)el.innerHTML=efBar('dash',S.ctx.employees.filter(e=>e.is_active),{inl:1,nolim:1});S.dash.p=1;drawDash()};
CHG.dashPer=t=>{S.dash.per=+t.value;S.dash.p=1;drawDash()};
CHG.dashSt=t=>{S.dash.st=t.value;S.dash.p=1;drawDash()};
ACT.dashPg=t=>{S.dash.p=+t.dataset.p;drawDash()};
ACT.procNow=async t=>{await busy(t,async()=>{const n=await reprocess(D.add(TODAY(),-7),TODAY());toast('Rekap diproses: '+n+' baris','ok');Ctx.reset();await render(true)})};

/* =====================================================================
   14. ADMIN: REKAP KEHADIRAN
   ===================================================================== */
S.rk={tab:'ringkasan',emp:null,date:null,dept:'',q:'',st:'',sort:'name'};
/* ---- Rekap Kehadiran: tampilan visual (kartu ringkasan, komposisi, peringkat, peta kehadiran) ---- */
const RK_C={present:'#16a34a',late:'#f59e0b',official_duty:'#6366f1',leave:'#3b82f6',sick:'#8b5cf6',permit:'#06b6d4',incomplete:'#fb923c',absent:'#ef4444'};
const RK_L={present:'Tepat waktu',late:'Terlambat',official_duty:'Dinas',leave:'Cuti',sick:'Sakit',permit:'Izin',incomplete:'Belum lengkap',absent:'Alpha'};
const RK_SCHED=['present','late','incomplete','absent','leave','sick','permit','official_duty'];
const rkWho=e=>`<div class="rk-who"><span class="rk-av">${esc(initials(e.full_name))}</span><span class="rk-nm"><b>${esc(e.full_name)}</b><small>${esc(e.employee_no)}</small></span></div>`;
const rkN=(v,c)=>v?`<span class="rk-n ${c}">${v}</span>`:'<span class="rk-z">-</span>';
const rkLv=p=>p>=95?'g':p>=85?'a':'r';
const rkPct=s=>s.sched?`<div class="rk-pct"><div class="rk-bar"><i class="${rkLv(s.pct)}" style="width:${Math.min(100,s.pct)}%"></i></div><b class="${rkLv(s.pct)}">${s.pct}%</b></div>`:'<span class="rk-z">-</span>';
const rkSum=list=>{const t={sched:0,present:0,late:0,incomplete:0,leave:0,sick:0,permit:0,official_duty:0,absent:0,late_min:0,worked:0,ot:0,hadir:0};list.forEach(x=>Object.keys(t).forEach(k=>{t[k]+=x.s[k]||0}));t.pct=t.sched?Math.round((t.hadir+t.official_duty)*100/t.sched):0;return t};
function rkEmps(){const c=S.ctx,rk=S.rk,q=(rk.q||'').toLowerCase();efFix('rk',c.employees.filter(e=>e.is_active));return c.employees.filter(e=>e.is_active&&efPass('rk',e)&&(!q||e.full_name.toLowerCase().includes(q)||String(e.employee_no).toLowerCase().includes(q))).sort((a,b)=>a.full_name.localeCompare(b.full_name))}
const rkMix=(t,sub)=>{
  const keys=['present','late','official_duty','leave','sick','permit','incomplete','absent'],tot=keys.reduce((a,k)=>a+t[k],0);if(!tot)return'';
  return`<section class="rk-mix"><div class="rk-mixh"><b>Komposisi kehadiran</b><span>${tot} ${sub||'catatan hari kerja'}</span></div><div class="rk-mixbar">${keys.filter(k=>t[k]).map(k=>`<i style="width:${t[k]*100/tot}%;background:${RK_C[k]}" title="${RK_L[k]}: ${t[k]}"></i>`).join('')}</div><div class="rk-leg">${keys.map(k=>`<span><i style="background:${RK_C[k]}"></i>${RK_L[k]} <b>${t[k]}</b><em>${Math.round(t[k]*100/tot)}%</em></span>`).join('')}</div></section>`;
};
const rkAttn=list=>{
  const r=list.filter(x=>x.s.absent||x.s.late||x.s.incomplete).map(x=>({e:x.e,s:x.s,sc:x.s.absent*3+x.s.incomplete*2+x.s.late})).sort((a,b)=>b.sc-a.sc).slice(0,5);
  return`<section class="rk-panel"><h4>${ic('alert','sm')}Perlu perhatian</h4>${r.length?r.map(x=>`<button class="rk-att" data-act="rkGo" data-e="${x.e.id}" title="Lihat rincian harian">${rkWho(x.e)}<span class="rk-chips">${x.s.absent?`<span class="rk-n r">Alpha ${x.s.absent}</span>`:''}${x.s.late?`<span class="rk-n a">Terlambat ${x.s.late}</span>`:''}${x.s.incomplete?`<span class="rk-n o">Belum lengkap ${x.s.incomplete}</span>`:''}</span></button>`).join(''):`<div class="rk-ok">${ic('check','sm')}Tidak ada alpha, keterlambatan, maupun absen belum lengkap pada periode ini.</div>`}</section>`;
};
const rkDept=list=>{
  const by={};list.forEach(x=>{const k=x.e.department_id||0;(by[k]=by[k]||[]).push(x)});
  const rows=Object.keys(by).map(k=>({n:+k?deptName(+k):'Tanpa departemen',t:rkSum(by[k]),c:by[k].length})).filter(r=>r.t.sched).sort((a,b)=>b.t.pct-a.t.pct);
  return`<section class="rk-panel"><h4>${ic('building','sm')}Kehadiran per departemen</h4>${rows.length?rows.map(r=>`<div class="rk-dr"><div class="rk-drh"><span>${esc(r.n)} <small>${r.c} pegawai</small></span><b class="${rkLv(r.t.pct)}">${r.t.pct}%</b></div><div class="rk-bar"><i class="${rkLv(r.t.pct)}" style="width:${Math.min(100,r.t.pct)}%"></i></div></div>`).join(''):'<div class="sk-empty">Belum ada data</div>'}</section>`;
};
reg('rekap',{
  range:()=>payRange(pmLabel()),
  render(){
    const c=S.ctx,rk=S.rk,[from,to]=payRange(pmLabel()),t=TODAY(),def=(t<from||t>to)?from:t;
    if(!rk.date||rk.date<from||rk.date>to)rk.date=def;
    const emps=rkEmps(),days=D.range(from,to),nLock=days.filter(ds=>isLockedDate(ds)).length;
    if(!rk.emp||!emps.some(e=>e.id===+rk.emp))rk.emp=emps[0]?emps[0].id:null;
    const lockedTag=nLock?bd(nLock>=days.length?'Periode dikunci':'Sebagian terkunci','r'):'';
    const head=`<div class="rk-top"><div class="rk-tl"><span class="rk-eye">${ic('list','sm')}Rekap Kehadiran</span>${payBar()}<div class="rk-meta"><span class="rk-mc">${ic('users','sm')}${emps.length} pegawai</span><span class="rk-mc">${ic('calendar','sm')}${days.length} hari</span>${lockedTag}</div></div><div class="sk-rkact"><button class="sk-b ghost sm rk-pr" data-act="rkProc" title="Hitung ulang dan simpan rekap periode ini">${ic('refresh','sm')}<span class="lbl">Proses rekap</span></button><button class="sk-b sec sm" data-act="rkExport" title="Unduh Excel" aria-label="Unduh Excel">${ic('download','sm')}<span class="lbl">Excel</span></button><button class="sk-b sec sm" data-act="rkPrint" title="Cetak atau simpan PDF" aria-label="Cetak atau simpan PDF">${ic('print','sm')}<span class="lbl">PDF</span></button></div></div>`;
    const sortSel=rk.tab==='ringkasan'?sel('sort',[['name','Urut: Nama A-Z'],['pct_desc','Kehadiran tertinggi'],['pct_asc','Kehadiran terendah'],['alpha','Alpha terbanyak'],['late','Terlambat terlama']],rk.sort,{attr:'data-chg="rkF" aria-label="Urutkan"'}):'';
    const bar=`<div class="sk-rkbar">${stabs([['ringkasan','Ringkasan'],['pegawai','Per Pegawai'],['harian','Per Hari']],rk.tab,'rkTab')}<div class="sk-rkf">${efBar('rk',c.employees.filter(e=>e.is_active),{inl:1})}<input class="sk-in" type="search" name="q" placeholder="Cari nama / NIP" aria-label="Cari pegawai" value="${esc(rk.q)}" data-inp="rkQ">${sortSel}${rk.tab==='pegawai'?sel('emp',empOptions(emps),rk.emp,{attr:'data-chg="rkF" aria-label="Pilih pegawai"'}):''}${rk.tab==='harian'?`<input class="sk-in" type="date" name="date" aria-label="Tanggal" value="${rk.date}" min="${from}" max="${to}" data-chg="rkF">`:''}</div></div>`;
    S._rkRows=null;
    if(!emps.length)return head+bar+'<div class="sk-empty">Tidak ada pegawai sesuai filter</div>';

    /* ---------- RINGKASAN ---------- */
    if(rk.tab==='ringkasan'){
      const list=emps.map(e=>({e,s:summarize(e.id,from,to)})),tot=rkSum(list);
      const nz=(a,b)=>(!a.s.sched)-(!b.s.sched);
      const fn={pct_desc:(a,b)=>nz(a,b)||b.s.pct-a.s.pct,pct_asc:(a,b)=>nz(a,b)||a.s.pct-b.s.pct,alpha:(a,b)=>b.s.absent-a.s.absent,late:(a,b)=>b.s.late_min-a.s.late_min}[rk.sort];
      const view=fn?list.slice().sort(fn):list;
      S._rkRows={h:['Nama','NIP','Departemen','Hari Kerja','Hadir','Terlambat','Belum Lengkap','Alpha','Cuti','Sakit','Izin','Dinas','Menit Terlambat','% Hadir'],r:view.map(({e,s})=>[e.full_name,e.employee_no,deptName(e.department_id),s.sched,s.hadir,s.late,s.incomplete,s.absent,s.leave,s.sick,s.permit,s.official_duty,s.late_min,s.pct+'%'])};
      const kp='<div class="sk-kpis rk-kpis">'
        +kpi('chart','#1d4ed8',tot.sched?tot.pct+'%':'-','Tingkat kehadiran',(tot.hadir+tot.official_duty)+' dari '+tot.sched+' hari kerja')
        +kpi('check','#059669',tot.hadir,'Total hadir',tot.present+' tepat waktu')
        +kpi('clock','#d97706',tot.late,'Terlambat',fmtMin(tot.late_min)+' total')
        +kpi('alert','#dc2626',tot.absent,'Alpha',tot.incomplete+' belum lengkap')
        +kpi('calendar','#7c3aed',tot.leave+tot.sick+tot.permit+tot.official_duty,'Cuti, sakit, izin & dinas','Cuti '+tot.leave+' · Sakit '+tot.sick+' · Izin '+tot.permit+' · Dinas '+tot.official_duty)+'</div>';
      const body=tbl(['Pegawai','Departemen',{t:'Hari Kerja',c:'num'},{t:'Hadir',c:'num'},{t:'Terlambat',c:'num'},{t:'Belum Lengkap',c:'num'},{t:'Alpha',c:'num'},{t:'Cuti',c:'num'},{t:'Sakit',c:'num'},{t:'Izin',c:'num'},{t:'Dinas',c:'num'},{t:'Menit Terlambat',c:'num'},'% Hadir'],
        view.map(({e,s})=>[`<button class="rk-lk" data-act="rkGo" data-e="${e.id}" title="Lihat rincian harian">${rkWho(e)}</button>`,esc(deptName(e.department_id)),s.sched,rkN(s.hadir,'g'),rkN(s.late,'a'),rkN(s.incomplete,'o'),rkN(s.absent,'r'),rkN(s.leave,'b'),rkN(s.sick,'v'),rkN(s.permit,'c'),rkN(s.official_duty,'i'),s.late_min?s.late_min+' mnt':'<span class="rk-z">-</span>',rkPct(s)]),{empty:'Tidak ada pegawai sesuai filter',cls:'sk-fit',ef:'rk'});
      return head+bar+kp+rkMix(tot)+'<div class="rk-panels">'+rkAttn(list)+rkDept(list)+'</div>'+card('Ringkasan per pegawai <span class="rk-cnt">'+view.length+'</span>',body,'<span class="sk-sub">Klik nama untuk melihat rincian harian</span>');
    }

    /* ---------- PER PEGAWAI ---------- */
    if(rk.tab==='pegawai'){
      const e=c.employees.find(x=>x.id===+rk.emp);if(!e)return head+bar+'<div class="sk-empty">Pilih pegawai</div>';
      const s=summarize(e.id,from,to),recs=days.map(ds=>[ds,recFor(e.id,ds)]),idx=emps.findIndex(x=>x.id===e.id);
      S._rkRows={h:['Tanggal','Hari','Status','Jadwal','Masuk','Pulang','Terlambat (mnt)','Pulang Cepat (mnt)','Jam Kerja','Lembur (mnt)','Catatan'],r:recs.map(([ds,r])=>[ds,dayName(ds),(ST[r.status]||[r.status])[0],r.scheduled_start?tm(r.scheduled_start)+'-'+tm(r.scheduled_end):'-',tm(r.check_in_at),tm(r.check_out_at),r.late_minutes||0,r.early_leave_minutes||0,r.worked_minutes!=null?fmtMin(r.worked_minutes):'-',r.overtime_minutes||0,r.note||''])};
      const prof=`<section class="rk-prof"><div class="rk-pl"><span class="rk-av lg">${esc(initials(e.full_name))}</span><div class="rk-pn"><b>${esc(e.full_name)}</b><small>${esc(e.employee_no)} · ${esc(deptName(e.department_id))}${e.position?' · '+esc(e.position):''}</small></div></div><div class="rk-pm"><div class="rk-big ${s.sched?rkLv(s.pct):''}">${s.sched?s.pct+'%':'-'}</div><small>Tingkat kehadiran</small></div><div class="rk-pnav"><button class="sk-ibtn" data-act="rkEmp" data-d="-1" ${idx<=0?'disabled':''} aria-label="Pegawai sebelumnya">${ic('left')}</button><span>${idx+1} / ${emps.length}</span><button class="sk-ibtn" data-act="rkEmp" data-d="1" ${idx>=emps.length-1?'disabled':''} aria-label="Pegawai berikutnya">${ic('right')}</button></div></section>`;
      const kp='<div class="sk-kpis rk-kpis">'
        +kpi('check','#059669',s.hadir,'Hadir','dari '+s.sched+' hari kerja')
        +kpi('clock','#d97706',s.late,'Terlambat',fmtMin(s.late_min)+' total')
        +kpi('alert','#dc2626',s.absent,'Alpha',s.incomplete+' belum lengkap')
        +kpi('calendar','#7c3aed',s.leave+s.sick+s.permit+s.official_duty,'Cuti, sakit, izin & dinas','Cuti '+s.leave+' · Sakit '+s.sick+' · Izin '+s.permit+' · Dinas '+s.official_duty)
        +kpi('brief','#0891b2',fmtMin(s.worked),'Total jam kerja','Lembur '+fmtMin(s.ot))+'</div>';
      const cnt={};recs.forEach(([ds,r])=>{cnt[r.status]=(cnt[r.status]||0)+1});
      const keys=Object.keys(ST).filter(k=>cnt[k]);
      const lead=D.iso(from)-1;
      const hm=`<section class="rk-panel"><h4>${ic('calendar','sm')}Peta kehadiran periode</h4><div class="rk-hm">${DAYS3.map(d=>`<span class="rk-hd">${d}</span>`).join('')}${'<span></span>'.repeat(lead)}${recs.map(([ds,r])=>{const x=ST[r.status]||[r.status,'s'];return`<span class="rk-hc ${x[1]}${ds===t?' now':''}" title="${esc(fmtDL(ds)+' · '+x[0])}">${D.parse(ds).d}</span>`}).join('')}</div></section>`;
      const brk=`<section class="rk-panel"><h4>${ic('chart','sm')}Rincian status</h4>${keys.map(k=>`<div class="rk-sr">${bd(ST[k][0],ST[k][1])}<div class="rk-bar"><i class="${ST[k][1]}" style="width:${Math.round(cnt[k]*100/days.length)}%"></i></div><b>${cnt[k]} <small>hari</small></b></div>`).join('')}</section>`;
      const body=tbl(['Tanggal','Status','Jadwal','Masuk','Pulang',{t:'Terlambat',c:'num'},{t:'Jam Kerja',c:'num'},'Catatan',''],recs.map(([ds,r])=>[`<b>${dayName(ds)}</b>, ${fmtD(ds)}`,stBadge(r.status)+(r.is_manual?' '+bd('Manual','v'):''),r.scheduled_start?tm(r.scheduled_start)+'-'+tm(r.scheduled_end):'<span class="rk-z">-</span>',tm(r.check_in_at),tm(r.check_out_at),r.late_minutes?`<span class="rk-n a">${r.late_minutes} mnt</span>`:'<span class="rk-z">-</span>',r.worked_minutes!=null?fmtMin(r.worked_minutes):'-',esc(r.note||''),(isLockedDate(ds)||r.is_locked)?ic('lock','sm'):`<button class="sk-b sec sm" data-act="rkEdit" data-e="${e.id}" data-d="${ds}">${ic('edit','sm')}Koreksi</button>`]),{maxh:'62vh'});
      return head+bar+prof+kp+'<div class="rk-panels">'+hm+brk+'</div>'+card('Rincian harian',body);
    }

    /* ---------- PER HARI ---------- */
    const rows=emps.map(e=>[e,recFor(e.id,rk.date)]),cnt={};
    rows.forEach(([e,r])=>{cnt[r.status]=(cnt[r.status]||0)+1});
    if(rk.st&&!cnt[rk.st])rk.st='';
    const shown=rk.st?rows.filter(([e,r])=>r.status===rk.st):rows,nSch=rows.filter(([e,r])=>RK_SCHED.includes(r.status)).length,lateSum=rows.reduce((a,[e,r])=>a+(r.late_minutes||0),0);
    S._rkRows={h:['Nama','NIP','Departemen','Status','Masuk','Pulang','Terlambat (mnt)','Catatan'],r:shown.map(([e,r])=>[e.full_name,e.employee_no,deptName(e.department_id),(ST[r.status]||[r.status])[0],tm(r.check_in_at),tm(r.check_out_at),r.late_minutes||0,r.note||''])};
    const dnav=`<div class="rk-dnav"><button class="sk-ibtn" data-act="rkDay" data-d="-1" ${rk.date<=from?'disabled':''} aria-label="Hari sebelumnya">${ic('left')}</button><div class="rk-dl"><b>${fmtDL(rk.date)}</b><small>${rows.length} pegawai${nLock&&isLockedDate(rk.date)?' · terkunci':''}</small></div><button class="sk-ibtn" data-act="rkDay" data-d="1" ${rk.date>=to?'disabled':''} aria-label="Hari berikutnya">${ic('right')}</button>${t>=from&&t<=to&&rk.date!==t?'<button class="sk-b sec sm" data-act="rkDay" data-d="t">Hari ini</button>':''}</div>`;
    const kp='<div class="sk-kpis rk-kpis">'
      +kpi('check','#059669',(cnt.present||0)+(cnt.late||0),'Hadir','dari '+nSch+' terjadwal')
      +kpi('clock','#d97706',cnt.late||0,'Terlambat',fmtMin(lateSum)+' total')
      +kpi('alert','#dc2626',cnt.absent||0,'Alpha',(cnt.incomplete||0)+' belum lengkap')
      +kpi('calendar','#7c3aed',(cnt.leave||0)+(cnt.sick||0)+(cnt.permit||0)+(cnt.official_duty||0),'Cuti, sakit, izin & dinas','Libur / tanpa jadwal '+((cnt.holiday||0)+(cnt.day_off||0)+(cnt.no_schedule||0)))+'</div>';
    const chips=`<div class="rk-chs"><button class="rk-chip${rk.st?'':' on'}" data-act="rkSt" data-k="">Semua <b>${rows.length}</b></button>${Object.keys(ST).filter(k=>cnt[k]).map(k=>`<button class="rk-chip ${ST[k][1]}${rk.st===k?' on':''}" data-act="rkSt" data-k="${k}"><i></i>${ST[k][0]} <b>${cnt[k]}</b></button>`).join('')}</div>`;
    const body=tbl(['Pegawai','Departemen','Status','Jadwal','Masuk','Pulang','Catatan',''],shown.map(([e,r])=>[`<button class="rk-lk" data-act="rkGo" data-e="${e.id}" title="Lihat rincian periode">${rkWho(e)}</button>`,esc(deptName(e.department_id)),stBadge(r.status)+(r.is_manual?' '+bd('Manual','v'):''),r.scheduled_start?tm(r.scheduled_start)+'-'+tm(r.scheduled_end):'<span class="rk-z">-</span>',tm(r.check_in_at)+(r.late_minutes?` <span class="rk-n a">+${r.late_minutes}m</span>`:''),tm(r.check_out_at),esc(r.note||''),(isLockedDate(rk.date)||r.is_locked)?ic('lock','sm'):`<button class="sk-b sec sm" data-act="rkEdit" data-e="${e.id}" data-d="${rk.date}">${ic('edit','sm')}Koreksi</button>`]),{maxh:'62vh',empty:'Tidak ada data untuk status ini',cls:'sk-fit',ef:'rk'});
    return head+bar+dnav+kp+chips+card('Kehadiran harian <span class="rk-cnt">'+shown.length+'</span>',body);
  }
});
ACT.rkTab=async t=>{S.rk.tab=t.dataset.k;await render()};
ACT.rkSt=async t=>{S.rk.st=t.dataset.k||'';await render()};
ACT.rkDay=async t=>{const[f,to]=payRange(pmLabel()),d=t.dataset.d;let n=d==='t'?TODAY():D.add(S.rk.date,+d);if(n<f)n=f;if(n>to)n=to;S.rk.date=n;await render()};
ACT.rkEmp=async t=>{const list=rkEmps(),i=list.findIndex(e=>e.id===+S.rk.emp),n=list[i+(+t.dataset.d)];if(n){S.rk.emp=n.id;await render()}};
ACT.rkGo=async t=>{S.rk.emp=+t.dataset.e;S.rk.tab='pegawai';await render();const v=$('#sk-view');if(v&&v.scrollTo)v.scrollTo({top:0,behavior:'smooth'});try{window.scrollTo({top:0,behavior:'smooth'})}catch(x){}};
CHG.rkF=async t=>{S.rk[t.name]=t.value;await render()};
INP.rkQ=debounce(async t=>{S.rk.q=t.value;const pos=t.selectionStart;await render();const i=$('[data-inp=rkQ]');if(i){i.focus();try{i.setSelectionRange(pos,pos)}catch(e){}}},400);
ACT.rkProc=async t=>{if(!await confirmBox('Hitung ulang dan simpan rekap seluruh pegawai aktif untuk '+payTitle()+'? Data yang sudah dikoreksi manual dan periode terkunci tidak diubah.',{ok:'Proses'}))return;
  await busy(t,async()=>{const n=await reprocess(...payRange(pmLabel()));toast('Rekap tersimpan: '+n+' baris','ok');Ctx.reset();await render(true)})};
ACT.rkExport=()=>{if(!S._rkRows)return;exportRows('rekap-'+S.rk.tab+'-'+pmLabel().slice(0,7),'Rekap',S._rkRows.h,S._rkRows.r)};
ACT.rkPrint=()=>{if(!S._rkRows)return;printTable('Rekap Kehadiran '+payTitle(),'Tampilan: '+S.rk.tab,S._rkRows.h,S._rkRows.r)};
ACT.rkEdit=t=>{
  const emp=+t.dataset.e,ds=t.dataset.d,r=recFor(emp,ds),sch=Engine.sched(S.ctx,emp,ds);
  const hm=iso=>iso?D.local(Date.parse(iso)).hm:'';
  modal({title:'Koreksi Manual - '+nameOf(emp),body:'<div class="sk-info">'+fmtDL(ds)+'. Koreksi manual HRD akan menandai baris ini sebagai "Manual" dan tidak ditimpa proses otomatis.</div><form id="rkf">'
    +'<div class="sk-grid sk-g2">'+fld('Jam masuk',inp('in',hm(r.check_in_at),{type:'time'}))+fld('Jam pulang',inp('out',hm(r.check_out_at),{type:'time'}))+'</div>'
    +fld('Status',sel('status',Object.entries(ST).map(([k,v])=>[k,v[0]]),r.status))+fld('Catatan (wajib)',inp('note',r.note||'',{ph:'Alasan koreksi'}))+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{
      const d=formData($('#rkf',m.el));if(!d.note.trim()){toast('Catatan wajib diisi','warn');return false}
      const base=sch.kind==='work'&&sch.end<=sch.start?D.add(ds,1):ds;
      const row={employee_id:emp,work_date:ds,schedule_type:sch.src||'none',scheduled_start:sch.kind==='work'?new Date(sch.startMs).toISOString():null,scheduled_end:sch.kind==='work'?new Date(sch.endMs).toISOString():null,
        check_in_at:d.in?new Date(D.epoch(ds,d.in)).toISOString():null,check_out_at:d.out?new Date(D.epoch(base,d.out)).toISOString():null,status:d.status,late_minutes:r.late_minutes||0,early_leave_minutes:r.early_leave_minutes||0,worked_minutes:r.worked_minutes==null?null:r.worked_minutes,overtime_minutes:r.overtime_minutes||0,request_id:r.request_id||null,is_manual:true,note:d.note.trim()};
      if(row.check_in_at&&row.check_out_at)row.worked_minutes=Math.max(0,Math.floor((Date.parse(row.check_out_at)-Date.parse(row.check_in_at))/6e4)-(sch.brk||0));
      if(d.status!=='late')row.late_minutes=d.status==='present'?0:row.late_minutes;
      await DB.upsert('attendance_records',[row],'employee_id,work_date');await audit('attendance.manual','attendance_records',emp+'|'+ds,null,row);
      toast('Koreksi tersimpan','ok');Ctx.reset();render(true)}}]});
};

/* =====================================================================
   15. ADMIN: PERSETUJUAN (tahap 2 + pembatalan + override)
   ===================================================================== */
S.ap={tab:'l2',q:''};
/* v6: pembantu tampilan bersama (Persetujuan & Akun Pengguna) */
const uxAv=(e,nm)=>{const id=(e&&+e.id)||0;return'<span class="ux-av c'+(id%6)+'"><span>'+esc(initials(nm||(e&&e.full_name)))+'</span>'+(e&&e.photo_url?pgImg(e,''):'')+'</span>'};
const uxDays=x=>x?-daysTo(D.local(Date.parse(x)).date):null;
const uxAgo=x=>{const n=uxDays(x);if(n==null)return'';return n<=0?'hari ini':n===1?'kemarin':n<30?n+' hari lalu':n<365?Math.floor(n/30)+' bulan lalu':Math.floor(n/365)+' tahun lalu'};
const APCAT={leave:['calendar','#2563eb'],permit:['file','#0891b2'],sick:['heart','#7c3aed'],official_duty:['pin','#4f46e5'],time_permit:['clock','#ea580c'],correction:['edit','#d97706'],overtime:['history','#059669']};
const APDEC={approved:'Disetujui',rejected:'Ditolak',cancel_approved:'Pembatalan disetujui',cancel_rejected:'Pembatalan ditolak'};
const apLast=id=>{const l=(S.ctx.approvals||[]).filter(a=>a.request_id===id).sort((a,b)=>a.id-b.id);return l.length?l[l.length-1]:null};
function apCard(r,tab){
  const rt=rtOf(r.request_type_id),cat=APCAT[rt.category]||['file','#64748b'],e=empIx()[r.employee_id]||null,nm=nameOf(r.employee_id);
  const waiting=tab!=='hist',last=apLast(r.id),age=waiting&&r.created_at?uxDays(r.created_at):null,ov=tab==='l1'?1:0;
  const btns=(tab==='l2'||tab==='l1')
    ?`<button class="sk-b sec sm ap-no" data-act="apDecide" data-id="${r.id}" data-d="rejected" data-ov="${ov}">${ic('x','sm')}Tolak</button><button class="sk-b ok sm" data-act="apDecide" data-id="${r.id}" data-d="approved" data-ov="${ov}">${ic('check','sm')}${tab==='l1'?'Setujui (override)':'Setujui'}</button>`
    :tab==='cancel'
    ?`<button class="sk-b sec sm ap-no" data-act="apCancel" data-id="${r.id}" data-d="rejected">${ic('x','sm')}Tolak batal</button><button class="sk-b ok sm" data-act="apCancel" data-id="${r.id}" data-d="approved">${ic('check','sm')}Setujui batal</button>`:'';
  const meta=rstBadge(r.status)+(r.cancel_requested?bd('Minta batal','o'):'')
    +(age!=null?`<span class="ux-pill${age>=2&&waiting?' warn':''}" title="Diajukan ${esc(fmtDT(r.created_at))}">${ic('clock','sm')}Diajukan ${esc(uxAgo(r.created_at))}</span>`:'')
    +(!waiting&&last?`<span class="ux-pill" title="${esc(fmtDT(last.decided_at))}">${ic(last.decision==='rejected'||last.decision==='cancel_rejected'?'x':'check','sm')}${esc(APDEC[last.decision]||last.decision)} oleh ${esc(last.approver_name||'-')}</span>`:'');
  return`<article class="ux-card ap-item s-${tab==='hist'?esc(r.status):tab}">`
    +`<div class="ux-h">${uxAv(e,nm)}<div class="ux-id"><b title="${esc(nm)}">${esc(nm)}</b><small>${esc(deptOf(r.employee_id))}</small></div><span class="sk-sub">#${r.id}</span></div>`
    +`<div class="ap-ty" style="--c:${cat[1]}"><span class="ap-ti">${ic(cat[0])}</span><div class="tx"><b>${esc(rt.name)}</b>${(CAT[rt.category]||'')&&CAT[rt.category]!==rt.name?`<small>${esc(CAT[rt.category])}</small>`:''}</div>${+r.total_days>0?`<div class="ap-days"><b>${nf(r.total_days)}</b><small>hari</small></div>`:''}</div>`
    +`<div class="ap-dt">${ic('calendar','sm')}<span>${esc(reqPeriod(r))}${esc(reqExtra(r))}</span></div>`
    +(r.reason?`<p class="ap-rs" title="${esc(r.reason)}">${esc(r.reason)}</p>`:'<p class="ap-rs none">Tanpa alasan</p>')
    +(tab==='cancel'?`<div class="ap-cx">${ic('alert','sm')}<span><b>Alasan pembatalan:</b> ${esc(r.cancel_reason||'-')}</span></div>`:'')
    +`<div class="ap-meta">${meta}</div>`
    +`<div class="ap-f"><button class="sk-b sec sm" data-act="apView" data-id="${r.id}">${ic('eye','sm')}Detail</button><span class="sp"></span>${btns}</div></article>`;
}
const EMPTY_AP='Tidak ada pengajuan';
function apRow(r,tab){
  const rt=rtOf(r.request_type_id),e=empIx()[r.employee_id]||null,nm=nameOf(r.employee_id),waiting=tab!=='hist',last=apLast(r.id),age=waiting&&r.created_at?uxDays(r.created_at):null,ov=tab==='l1'?1:0;
  const btns=(tab==='l2'||tab==='l1')
    ?`<button class="sk-b sec sm ap-no" data-act="apDecide" data-id="${r.id}" data-d="rejected" data-ov="${ov}">Tolak</button><button class="sk-b ok sm" data-act="apDecide" data-id="${r.id}" data-d="approved" data-ov="${ov}">${tab==='l1'?'Setujui (override)':'Setujui'}</button>`
    :tab==='cancel'
    ?`<button class="sk-b sec sm ap-no" data-act="apCancel" data-id="${r.id}" data-d="rejected">Tolak batal</button><button class="sk-b ok sm" data-act="apCancel" data-id="${r.id}" data-d="approved">Setujui batal</button>`:'';
  const st=rstBadge(r.status)+(r.cancel_requested?' '+bd('Minta batal','o'):'')
    +(age!=null?`<div class="sk-sub${age>=2&&waiting?' ap-late':''}">Diajukan ${esc(uxAgo(r.created_at))}</div>`:'')
    +(!waiting&&last?`<div class="sk-sub">${esc(APDEC[last.decision]||last.decision)} oleh ${esc(last.approver_name||'-')}</div>`:'');
  const why=(r.reason?esc(r.reason):'<span class="sk-sub">Tanpa alasan</span>')+(tab==='cancel'?`<div class="sk-sub"><b>Batal:</b> ${esc(r.cancel_reason||'-')}</div>`:'');
  return[
    `<div class="ap-emp">${uxAv(e,nm)}<div><b>${esc(nm)}</b><div class="sk-sub">${esc(deptOf(r.employee_id))} &middot; #${r.id}</div></div></div>`,
    esc(rt.name),
    esc(reqPeriod(r)+reqExtra(r)),
    +r.total_days>0?nf(r.total_days):'-',
    `<div class="ap-why" title="${esc(r.reason||'')}">${why}</div>`,
    st,
    `<div class="ap-btns"><button class="sk-b sec sm" data-act="apView" data-id="${r.id}">Detail</button>${btns}</div>`
  ];
}
reg('persetujuan',{
  render(){
    const c=S.ctx,a=S.ap,rs=c.requests;
    const l2=rs.filter(r=>r.status==='waiting_l2'),cr=rs.filter(r=>r.status==='approved'&&r.cancel_requested),l1=rs.filter(r=>r.status==='waiting_l1'),hist=rs.filter(r=>['approved','rejected','cancelled'].includes(r.status)&&!(r.status==='approved'&&r.cancel_requested));
    const sets={l2,cancel:cr,l1,hist};if(!sets[a.tab])a.tab='l2';const tab=a.tab;
    const q=(a.q||'').trim().toLowerCase();
    let list=sets[tab];efFix('ap',c.employees);list=list.filter(r=>efPassId('ap',r.employee_id));
    if(q)list=list.filter(r=>[nameOf(r.employee_id),deptOf(r.employee_id),rtOf(r.request_type_id).name,r.reason||''].join(' ').toLowerCase().includes(q));
    list=list.slice().sort(tab==='hist'?(x,y)=>y.id-x.id:(x,y)=>String(x.start_date).localeCompare(String(y.start_date))||x.id-y.id);
    const pend=l2.length+cr.length;
    const sub=pend?l2.length+' pengajuan menunggu keputusan akhir'+(cr.length?' dan '+cr.length+' permintaan pembatalan':'')+'.':'Tidak ada pengajuan yang menunggu keputusan Anda saat ini.';
    const hero=`<div class="rk-top ux-hero"><div class="rk-tl"><span class="rk-eye">${ic('check','sm')}Persetujuan</span><div class="ux-ht">Kotak Persetujuan</div><div class="ux-hs">${esc(sub)}</div><div class="rk-meta"><span class="rk-mc">${ic('inbox','sm')}${l2.length} menunggu HRD</span><span class="rk-mc">${ic('x','sm')}${cr.length} pembatalan</span><span class="rk-mc">${ic('users','sm')}${l1.length} di Kepala Dept.</span></div></div><div class="ux-hr">${monthBar()}</div></div>`;
    const T=[['l2','Menunggu HRD',l2.length,'inbox','#2563eb'],['cancel','Permintaan pembatalan',cr.length,'x','#ea580c'],['l1','Menunggu Kepala Dept.',l1.length,'users','#d97706'],['hist','Riwayat keputusan',hist.length,'history','#64748b']];
    const tabs='<div class="ux-tabs" role="tablist">'+T.map(([k,l,n,i,col])=>`<button type="button" role="tab" aria-selected="${k===tab}" class="ux-tab${k===tab?' on':''}" style="--c:${col}" data-act="apTab" data-k="${k}"><span class="i">${ic(i)}</span><span class="t"><b>${n}</b><small>${esc(l)}</small></span></button>`).join('')+'</div>';
    const tool=`<div class="ux-tool"><label class="ux-sr">${ic('search','sm')}<input class="sk-in" type="search" placeholder="Cari pegawai, jenis, atau alasan" aria-label="Cari pengajuan" value="${esc(a.q||'')}" data-inp="apQ"></label>${efBar('ap',c.employees,{inl:1})}<span class="ux-cnt">${list.length} pengajuan${tab==='hist'?' &middot; bulan terpilih':''}</span></div>`;
    const EM={l2:['Tidak ada pengajuan yang menunggu','Semua pengajuan sudah Anda putuskan.'],cancel:['Tidak ada permintaan pembatalan','Belum ada pegawai yang meminta pembatalan.'],l1:['Tidak ada yang menunggu Kepala Departemen','Semua pengajuan sudah diputuskan di tahap 1.'],hist:['Belum ada riwayat keputusan','Tidak ada pengajuan yang diputuskan pada bulan terpilih.']}[tab];
    const body=list.length?card('',tbl(['Pegawai','Jenis',{t:'Periode',c:'nw'},{t:'Hari',c:'num'},'Alasan','Status',{t:'Aksi',c:'ap-act'}],list.map(r=>apRow(r,tab)),{cls:'ap-tbl',page:30,empty:EMPTY_AP,ef:'ap'}))
      :`<div class="ux-empty"><div class="ei">${ic((q||efOn('ap'))?'search':'check')}</div><b>${(q||efOn('ap'))?'Tidak ada hasil':EM[0]}</b><p>${(q||efOn('ap'))?'Coba kata kunci atau filter lain.':EM[1]}</p></div>`;
    return hero+tabs
      +(!isMgr()?'<div class="sk-info">Anda dapat melihat seluruh pengajuan, tetapi keputusan hanya dapat diberikan oleh HRD Manager atau Super Admin.</div>':'')
      +(tab==='l1'?'<div class="sk-warn">Pengajuan di tab ini seharusnya diputuskan Kepala Departemen. HRD Manager / Super Admin hanya mengambil alih bila kepala departemen berhalangan; keputusan akan dicatat sebagai override.</div>':'')
      +tool+body;
  }
});
INP.apQ=debounce(async t=>{S.ap.q=t.value;const p=t.selectionStart;await render();const i=$('[data-inp=apQ]');if(i){i.focus();try{i.setSelectionRange(p,p)}catch(e){}}},300);
ACT.apTab=async t=>{S.ap.tab=t.dataset.k;await render()};
ACT.apView=t=>{const r=S.ctx.requests.find(x=>x.id===+t.dataset.id);if(r)modal({title:'Detail Pengajuan #'+r.id,size:'wide',body:requestDetail(r,true),actions:[{t:'Tutup',c:'sec',fn:()=>{}}]})};
async function afterRequestChange(res){
  if(res&&res.employee_id&&res.start_date){const n=await applyRange(res.employee_id,String(res.start_date).slice(0,10),String(res.end_date).slice(0,10));if(res.id)await rpc('adm_mark_applied',{p_ids:[res.id]}).catch(()=>{})}
}
ACT.apDecide=async t=>{
  if(!isMgr()){toast('Hanya HRD Manager atau Super Admin yang dapat memutuskan pengajuan','warn');return}
  const id=+t.dataset.id,dec=t.dataset.d,ov=t.dataset.ov==='1',r=S.ctx.requests.find(x=>x.id===id);
  const note=await promptBox(dec==='approved'?'Setujui pengajuan':'Tolak pengajuan',dec==='approved'?'Catatan (opsional)':'Alasan penolakan (wajib)',{required:dec!=='approved',ok:dec==='approved'?'Setujui':'Tolak',danger:dec!=='approved',hint:ov?'Pengajuan ini belum diputuskan Kepala Departemen. Keputusan Anda dicatat sebagai override.':(r?nameOf(r.employee_id)+' - '+rtOf(r.request_type_id).name+' '+reqPeriod(r):'')});
  if(note===null)return;
  await busy(t,async()=>{const res=await rpc('adm_decide_final',{p_id:id,p_decision:dec,p_note:note,p_override:ov});await afterRequestChange(res);toast(dec==='approved'?'Pengajuan disetujui':'Pengajuan ditolak','ok');Ctx.reset();await render(true)});
};
ACT.apCancel=async t=>{
  if(!isMgr()){toast('Hanya HRD Manager atau Super Admin yang dapat memutuskan pembatalan','warn');return}
  const id=+t.dataset.id,dec=t.dataset.d;const note=await promptBox(dec==='approved'?'Setujui pembatalan':'Tolak pembatalan','Catatan'+(dec==='approved'?' (opsional)':' (opsional)'),{required:false,ok:'Kirim'});if(note===null)return;
  await busy(t,async()=>{const res=await rpc('adm_decide_cancel',{p_id:id,p_decision:dec,p_note:note});if(dec==='approved')await afterRequestChange({...res,id:null});toast('Keputusan pembatalan tersimpan','ok');Ctx.reset();await render(true)});
};

/* =====================================================================
   16. ADMIN: PEGAWAI & DEPARTEMEN
   ===================================================================== */
const EMAIL_RE=/^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$/i,PHONE_RE=/^\+?[0-9][0-9 ()-]{5,19}$/;
const BPJS_RE=/^\d{8,20}$/;
function empValidate(d){
  if(d.email&&!EMAIL_RE.test(String(d.email).trim()))return'Format email tidak valid';
  if(d.phone&&!PHONE_RE.test(String(d.phone).trim()))return'Format No. HP tidak valid (6-20 digit, boleh diawali +)';
  if(d.join_date&&d.end_date&&d.end_date<d.join_date)return'Tanggal berhenti tidak boleh sebelum tanggal bergabung';
  if(d.gender&&!GENDER[d.gender])return'Jenis kelamin tidak valid';
  if(d.marital_status&&!MARITAL[d.marital_status])return'Status pernikahan tidak valid';
  if(d.education_level&&!EDU.includes(d.education_level))return'Pendidikan terakhir tidak valid';
  if(d.birth_date&&d.birth_date>TODAY())return'Tanggal lahir tidak boleh di masa depan';
  if(d.birth_date&&d.join_date&&d.join_date<d.birth_date)return'Tanggal bergabung tidak boleh sebelum tanggal lahir';
  if(d.join_date&&d.contract_end_date&&d.contract_end_date<d.join_date)return'Tanggal akhir kontrak tidak boleh sebelum tanggal bergabung';
  if(d.bpjs_kesehatan&&!BPJS_RE.test(d.bpjs_kesehatan))return'Nomor BPJS Kesehatan tidak valid (hanya angka, 8-20 digit)';
  if(d.bpjs_ketenagakerjaan&&!BPJS_RE.test(d.bpjs_ketenagakerjaan))return'Nomor BPJS Ketenagakerjaan tidak valid (hanya angka, 8-20 digit)';
  return''}
/* tautan Google Drive (file/d/ID, open?id=, uc?id=) atau ID polos -> https://lh3.googleusercontent.com/d/ID */
function gdriveUrl(u){u=String(u||'').trim();if(!u)return'';
  const m=u.match(/drive\.google\.com\/file\/d\/([\w-]{10,})/i)||u.match(/drive\.google\.com\/(?:open|uc|thumbnail)\?(?:[^#]*&)?id=([\w-]{10,})/i)||u.match(/^https?:\/\/lh\d\.googleusercontent\.com\/d\/([\w-]{10,})/i)||u.match(/^([\w-]{25,})$/);
  return m?'https://lh3.googleusercontent.com/d/'+m[1]:u}
const photoUrlOk=u=>/^https:\/\/\S+$/i.test(u)||/^data:image\//i.test(u);
const daysTo=x=>Math.round((Date.parse(String(x).slice(0,10)+'T00:00:00Z')-Date.parse(TODAY()+'T00:00:00Z'))/864e5);
/* kolom tambahan -> objek siap simpan (kosong = null) */
function empX(d){const o={};EMP_X.forEach(k=>{const v=d[k];o[k]=v==null||v===''?null:v});return o}
/* RPC adm_create_employees lama belum mengenal kolom tambahan: simpan lewat UPDATE setelah insert */
async function empXSave(rows){
  if(IS_PREVIEW||S.legacyEmp)return 0;let fail=0;
  for(const part of chunk(rows,8))await Promise.all(part.map(async r=>{const x=empX(r);if(r.photo_url)x.photo_url=r.photo_url;
    if(!Object.values(x).some(Boolean))return;
    try{await DB.update('employees',x,{employee_no:r.employee_no})}catch(er){fail++;console.warn('simpan kolom tambahan gagal',er)}}));
  return fail}
const pgEdu=e=>[e.education_level,e.education_major].filter(Boolean).join(' ')+(e.education_institution?' · '+e.education_institution:'');
function pgContract(e){const base='Kontrak s.d. '+esc(fmtD(String(e.contract_end_date).slice(0,10)));
  if(!e.is_active||e.employment_status!=='contract')return base;const n=daysTo(e.contract_end_date);
  return n<0?base+' <b style="color:#b91c1c">(berakhir)</b>':n<=30?base+' <b style="color:#b45309">('+n+' hari lagi)</b>':base}
function pgNotice(act){let h='';
  if(S.legacyEmp)h+='<div class="sk-warn" style="margin-bottom:12px">Kolom data karyawan yang baru (jenis kelamin, pendidikan, BPJS, dll.) belum ada di database. Jalankan <b>migrasi-karyawan.sql</b> di Supabase &gt; SQL Editor, lalu muat ulang halaman.</div>';
  const soon=act.filter(e=>e.employment_status==='contract'&&e.contract_end_date&&daysTo(e.contract_end_date)<=30);
  if(soon.length)h+='<div class="sk-warn" style="margin-bottom:12px">'+soon.length+' kontrak berakhir dalam 30 hari atau sudah lewat: '+esc(soon.slice(0,5).map(e=>e.full_name).join(', '))+(soon.length>5?', dll.':'')+'</div>';
  return h}
function pgDivCount(c){const s=new Set(),ok=e=>e.department_id!=null&&e.department_id!==''&&c.departments.some(d=>d.id===e.department_id);c.employees.filter(e=>e.is_active).forEach(e=>{const v=efDiv(e);if(v)s.add((ok(e)?e.department_id:'-')+'|'+v.toLowerCase())});return s.size}
const PG_FULL_H=['No.','NIP','Nama Lengkap','Gelar','Jenis Kelamin','Tgl Bergabung','Tempat Lahir','Tgl Lahir','Akhir Kontrak','Departemen','Divisi','Jabatan','Pendidikan Terakhir','Nama Pendidikan','Jurusan','BPJS Kesehatan','BPJS Ketenagakerjaan','Status Karyawan','No. Handphone','Email','Status Pernikahan','Pas Photo',''];
function pgFullRow(e,i){
  const dash='<span class="sk-sub">-</span>',t=v=>v&&v!=='-'?esc(v):dash,d=v=>v?esc(fmtD(String(v).slice(0,10))):dash,
    ph=e.photo_url?'<span class="pg-th'+(e.is_active?'':' off')+'"><span class="pg-init">'+pgInit(e)+'</span>'+pgImg(e,'')+'</span>'+(/^https?:/i.test(e.photo_url)?' <a href="'+esc(e.photo_url)+'" target="_blank" rel="noopener noreferrer">Buka</a>':''):dash;
  return[i+1,esc(e.employee_no),'<b>'+esc(e.full_name)+'</b>'+(e.is_active?'':' '+bd('Nonaktif','r')),t(e.title_suffix),t(GENDER[e.gender]),d(e.join_date),t(e.birth_place),d(e.birth_date),e.contract_end_date?pgContract(e):dash,
    t(deptName(e.department_id)),t(e.division),t(e.position),t(e.education_level),t(e.education_institution),t(e.education_major),t(e.bpjs_kesehatan),t(e.bpjs_ketenagakerjaan),pgStatBd(e),
    pgTel(e)||dash,pgMail(e)||dash,t(MARITAL[e.marital_status]),ph,'<button class="sk-b sec sm" data-act="empEdit" data-id="'+e.id+'">'+ic('edit','sm')+'Ubah</button>']}
function tenure(j,e){if(!j)return'-';const a=D.parse(String(j).slice(0,10)),b=D.parse(e?String(e).slice(0,10):TODAY());let m=(b.y-a.y)*12+(b.m-a.m)-(b.d<a.d?1:0);if(m<0)return'Belum mulai';const y=Math.floor(m/12);m%=12;return(y?y+' th ':'')+m+' bln'}
const empName=e=>{const n=e.full_name||'',t=(e.title_suffix||'').trim();return t&&!n.toLowerCase().includes(t.toLowerCase())?n+', '+t:n};
const canRole=()=>S.mode==='admin'&&['admin_sistem','manajer_hrd'].includes(S.admin.role);
S.pg={tab:'pegawai',q:'',dept:'',type:'',act:'1',st:'',div:'',gender:'',view:'card',lim:24,per:24,ddept:''};
const PG_STEP=24,normSp=v=>String(v==null?'':v).replace(/\s+/g,' ').trim();
/* rapikan isian teks pegawai: spasi ganda, koma di ujung gelar, huruf kecil email */
function normEmp(d){d.full_name=normSp(d.full_name);d.title_suffix=normSp(d.title_suffix).replace(/^[,\s]+|[,\s]+$/g,'');d.position=normSp(d.position);d.division=normSp(d.division);d.phone=normSp(d.phone);d.email=normSp(d.email).toLowerCase();d.nip=normSp(d.nip);for(const k of ['birth_place','education_institution','education_major'])d[k]=normSp(d[k]);for(const k of ['bpjs_kesehatan','bpjs_ketenagakerjaan'])d[k]=String(d[k]==null?'':d[k]).replace(/[\s.\-]/g,'');return d}
const nipKey=v=>String(v||'').replace(/\s+/g,'');
const pgNm=e=>String(e.full_name||'').replace(/^((Prof|Dr|Dra|Drs|Ir|H|Hj)\.\s*)+/i,'');
const pgInit=e=>esc(initials(pgNm(e)));
/* foto pegawai: berkas/URL dikirim apa adanya, bila gagal dimuat tampil inisial */
const pgImg=(e,alt)=>e.photo_url?'<img src="'+esc(e.photo_url)+'" alt="'+esc(alt||'')+'" loading="lazy" decoding="async" referrerpolicy="no-referrer" data-pgf="1">':'';
document.addEventListener('error',ev=>{const t=ev.target;if(t&&t.tagName==='IMG'&&t.dataset&&t.dataset.pgf)t.style.display='none'},true);
const pgAvatar=e=>'<span class="sk-av pg-av'+(e.type==='teacher'?' t':'')+'"><span>'+pgInit(e)+'</span>'+pgImg(e,'')+'</span>';
const pgStatBd=e=>bd(EMPST[e.employment_status]||'-',e.employment_status==='permanent'?'g':e.employment_status==='contract'?'a':'o');
const pgRoles=e=>{const a=(e.roles||[]).filter(r=>r!=='pegawai').map(r=>bd(ROLE[r]||r,'b'));if((S.ctx.departments||[]).some(d=>d.head_employee_id===e.id))a.unshift(bd('Manager / Kepala Dept.','v'));return a.join(' ')};
const pgTenure=e=>tenure(e.join_date,e.end_date)+(e.end_date?' (berhenti '+fmtD(e.end_date)+')':'');
const pgMail=e=>e.email?'<a href="mailto:'+esc(e.email)+'">'+esc(e.email)+'</a>':'';
const pgTel=e=>e.phone?'<a href="tel:'+esc(String(e.phone).replace(/[^\d+]/g,''))+'">'+esc(e.phone)+'</a>':'';
const pgWhere=e=>[deptName(e.department_id),e.division].filter(v=>v&&v!=='-').join(' · ');
const PG_ICO={
  mail:'<svg class="sk-ic" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/></svg>',
  tel:'<svg class="sk-ic" viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2z"/></svg>'};
/* kartu pegawai ringkas: foto 3:4 (klik = detail), nama + gelar, jabatan, NIP, tombol WhatsApp */
const PG_WA=`<svg class="fx-wa-ic" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M17.47 14.38c-.3-.15-1.76-.87-2.03-.97-.27-.1-.47-.15-.67.15-.2.3-.77.97-.94 1.16-.17.2-.35.22-.64.07-.3-.15-1.26-.46-2.39-1.48-.88-.79-1.48-1.76-1.65-2.06-.17-.3-.02-.46.13-.6.13-.14.3-.35.45-.52.15-.17.2-.3.3-.5.1-.2.05-.37-.03-.52-.07-.15-.67-1.61-.92-2.2-.24-.58-.49-.5-.67-.51h-.57c-.2 0-.52.07-.79.37-.27.3-1.04 1.02-1.04 2.48s1.07 2.88 1.21 3.07c.15.2 2.1 3.2 5.08 4.49.71.3 1.26.49 1.69.63.71.23 1.36.2 1.87.12.57-.09 1.76-.72 2-1.41.25-.7.25-1.29.17-1.41-.07-.13-.27-.2-.57-.35M12.05 21.79h-.01a9.87 9.87 0 0 1-5.03-1.38l-.36-.21-3.74.98 1-3.65-.24-.37a9.86 9.86 0 0 1-1.51-5.26c0-5.45 4.44-9.89 9.89-9.89 2.64 0 5.12 1.03 6.99 2.9a9.83 9.83 0 0 1 2.89 7c0 5.45-4.44 9.88-9.88 9.88M20.46 3.49A11.8 11.8 0 0 0 12.05 0C5.5 0 .16 5.34.16 11.89c0 2.1.55 4.14 1.59 5.95L.06 24l6.3-1.65a11.9 11.9 0 0 0 5.68 1.45h.01c6.55 0 11.89-5.34 11.89-11.89 0-3.18-1.24-6.17-3.48-8.42"/></svg>`;
const waNum=p=>{let d=String(p||'').replace(/[^\d+]/g,'').replace(/\+/g,'');if(!d)return'';if(d.startsWith('00'))d=d.slice(2);if(d.startsWith('0'))d='62'+d.slice(1);else if(d.startsWith('8'))d='62'+d;return d.length>=10?d:''};
const pgWA=(e,full)=>{
  const n=waNum(e.phone),lb=full?'<span>WhatsApp</span>':'',cl='fx-wa'+(full?' full':'');
  return n?`<a class="${cl}" href="https://wa.me/${n}?text=${encodeURIComponent('Assalamu\u2019alaikum '+empName(e)+', ')}" target="_blank" rel="noopener noreferrer" title="Chat WhatsApp ${esc(e.phone)}" aria-label="Chat WhatsApp ${esc(empName(e))}">${PG_WA}${lb}</a>`
    :`<span class="${cl} off" title="Nomor HP belum diisi" aria-label="Nomor HP belum diisi">${PG_WA}${lb}</span>`};
function pgCard(e){
  const nm=empName(e),off=!e.is_active;
  return`<article class="fx-card${off?' off':''}"><button type="button" class="fx-ph ${e.type==='teacher'?'t':'s'}" data-act="empView" data-id="${e.id}" title="Lihat detail" aria-label="Lihat detail ${esc(nm)}"><span class="fx-init">${pgInit(e)}</span>${pgImg(e,'Foto '+e.full_name)}<span class="fx-dot${off?' off':''}" title="${off?'Nonaktif':'Aktif'}"></span></button><div class="fx-body"><b class="fx-nm" title="${esc(nm)}">${esc(nm)}</b><span class="fx-pos">${esc(e.position||'Jabatan belum diisi')}</span><span class="fx-nipc">NIP: ${esc(e.employee_no)}</span></div></article>`}
/* urutan kartu: per divisi, laki-laki dulu lalu perempuan, lalu nama */
const pgGRank=e=>e.gender==='L'?0:e.gender==='P'?1:2;
const pgOrd=(a,b)=>{const da=(a.division||'').trim(),db=(b.division||'').trim();if(!da!==!db)return da?-1:1;return da.localeCompare(db,'id')||pgGRank(a)-pgGRank(b)||a.full_name.localeCompare(b.full_name,'id')};
function pgGroups(arr){const m=new Map();arr.forEach(e=>{const k=(e.division||'').trim()||'Tanpa Divisi';if(!m.has(k))m.set(k,[]);m.get(k).push(e)});return[...m.entries()].map(([k,v])=>'<div class="fx-gh"><span>'+esc(k)+'</span><small>'+v.length+'</small></div><div class="fx-grid">'+v.map(pgCard).join('')+'</div>').join('')}
/* detail karyawan (modal) - minimalis; baris mengikuti kolom tabel employees */
const fxNa='<i class="fx-na">&mdash;</i>';
const fxV=v=>v&&v!=='-'?esc(v):fxNa;
const fxD=v=>v?esc(fmtD(String(v).slice(0,10))):fxNa;
const fxKv=(l,v)=>`<div class="fx-kv"><span>${esc(l)}</span><b>${v}</b></div>`;
const fxSec=(title,rows)=>`<section class="fx-sec"><h4>${esc(title)}</h4><div class="fx-kvs">${rows.join('')}</div></section>`;
function fxAge(b){if(!b)return'';const a=D.parse(String(b).slice(0,10)),n=D.parse(TODAY());let y=n.y-a.y;if(n.m<a.m||(n.m===a.m&&n.d<a.d))y--;return y>0&&y<100?' ('+y+' th)':''}
ACT.empView=t=>{
  const e=S.ctx.employees.find(x=>x.id===+t.dataset.id);if(!e)return;
  const nm=empName(e),roles=pgRoles(e),adm=S.mode==='admin',
    tel=e.phone?`<a class="fx-qb" href="tel:${esc(String(e.phone).replace(/[^\d+]/g,''))}">${PG_ICO.tel}Telepon</a>`:'',
    mail=e.email?`<a class="fx-qb" href="mailto:${esc(e.email)}">${PG_ICO.mail}Email</a>`:'';
  const fps=adm?(S.ctx.fingerprints||[]).filter(f=>f.employee_id===e.id).map(f=>f.fingerprint_id):[];
  const lk=!!(e.locked_until&&Date.parse(e.locked_until)>Date.now());
  const kepeg=[fxKv('NIP',fxV(String(e.nip||e.employee_no||''))),
    ...(adm?[fxKv('ID fingerprint',fps.length?esc(fps.join(', ')):fxNa)]:[]),
    fxKv('Tipe',fxV(TYPE[e.type])),fxKv('Status karyawan',pgStatBd(e)),
    fxKv('Departemen',fxV(deptName(e.department_id))),fxKv('Divisi',fxV(e.division)),fxKv('Jabatan',fxV(e.position)),
    ...(adm?[]:[fxKv('Peran',roles||bd('Pegawai','o'))]),
    fxKv('Tanggal bergabung',fxD(e.join_date)),fxKv('Masa kerja',esc(tenure(e.join_date,e.end_date))),
    ...((e.employment_status==='contract'||e.contract_end_date)?[fxKv('Akhir kontrak',e.contract_end_date?pgContract(e):fxNa)]:[]),
    ...(e.end_date?[fxKv('Tanggal berhenti',fxD(e.end_date))]:[])];
  const a0=adm?curAssign(e.id,TODAY()):null,ws0=a0?(S.ctx.work_schedules||[]).find(w=>w.id===a0.work_schedule_id):null,dp0=(S.ctx.departments||[]).find(d=>d.id===e.department_id),hd0=dp0&&dp0.head_employee_id?S.ctx.employees.find(x=>x.id===dp0.head_employee_id):null,
    pbtn=(act,l)=>'<button type="button" class="sk-b sec sm" data-act="'+act+'" data-id="'+e.id+'">'+ic('edit','sm')+l+'</button>';
  const jp=adm?[fxKv('Jadwal kerja',(ws0?bd(ws0.name,'b')+' <span class="sk-sub">sejak '+esc(fmtD(String(a0.effective_from).slice(0,10)))+'</span>':(hasShifts(e.id)?bd('Jadwal shift','v'):bd('Belum ada','a')))+' '+pbtn('empJadwal','Ubah')),
    fxKv('Peran',(roles||bd('Pegawai','o'))+(canRole()?' '+pbtn('empRole','Ubah'):'')),
    fxKv('Atasan langsung (tahap 1)',hd0&&hd0.id!==e.id?esc(empName(hd0))+' <span class="sk-sub">Kepala '+esc(dp0.name)+'</span>':'<span class="sk-sub">'+(hd0?'Kepala departemennya sendiri: pengajuan langsung ke HRD':'Departemen belum punya kepala: pengajuan langsung ke HRD')+'</span>')]:[];
  const body=`<div class="fx-dv"><div class="fx-hero"><div class="fx-hph ${e.type==='teacher'?'t':'s'}"><span class="fx-init">${pgInit(e)}</span>${pgImg(e,'Foto '+e.full_name)}</div><div class="fx-hin"><div class="fx-hn">${esc(nm)}</div><div class="fx-hp">${esc(e.position||'Jabatan belum diisi')}</div><div class="fx-hw">${esc(pgWhere(e)||'Departemen belum diisi')}</div><div class="fx-bds">${pgStatBd(e)}${e.is_active?'':' '+bd('Nonaktif','r')}</div></div></div>`
   +`<div class="fx-qa">${pgWA(e,true)}${tel}${mail}</div>`
   +fxSec('Kepegawaian',kepeg)+(adm?fxSec('Jadwal Kerja & Peran',jp):'')
   +fxSec('Data Pribadi',[fxKv('Jenis kelamin',fxV(GENDER[e.gender])),fxKv('Status pernikahan',fxV(MARITAL[e.marital_status])),fxKv('Tempat lahir',fxV(e.birth_place)),fxKv('Tanggal lahir',e.birth_date?fxD(e.birth_date)+esc(fxAge(e.birth_date)):fxNa)])
   +fxSec('Pendidikan',[fxKv('Pendidikan terakhir',fxV(e.education_level)),fxKv('Nama pendidikan',fxV(e.education_institution)),fxKv('Jurusan',fxV(e.education_major))])
   +fxSec('Kontak',[fxKv('No. Handphone',pgTel(e)||fxNa),fxKv('Email',pgMail(e)||fxNa)])
   +fxSec('BPJS',[fxKv('BPJS Kesehatan',fxV(e.bpjs_kesehatan)),fxKv('BPJS Ketenagakerjaan',fxV(e.bpjs_ketenagakerjaan))])
   +(adm?fxSec('Akun',[fxKv('Status akun',!e.is_active?bd('Nonaktif','r'):lk?bd('Terkunci','r'):bd('Aktif','g')),fxKv('Login terakhir',e.last_login_at?esc(fmtDT(e.last_login_at)):fxNa),fxKv('Kata sandi',e.must_change_password?bd('Wajib diganti','a'):'Sudah diganti')]):'')
   +`</div>`;
  modal({title:'Detail Karyawan',size:'wide fx-md',body,actions:[{t:'Tutup',c:'sec',fn:()=>{}},...(e.is_active?[{t:'Rekap kehadiran',c:'sec',i:'list',fn:()=>{S.rk.emp=e.id;S.rk.tab='pegawai';S.rk.dept='';S.rk.q='';go('rekap')}}]:[]),{t:'Ubah data',i:'edit',fn:()=>{setTimeout(()=>ACT.empEdit({dataset:{id:String(e.id)}}),0)}}]});
};
/* ---- Detail Karyawan: ubah jadwal kerja & peran ---- */
const closeOv=t=>{const o=t&&t.closest?t.closest('.sk-ov'):null;if(o)o.remove()};
const reopenEmp=async id=>{Ctx.reset();await render(true);setTimeout(()=>ACT.empView({dataset:{id:String(id)}}),30)};
ACT.empJadwal=t=>{
  const e=S.ctx.employees.find(x=>x.id===+t.dataset.id);if(!e)return;closeOv(t);
  const a=curAssign(e.id,TODAY()),wl=(S.ctx.work_schedules||[]).filter(w=>w.is_active);
  modal({title:'Jadwal kerja - '+e.full_name,body:'<form id="jaf">'+(hasShifts(e.id)&&!a?'<div class="sk-info">Pegawai ini memakai jadwal shift (diatur di menu Jadwal Shift). Jadwal normal di sini hanya berlaku pada tanggal yang tidak punya jadwal shift terbit.</div>':'')
    +fld('Jadwal normal',sel('ws',wl.map(w=>[w.id,w.name]),a?a.work_schedule_id:''))+fld('Berlaku mulai tanggal',inp('from',TODAY(),{type:'date',req:1}),'Riwayat sebelum tanggal ini tidak berubah')+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{setTimeout(()=>ACT.empView({dataset:{id:String(e.id)}}),0)}},{t:'Simpan',i:'check',keep:true,fn:async m=>{
      const d=formData($('#jaf',m.el));if(!d.ws){toast('Belum ada jadwal kerja aktif. Buat dulu di menu Jadwal Normal.','warn');return false}if(!d.from){toast('Tanggal mulai wajib diisi','warn');return false}
      await assignSchedule([e.id],+d.ws,d.from);await audit('schedule.assign','schedule_assignments',e.id,null,d);if(d.from<=TODAY())await reprocess(d.from,TODAY(),[e.id]);
      toast('Jadwal kerja tersimpan','ok');m.close();await reopenEmp(e.id)}}]});
};
ACT.empRole=t=>{
  const e=S.ctx.employees.find(x=>x.id===+t.dataset.id);if(!e)return;if(!canRole()){toast('Peran hanya dapat diubah HRD Manager / Super Admin','warn');return}closeOv(t);
  const deps=(S.ctx.departments||[]).filter(d=>d.is_active||d.head_employee_id===e.id),roles=e.roles||['pegawai'];
  const info='<div class="sk-info"><b>Karyawan / Staff</b>: edit data pribadi, absensi masuk-pulang, melihat riwayat, dan mengajukan izin atau cuti, dll (dasar semua akun).<br><b>Manager / Kepala Departemen</b>: memantau tim dan memberikan persetujuan <u>tahap pertama</u> sesuai kewenangan, lalu pengajuan diteruskan ke HRD. Bila departemen belum punya kepala, atau pengaju adalah kepalanya sendiri, pengajuan langsung ke HRD.<br><b>Supervisor / Atasan Langsung</b>: memeriksa dan memberikan persetujuan awal jika alur pengajuan memerlukannya.<br><span class="sk-sub">Peran admin (Super Admin, HRD Manager, HRD Staff / Admin HR) diatur di menu Akun Pengguna; HRD Manager memberi persetujuan akhir.</span></div>';
  modal({title:'Peran - '+e.full_name,size:'wide',body:info+'<form id="rf"><div class="pg-sec">Peran</div><div class="sk-row">'+chk('r_supervisor','Supervisor / Atasan Langsung',roles.includes('supervisor'))+'</div>'
    +'<div class="pg-sec">Manager / Kepala Departemen</div><div class="sk-hint" style="margin-bottom:6px">Centang departemen yang dipimpin pegawai ini. Departemen yang sudah punya kepala lain akan berganti kepala.</div><div class="sk-grid sk-g2">'
    +deps.map(d=>chk('hd'+d.id,d.name+(d.head_employee_id&&d.head_employee_id!==e.id?' (kepala saat ini: '+nameOf(d.head_employee_id)+')':''),d.head_employee_id===e.id)).join('')+'</div></form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{setTimeout(()=>ACT.empView({dataset:{id:String(e.id)}}),0)}},{t:'Simpan',i:'check',keep:true,fn:async m=>{
      const d=formData($('#rf',m.el)),nr=rolesFrom(d),ts=new Date().toISOString();
      const chg=deps.filter(dp=>!!d['hd'+dp.id]!==(dp.head_employee_id===e.id));
      if(chg.some(dp=>d['hd'+dp.id]&&dp.head_employee_id&&dp.head_employee_id!==e.id)&&!await confirmBox('Satu atau lebih departemen akan berganti kepala. Pengajuan baru akan masuk ke kepala yang baru. Lanjutkan?',{ok:'Lanjutkan'}))return false;
      await DB.update('employees',{roles:nr,updated_at:ts},{id:e.id});
      for(const dp of chg)await DB.update('departments',{head_employee_id:d['hd'+dp.id]?e.id:null},{id:dp.id});
      await audit('employee.roles','employees',e.id,{roles:e.roles,kepala:deps.filter(x=>x.head_employee_id===e.id).map(x=>x.name)},{roles:nr,kepala:deps.filter(x=>d['hd'+x.id]).map(x=>x.name)});
      toast('Peran tersimpan','ok');m.close();await reopenEmp(e.id)}}]});
};
function pgToolbar(g,c,divs){
  const pl=c.employees.filter(e=>g.act===''||(g.act==='1')===!!e.is_active),eoD=efOpts(pl,'').depts,eoV=efOpts(pl,g.dept).divs;
  const dirty=['q','dept','div','gender'].some(k=>g[k]);
  return'<div class="pg-bar pg-bar2"><div class="pg-search">'+ic('search','sm')+'<input class="sk-in" type="search" placeholder="Cari nama, NIP, jabatan" title="Cari juga berdasarkan pendidikan, BPJS, email, dan HP" value="'+esc(g.q)+'" data-inp="pgQ" aria-label="Cari pegawai"></div>'
   +'<div class="pg-filters">'+sel('dept',[['','Departemen'],...eoD],g.dept,{attr:'data-chg="pgF" aria-label="Filter departemen"'})+sel('div',[['','Divisi'],...eoV],g.div,{attr:'data-chg="pgF" aria-label="Filter divisi"'})
   +sel('gender',[['','Jenis Kelamin'],['L','Laki-laki'],['P','Perempuan']],g.gender,{attr:'data-chg="pgF" aria-label="Filter jenis kelamin"'})+sel('lim',[['12','Tampil 12'],['24','Tampil 24'],['48','Tampil 48'],['96','Tampil 96'],['all','Tampil semua']],g.per>=1e9?'all':String(g.per),{attr:'data-chg="pgF" aria-label="Jumlah data ditampilkan"'})+(dirty?'<button class="sk-b ghost sm" data-act="pgReset">'+ic('x','sm')+'Reset</button>':'')+'</div>'
   +'<div class="pg-seg" role="group" aria-label="Tampilan"><button data-act="pgView" data-k="card" class="'+(g.view==='card'?'on':'')+'">'+ic('users','sm')+'Kartu</button><button data-act="pgView" data-k="table" class="'+(g.view==='table'?'on':'')+'">'+ic('list','sm')+'Tabel</button><button data-act="pgView" data-k="full" class="'+(g.view==='full'?'on':'')+'">'+ic('book','sm')+'Lengkap</button></div></div>'}
/* foto: potong tengah ke 3:4 (condong ke atas agar wajah tidak terpotong), perkecil, simpan sebagai JPEG */
async function pgCrop34(f){
  if(!/^image\/(jpeg|png|webp)$/.test(f.type))throw new Error('Format foto harus JPG, PNG, atau WebP');
  if(f.size>12*1024*1024)throw new Error('Ukuran foto terlalu besar (maks. 12 MB sebelum dikompres)');
  const url=URL.createObjectURL(f);
  try{
    const img=await new Promise((ok,no)=>{const i=new Image();i.onload=()=>ok(i);i.onerror=()=>no(new Error('Foto tidak dapat dibaca'));i.src=url});
    const w=img.naturalWidth,h=img.naturalHeight;let cw=w,ch=h;
    if(w/h>3/4)cw=Math.round(h*3/4);else ch=Math.round(w*4/3);
    const sx=Math.round((w-cw)/2),sy=Math.round((h-ch)*.2),tw=Math.min(480,cw),th=Math.round(tw*4/3);
    const cv=document.createElement('canvas');cv.width=tw;cv.height=th;const cx=cv.getContext('2d');
    cx.fillStyle='#fff';cx.fillRect(0,0,tw,th);cx.imageSmoothingQuality='high';cx.drawImage(img,sx,sy,cw,ch,0,0,tw,th);
    const blob=await new Promise(ok=>cv.toBlob(ok,'image/jpeg',.86));
    if(!blob)throw new Error('Foto gagal diproses');
    return{blob,dataUrl:cv.toDataURL('image/jpeg',.86)}
  }finally{URL.revokeObjectURL(url)}
}
const PG_BUCKET='employee-photos';
const pgPathOf=u=>{const k=String(u||'').split('/'+PG_BUCKET+'/')[1];return k?decodeURIComponent(k.split('?')[0]):null};
function pgPhotoBox(e){
  const ext=e&&e.photo_url&&!pgPathOf(e.photo_url)&&!/^data:/i.test(e.photo_url)?e.photo_url:'';
  return'<div class="pg-fphoto"><div class="pg-fprev" id="pg-prev"></div><div class="pg-fbtn"><div class="row"><button type="button" class="sk-b sec sm" id="pg-pick">'+ic('upload','sm')+'Pilih foto</button><button type="button" class="sk-b sec sm sk-hide" id="pg-rm">'+ic('trash','sm')+'Hapus foto</button></div>'
   +'<input type="file" id="pg-file" class="sk-hide" accept="image/jpeg,image/png,image/webp"><div class="sk-sub">Foto dipotong otomatis ke rasio 3:4 dan dikompres. Format JPG, PNG, atau WebP.</div>'
   +'<input class="sk-in" id="pg-url" name="photo_url_in" type="text" inputmode="url" autocomplete="off" placeholder="atau tempel URL pas foto / tautan Google Drive" value="'+esc(ext)+'">'
   +'<div class="sk-sub">Tautan Google Drive otomatis diubah ke format gambar. Pastikan berkas dibagikan sebagai &quot;Siapa saja yang memiliki link&quot;.</div></div></div>'}
function pgPhotoInit(m,e){
  const st=m._photo={blob:null,dataUrl:null,removed:false,url:''},prev=$('#pg-prev',m.el),file=$('#pg-file',m.el),rm=$('#pg-rm',m.el),nm=$('input[name=full_name]',m.el),url=$('#pg-url',m.el);
  const store=e&&e.photo_url&&pgPathOf(e.photo_url)?e.photo_url:'';
  const show=u=>{prev.innerHTML='<span class="pg-init">'+(esc(initials(pgNm({full_name:nm?nm.value:''})))||'+')+'</span>'+(u?'<img src="'+esc(u)+'" alt="Pratinjau foto" referrerpolicy="no-referrer" data-pgf="1">':'');rm.classList.toggle('sk-hide',!u)};
  if(url&&url.value.trim())st.url=gdriveUrl(url.value);
  show(st.url||(e&&e.photo_url)||'');
  $('#pg-pick',m.el).onclick=()=>file.click();
  file.onchange=async ev=>{const f=ev.target.files[0];ev.target.value='';if(!f)return;
    try{const r=await pgCrop34(f);st.blob=r.blob;st.dataUrl=r.dataUrl;st.removed=false;st.url='';if(url)url.value='';show(r.dataUrl)}catch(er){toast(errMsg(er),'warn')}};
  rm.onclick=()=>{st.blob=null;st.dataUrl=null;st.removed=true;st.url='';if(url)url.value='';show('')};
  if(url)url.oninput=()=>{const v=gdriveUrl(url.value);st.url=v;st.blob=null;st.dataUrl=null;st.removed=!v&&!store&&!!(e&&e.photo_url);show(v||store)};
  if(nm)nm.addEventListener('input',()=>{if(!st.blob&&!st.url&&!(e&&e.photo_url&&!st.removed))show('')});
}
/* hasil: URL foto yang harus disimpan di kolom photo_url (null = tanpa foto) */
async function pgPhotoSave(m,e,no){
  const st=m._photo||{};
  if(st.blob){
    if(IS_PREVIEW)return st.dataUrl;
    const sb=await ensureSB(),path='p/'+String(no).toLowerCase().replace(/[^a-z0-9]+/g,'-')+'-'+Date.now().toString(36)+'.jpg';
    const up=await sb.storage.from(PG_BUCKET).upload(path,st.blob,{contentType:'image/jpeg',cacheControl:'31536000',upsert:false});
    if(up.error)throw new Error('Gagal mengunggah foto: '+up.error.message);
    return sb.storage.from(PG_BUCKET).getPublicUrl(path).data.publicUrl}
  if(st.url){if(!photoUrlOk(st.url))throw new Error('URL foto harus diawali https:// atau berupa tautan Google Drive');return st.url}
  if(st.removed)return null;
  return e&&e.photo_url||null}
/* hapus berkas foto lama dari Storage setelah data tersimpan (best-effort) */
async function pgPhotoClean(oldUrl,keepUrl){
  if(IS_PREVIEW||!oldUrl||oldUrl===keepUrl)return;const k=pgPathOf(oldUrl);if(!k)return;
  try{const sb=await ensureSB();await sb.storage.from(PG_BUCKET).remove([k])}catch(er){console.warn('hapus foto lama gagal',er)}}

S.pg.sel=new Set();
/* pilih (centang) pegawai untuk ubah / hapus massal */
const pgSelH=()=>{const ids=S._pgIds||[],all=ids.length&&ids.every(x=>S.pg.sel.has(x));return{t:'<input type="checkbox" class="pg-ck" data-chg="pgSelAll"'+(all?' checked':'')+' aria-label="Pilih semua hasil filter" title="Pilih semua hasil filter">',c:'ck'}};
const pgCk=e=>'<input type="checkbox" class="pg-ck" data-chg="pgSel" data-id="'+e.id+'"'+(S.pg.sel.has(e.id)?' checked':'')+' aria-label="Pilih '+esc(e.full_name)+'">';
function pgBulkBar(){
  if(S.pg.view==='card')return'';
  const n=S.pg.sel.size,tot=(S._pgIds||[]).length;
  if(!n)return'<div class="pg-bulk"><span class="sk-sub">Centang pegawai pada tabel untuk mengubah atau menghapus sekaligus. Kotak di judul kolom memilih semua '+tot+' hasil filter.</span></div>';
  return'<div class="pg-bulk on"><b>'+n+' pegawai dipilih</b><span class="sp"></span><button class="sk-b sm" data-act="pgBulkEdit">'+ic('edit','sm')+'Ubah massal</button>'
    +(canRole()?'<button class="sk-b sec sm" data-act="pgBulkOff">Nonaktifkan</button>':'')
    +(isSysAdmin()?'<button class="sk-b bad sm" data-act="pgBulkDel">'+ic('trash','sm')+'Hapus</button>':'')
    +'<button class="sk-b ghost sm" data-act="pgSelClear">Batal pilih</button></div>'}
function pgSelSync(){
  const w=$('#pg-bulk-w');if(w)w.innerHTML=pgBulkBar();
  const ids=S._pgIds||[],all=ids.length&&ids.every(x=>S.pg.sel.has(x)),some=ids.some(x=>S.pg.sel.has(x));
  $$('.pg-ck[data-chg=pgSelAll]').forEach(h=>{h.checked=!!all;h.indeterminate=!all&&some});
  $$('.pg-ck[data-id]').forEach(k=>{k.checked=S.pg.sel.has(+k.dataset.id)})}
CHG.pgSel=t=>{const id=+t.dataset.id;if(t.checked)S.pg.sel.add(id);else S.pg.sel.delete(id);pgSelSync()};
CHG.pgSelAll=t=>{(S._pgIds||[]).forEach(id=>{if(t.checked)S.pg.sel.add(id);else S.pg.sel.delete(id)});pgSelSync()};
ACT.pgSelClear=()=>{S.pg.sel.clear();pgSelSync()};
const pgSelEmps=()=>S.ctx.employees.filter(e=>S.pg.sel.has(e.id));
async function pgBulkRun(list,fn){let fail=0;for(const part of chunk(list,8))await Promise.all(part.map(async e=>{try{await fn(e)}catch(er){fail++;console.warn('aksi massal gagal',er)}}));return fail}
ACT.pgBulkEdit=()=>{
  const list=pgSelEmps();if(!list.length)return;const nb='- tidak diubah -';
  const body='<div class="sk-info">Hanya isian yang diisi yang diubah untuk '+list.length+' pegawai terpilih. Isian kosong dibiarkan seperti semula.</div><form id="bef"><div class="sk-grid sk-g2">'
    +fld('Departemen',sel('department_id',deptOptions(),'',{blank:nb}))+fld('Divisi',inp('division','',{ph:nb}))+fld('Jabatan',inp('position','',{ph:nb}))
    +fld('Tipe',sel('type',[['teacher','Guru'],['staff','Karyawan Umum']],'',{blank:nb}))+fld('Status karyawan',sel('employment_status',Object.entries(EMPST),'',{blank:nb}))
    +(canRole()?fld('Status akun',sel('is_active',[['1','Aktif'],['0','Nonaktif']],'',{blank:nb})):'')+'</div></form>';
  modal({title:'Ubah '+list.length+' Pegawai',size:'wide',body,actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Terapkan',i:'check',keep:true,fn:async m=>{
    const d=formData($('#bef',m.el)),up={};
    if(d.department_id)up.department_id=+d.department_id;
    if(normSp(d.division))up.division=normSp(d.division);
    if(normSp(d.position))up.position=normSp(d.position);
    if(d.type)up.type=d.type;if(d.employment_status)up.employment_status=d.employment_status;
    if(canRole()&&d.is_active)up.is_active=d.is_active==='1';
    if(!Object.keys(up).length){toast('Belum ada isian yang diubah','warn');return false}
    up.updated_at=new Date().toISOString();
    const fail=await pgBulkRun(list,e=>DB.update('employees',up,{id:e.id}));
    await audit('employee.bulk_update','employees','bulk',null,{jumlah:list.length,...up});
    toast(fail?(list.length-fail)+' pegawai diubah, '+fail+' gagal':list.length+' pegawai berhasil diubah',fail?'warn':'ok');
    S.pg.sel.clear();m.close();Ctx.reset();render(true)}}]});
};
ACT.pgBulkOff=async()=>{
  const list=pgSelEmps().filter(e=>e.is_active);if(!list.length){toast('Pegawai terpilih sudah nonaktif semua','warn');return}
  if(!await confirmBox('Nonaktifkan '+list.length+' pegawai terpilih? Akun mereka tidak bisa masuk lagi, data tetap tersimpan.',{ok:'Nonaktifkan'}))return;
  const fail=await pgBulkRun(list,e=>DB.update('employees',{is_active:false,updated_at:new Date().toISOString()},{id:e.id}));
  await audit('employee.bulk_deactivate','employees','bulk',null,{jumlah:list.length});
  toast(fail?(list.length-fail)+' dinonaktifkan, '+fail+' gagal':list.length+' pegawai dinonaktifkan',fail?'warn':'ok');
  S.pg.sel.clear();Ctx.reset();await render(true)};
ACT.pgBulkDel=async()=>{
  if(!isSysAdmin())return;const list=pgSelEmps();if(!list.length)return;
  if(!await confirmBox('Hapus permanen '+list.length+' pegawai terpilih beserta SELURUH datanya (rekap, pengajuan, log)? Tindakan ini tidak dapat dibatalkan. Untuk pegawai yang sudah berhenti, gunakan Nonaktifkan.',{danger:true,ok:'Hapus '+list.length+' pegawai'}))return;
  const fail=await pgBulkRun(list,async e=>{await DB.remove('employees',{id:e.id});await pgPhotoClean(e.photo_url,null)});
  await audit('employee.bulk_delete','employees','bulk',{jumlah:list.length,nama:list.slice(0,20).map(e=>e.full_name)},null);
  toast(fail?(list.length-fail)+' dihapus, '+fail+' gagal':list.length+' pegawai dihapus',fail?'warn':'ok');
  S.pg.sel.clear();Ctx.reset();await render(true)};
reg('pegawai',{
  range:()=>[TODAY(),TODAY()],
  render(){
    const c=S.ctx,g=S.pg,tabs=stabs([['pegawai','Pegawai ('+c.employees.filter(e=>e.is_active).length+')'],['div','Divisi ('+pgDivCount(c)+')'],['subs','Pengganti Kepala Dept.'],['fp','ID Mesin Fingerprint']],g.tab,'pgTab');
    if(g.tab==='div'){
      const act=c.employees.filter(e=>e.is_active),hasD=e=>e.department_id!=null&&e.department_id!==''&&c.departments.some(d=>d.id===e.department_id);
      const noDept=act.filter(e=>!hasD(e)),noHead=c.departments.filter(d=>d.is_active&&!d.head_employee_id).length,dd=String(g.ddept||'');
      const groupDiv=m=>{const mp=new Map();m.forEach(e=>{const v=efDiv(e),k=v.toLowerCase();if(!mp.has(k))mp.set(k,{n:v,l:[]});mp.get(k).l.push(e)});return[...mp.values()].sort((x,y)=>(x.n===''?1:0)-(y.n===''?1:0)||x.n.localeCompare(y.n,'id'))};
      const dvCard=(d,x)=>{const m=x.l.slice().sort((p,q)=>p.full_name.localeCompare(q.full_name,'id')),gr=m.filter(e=>e.type==='teacher').length,lk=m.filter(e=>e.gender==='L').length,pr=m.filter(e=>e.gender==='P').length,show=m.slice(0,5),rest=m.length-show.length;
        return'<article class="pg-ds'+(x.n?'':' none')+'"><div class="pg-ds-t"><b title="'+esc(x.n||'Tanpa divisi')+'">'+esc(x.n||'Tanpa divisi')+'</b><span class="pg-ds-n">'+m.length+'</span></div>'
         +'<div class="pg-ds-m">'+gr+' guru · '+(m.length-gr)+' karyawan · '+lk+' L / '+pr+' P</div>'
         +'<div class="pg-stack" title="Anggota divisi">'+show.map(y=>pgAvatar(y)).join('')+(rest>0?'<span class="more">+'+rest+'</span>':'')+'</div>'
         +'<button class="sk-b sec sm" data-act="pgDivGo" data-d="'+(d?d.id:EF_NONE)+'" data-v="'+esc(x.n||EF_NONE)+'">'+ic('users','sm')+'Lihat pegawai</button></article>'};
      const sec=(d,m)=>{const divs=groupDiv(m),nd=divs.filter(x=>x.n).length,h=d&&d.head_employee_id?c.employees.find(e=>e.id===d.head_employee_id):null;
        return'<section class="pg-dsec'+(d&&!d.is_active?' off':'')+'"><header class="pg-dsh"><span class="sk-av pg-av d">'+ic('building')+'</span><div class="pg-id"><b>'+esc(d?d.name:'Tanpa Departemen')+'</b><span>'+nd+' divisi · '+m.length+' pegawai aktif'+(d&&!d.is_active?' · nonaktif':'')+'</span></div>'
         +(d?'<div class="pg-dsk">'+(h?'<span class="pg-dhd" title="Kepala Departemen">'+pgAvatar(h)+'<span>'+esc(empName(h))+'<br><small class="sk-sub">Kepala Departemen</small></span></span>':bd('Belum ada kepala','a'))+'</div><div class="pg-act"><button class="sk-b sec sm" data-act="deptEdit" data-id="'+d.id+'">'+ic('edit','sm')+'Ubah</button><button class="sk-b sec sm" data-act="deptDel" data-id="'+d.id+'" aria-label="Hapus '+esc(d.name)+'">'+ic('trash','sm')+'</button></div>':'')+'</header>'
         +(divs.length?'<div class="pg-dsgrid">'+divs.map(x=>dvCard(d,x)).join('')+'</div>':'<div class="sk-empty" style="padding:14px">Belum ada pegawai aktif di departemen ini</div>')+'</section>'};
      const deps=c.departments.slice().sort((x,y)=>(y.is_active-x.is_active)||x.name.localeCompare(y.name,'id')).filter(d=>!dd||String(d.id)===dd);
      let body=deps.map(d=>sec(d,act.filter(e=>e.department_id===d.id))).join('');
      if(noDept.length&&(!dd||dd===EF_NONE))body+=sec(null,noDept);
      if(!body)body='<div class="sk-empty">Belum ada departemen</div>';
      const ddOpts=[['','Semua departemen'],...c.departments.slice().sort((x,y)=>x.name.localeCompare(y.name,'id')).map(d=>[String(d.id),d.name]),...(noDept.length?[[EF_NONE,'Tanpa departemen']]:[])];
      return tabs+'<div class="sk-kpis">'+kpi('building','#1d4ed8',c.departments.filter(d=>d.is_active).length,'Departemen aktif')+kpi('grid','#4f46e5',pgDivCount(c),'Divisi')+kpi('users','#059669',act.length,'Pegawai aktif')+kpi('alert',noHead?'#d97706':'#64748b',noHead,'Belum ada kepala')+'</div>'
       +(noDept.length?'<div class="sk-warn" style="margin-bottom:12px">'+noDept.length+' pegawai aktif belum memiliki departemen. Lengkapi lewat tombol Ubah pada data pegawai.</div>':'')
       +card('Divisi per Departemen ('+pgDivCount(c)+')',body,'<div class="sk-tool">'+sel('ddept',ddOpts,dd,{attr:'data-chg="pgDD" aria-label="Filter departemen"'})+'<button class="sk-b" data-act="deptEdit">'+ic('plus','sm')+'Tambah Departemen</button></div>')
       +'<div class="sk-info">Daftar divisi mengikuti data pegawai di database (kolom Departemen dan Divisi). Untuk menambah atau memindahkan divisi, ubah data pegawai yang bersangkutan. Kepala Departemen otomatis menjadi penyetuju tahap 1 untuk seluruh pegawai di departemennya.</div>';
    }
    if(g.tab==='subs'){
      const rows=(c.substitutes||[]).map(s=>[esc(deptName(s.department_id)),esc(nameOf(s.substitute_employee_id)),fmtD(s.start_date),fmtD(s.end_date),(s.start_date<=TODAY()&&s.end_date>=TODAY())?bd('Berlaku','g'):(s.end_date<TODAY()?bd('Selesai','s'):bd('Akan datang','b')),'<button class="sk-b sec sm" data-act="subDel" data-id="'+s.id+'">'+ic('trash','sm')+'</button>']);
      return tabs+card('Pengganti Kepala Departemen (saat berhalangan)',tbl(['Departemen','Pengganti','Mulai','Selesai','Status',''],rows,{empty:'Belum ada pengganti'}),'<button class="sk-b" data-act="subAdd">'+ic('plus','sm')+'Tambah Pengganti</button>')+'<div class="sk-info">Selama periode berlaku, pengganti dapat memberi persetujuan tahap 1 untuk departemen tersebut.</div>';
    }
    if(g.tab==='fp'){
      const fpm={};(c.fingerprints||[]).forEach(f=>fpm[f.employee_id]=f);
      efFix('fp',c.employees.filter(e=>e.is_active));const rows=c.employees.filter(e=>e.is_active&&efPass('fp',e)).map(e=>{const f=fpm[e.id];return[esc(e.full_name),esc(e.employee_no),esc(deptName(e.department_id)),'<code>'+esc(f?f.fingerprint_id:e.employee_no)+'</code>'+((f&&f.fingerprint_id!==e.employee_no)?' '+bd('Beda dari NIP','a'):''),'<button class="sk-b sec sm" data-act="fpEdit" data-e="'+e.id+'">'+ic('edit','sm')+'Ubah</button>']});
      const un=(c.fingerprints||[]).length;
      return tabs+card('ID pegawai di mesin fingerprint',tbl(['Pegawai','NIP','Departemen','ID di mesin',''],rows,{maxh:'65vh',ef:'fp'}),'<div class="sk-tool">'+efBar('fp',c.employees.filter(e=>e.is_active),{inl:1})+'<span class="sk-sub">'+un+' pegawai sudah dipetakan</span></div>')+'<div class="sk-info">ID ini dipakai saat impor log fingerprint. Log dengan ID yang belum dipetakan disimpan sebagai "tidak cocok" dan dicocokkan otomatis setelah ID diisi di sini.</div>';
    }
    const q=g.q.toLowerCase(),divs=[...new Set(c.employees.map(e=>(e.division||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id'));
    const list=c.employees.filter(e=>(g.act===''||(g.act==='1')===!!e.is_active)&&efDeptOk(e,g.dept)&&(!g.gender||e.gender===g.gender)&&(!g.type||e.type===g.type)&&(!g.st||e.employment_status===g.st)&&efDivOk(e,g.div)&&(!q||[e.full_name,e.title_suffix,e.employee_no,e.nip,e.position,e.division,e.email,e.phone,e.birth_place,e.education_level,e.education_institution,e.education_major,e.bpjs_kesehatan,e.bpjs_ketenagakerjaan].some(v=>(v||'').toLowerCase().includes(q)))).sort((a,b)=>a.full_name.localeCompare(b.full_name,'id'));
    if(g.view==='card')list.sort(pgOrd);
    S._pgIds=list.map(e=>e.id);{const have=new Set(c.employees.map(e=>e.id));[...g.sel].forEach(x=>{if(!have.has(x))g.sel.delete(x)})}
    const act=c.employees.filter(e=>e.is_active),shown=list.slice(0,g.lim);
    let body;
    if(!list.length)body='<div class="sk-empty">Tidak ada pegawai yang cocok dengan pencarian atau filter.</div>';
    else if(g.view==='full')body=tbl([pgSelH(),...PG_FULL_H],shown.map((e,i)=>[pgCk(e),...pgFullRow(e,i)]),{maxh:'66vh',cls:'pg-t pg-full'});
    else if(g.view==='table')body=tbl([pgSelH(),'Pegawai','Departemen & jabatan','Kepegawaian','Peran','Kontak',''],shown.map(e=>[pgCk(e),'<div class="pg-cell fx-click" data-act="empView" data-id="'+e.id+'" title="Lihat detail"><span class="pg-th'+(e.is_active?'':' off')+'"><span class="pg-init">'+pgInit(e)+'</span>'+pgImg(e,'')+'</span><div><b>'+esc(empName(e))+'</b><div class="sk-sub">'+esc([e.nip,e.employee_no].filter(Boolean).join(' · '))+'</div></div></div>',
        '<b>'+esc(e.position||'-')+'</b><div class="sk-sub">'+esc(pgWhere(e)||'-')+'</div>',pgStatBd(e)+' '+bd(TYPE[e.type]||e.type,e.type==='teacher'?'i':'c')+(e.is_active?'':' '+bd('Nonaktif','r'))+'<div class="sk-sub">'+esc(pgTenure(e))+'</div>'+(e.contract_end_date?'<div class="sk-sub">'+pgContract(e)+'</div>':''),
        pgRoles(e)||'<span class="sk-sub">Pegawai</span>',[pgMail(e),pgTel(e)].filter(Boolean).join('<br>')||'<span class="sk-sub">-</span>',
        '<button class="sk-b sec sm" data-act="empEdit" data-id="'+e.id+'">'+ic('edit','sm')+'Ubah</button>']),{maxh:'66vh',cls:'pg-t'});
    else body=pgGroups(shown);
    const more=list.length>shown.length?'<div class="pg-more"><button class="sk-b sec" data-act="pgMore">Tampilkan '+Math.min(g.per,list.length-shown.length)+' lagi</button><span class="sk-sub">'+shown.length+' dari '+list.length+' pegawai</span></div>':'';
    return tabs+pgNotice(act)+'<div class="sk-kpis pg-kpis">'+kpi('users','#2563eb',act.length,'Pegawai aktif','dari '+c.employees.length+' terdaftar')+kpi('book','#4f46e5',act.filter(e=>e.type==='teacher').length,'Guru')+kpi('brief','#0891b2',act.filter(e=>e.type==='staff').length,'Karyawan umum')+kpi('building','#059669',c.departments.filter(d=>d.is_active).length,'Departemen aktif')+'</div>'
     +card('Data Karyawan ('+list.length+')',pgToolbar(g,c,divs)+'<div id="pg-bulk-w">'+pgBulkBar()+'</div>'+body+more,'<div class="sk-row"><button class="sk-b sec sm" data-act="empImport">'+ic('upload','sm')+'Impor</button><button class="sk-b sec sm" data-act="empExport">'+ic('download','sm')+'Ekspor</button><button class="sk-b sm" data-act="empEdit">'+ic('plus','sm')+'Tambah Pegawai</button></div>');
  }
});
ACT.pgTab=async t=>{S.pg.tab=t.dataset.k;await render()};
CHG.pgDD=async t=>{S.pg.ddept=t.value;await render()};
ACT.pgDivGo=async t=>{const d=t.dataset.d,v=t.dataset.v;Object.assign(S.pg,{tab:'pegawai',q:'',act:'1',type:'',st:'',gender:'',dept:String(d),div:v,lim:S.pg.per});await render()};
CHG.pgF=async t=>{if(t.name==='lim')S.pg.per=t.value==='all'?1e9:(+t.value||PG_STEP);else S.pg[t.name]=t.value;S.pg.lim=S.pg.per;if(t.name==='dept'||t.name==='act'){const pl=S.ctx.employees.filter(e=>S.pg.act===''||(S.pg.act==='1')===!!e.is_active);efValid(S.pg,pl)}await render()};
ACT.pgView=async t=>{S.pg.view=t.dataset.k;await render()};
ACT.pgMore=async()=>{S.pg.lim+=S.pg.per;await render()};
ACT.pgReset=async()=>{Object.assign(S.pg,{q:'',dept:'',type:'',act:'1',st:'',div:'',gender:'',lim:S.pg.per});await render()};
INP.pgQ=debounce(async t=>{S.pg.q=t.value;S.pg.lim=S.pg.per;const p=t.selectionStart;await render();const i=$('[data-inp=pgQ]');if(i){i.focus();try{i.setSelectionRange(p,p)}catch(e){}}},400);
ACT.empExport=()=>{const c=S.ctx,iso=v=>v?String(v).slice(0,10):'';exportRows('data-karyawan','Karyawan',['ID','NIP','Nama Lengkap','Gelar','Jenis Kelamin','Tanggal Bergabung','Tempat Lahir','Tanggal Lahir','Tanggal Akhir Kontrak','Departemen','Divisi','Jabatan','Pendidikan Terakhir','Nama Pendidikan','Jurusan','BPJS Kesehatan','BPJS Ketenagakerjaan','Status Karyawan','No. Handphone','Email','Status Pernikahan','URL Pas Photo','Masa Kerja','Tipe','Peran','Status Akun'],c.employees.slice().sort((a,b)=>a.full_name.localeCompare(b.full_name,'id')).map(e=>[e.id,String(e.employee_no),e.full_name,e.title_suffix||'',GENDER[e.gender]||'',iso(e.join_date),e.birth_place||'',iso(e.birth_date),iso(e.contract_end_date),deptName(e.department_id),e.division||'',e.position||'',e.education_level||'',e.education_institution||'',e.education_major||'',e.bpjs_kesehatan||'',e.bpjs_ketenagakerjaan||'',EMPST[e.employment_status]||'',e.phone||'',e.email||'',MARITAL[e.marital_status]||'',/^https?:/i.test(e.photo_url||'')?e.photo_url:'',tenure(e.join_date,e.end_date),TYPE[e.type]||'',(e.roles||[]).map(r=>ROLE[r]||r).join(', '),e.is_active?'Aktif':'Nonaktif']))};
function empForm(e){
  const roles=e?e.roles:['pegawai'],v=k=>e?e[k]||'':'',dt=k=>e&&e[k]?String(e[k]).slice(0,10):'',H=t=>'<div class="pg-sec">'+t+'</div>';
  return'<form id="ef">'+pgPhotoBox(e)
   +H('Data pribadi')+'<div class="sk-grid sk-g2">'
   +fld('Nama lengkap',inp('full_name',v('full_name'),{req:1}))
   +fld('Gelar',inp('title_suffix',v('title_suffix'),{ph:'mis. S.Pd, M.Pd'}),'Gelar belakang; gelar depan (Dr., Drs.) ditulis di nama')
   +fld('Jenis kelamin',sel('gender',Object.entries(GENDER),v('gender'),{blank:'- pilih -'}))
   +fld('Status pernikahan',sel('marital_status',Object.entries(MARITAL),v('marital_status'),{blank:'- pilih -'}))
   +fld('Tempat lahir',inp('birth_place',v('birth_place')))
   +fld('Tanggal lahir',inp('birth_date',dt('birth_date'),{type:'date'}))
   +fld('No. Handphone',inp('phone',v('phone'),{ph:'08xxxxxxxxxx'}))
   +fld('Email (untuk login)',inp('email',v('email'),{type:'email',req:1,ph:'nama@fiwa.sch.id',attr:'autocapitalize="none"'}),'Pegawai masuk ke aplikasi dengan email ini')+'</div>'
   +H('Kepegawaian')+'<div class="sk-grid sk-g2">'
   +fld('NIP (Nomor Induk)',inp('employee_no',v('employee_no'),{req:1,dis:!!e}),'Sama dengan ID mesin fingerprint')
   +fld('Tanggal bergabung',inp('join_date',e?String(e.join_date).slice(0,10):TODAY(),{type:'date'}))
   +fld('Status karyawan',sel('employment_status',Object.entries(EMPST),e?e.employment_status:'permanent'))
   +fld('Tanggal akhir kontrak',inp('contract_end_date',dt('contract_end_date'),{type:'date'}),'Isi untuk status Kontrak; sistem mengingatkan 30 hari sebelum berakhir')
   +fld('Departemen',sel('department_id',deptOptions(),e?e.department_id:'',{blank:'- pilih -'}))
   +fld('Divisi',inp('division',v('division'),{ph:'mis. Kurikulum'}))
   +fld('Jabatan',inp('position',v('position')))
   +fld('Tipe',sel('type',[['teacher','Guru'],['staff','Karyawan Umum']],e?e.type:'staff'))
   +(e?fld('Tanggal berhenti (bila ada)',inp('end_date',dt('end_date'),{type:'date'})):fld('Kata sandi awal',inp('password','',{ph:'kosong = NIP'}),'Pegawai wajib menggantinya saat login pertama'))+'</div>'
   +H('Pendidikan')+'<div class="sk-grid sk-g2">'
   +fld('Pendidikan terakhir',sel('education_level',EDU.map(x=>[x,x]),v('education_level'),{blank:'- pilih -'}))
   +fld('Nama pendidikan',inp('education_institution',v('education_institution'),{ph:'Nama sekolah / kampus'}))
   +fld('Jurusan',inp('education_major',v('education_major'),{ph:'mis. Pendidikan Matematika'}))+'</div>'
   +H('BPJS')+'<div class="sk-grid sk-g2">'
   +fld('BPJS Kesehatan',inp('bpjs_kesehatan',v('bpjs_kesehatan'),{ph:'Nomor kartu (angka)',attr:'inputmode="numeric" autocomplete="off"'}))
   +fld('BPJS Ketenagakerjaan',inp('bpjs_ketenagakerjaan',v('bpjs_ketenagakerjaan'),{ph:'Nomor kartu (angka)',attr:'inputmode="numeric" autocomplete="off"'}))+'</div>'
   +H('Peran & akun')+'<div class="sk-row">'+chk('r_supervisor','Supervisor / Atasan Langsung',roles.includes('supervisor'),canRole()?'':'disabled')+(e?chk('is_active','Akun aktif',e.is_active,canRole()?'':'disabled'):'')+'</div>'
   +'<div class="sk-hint">Manager / Kepala Departemen ditentukan dari tab Divisi (tombol Ubah pada departemen).'+(canRole()?'':' Peran & status akun hanya dapat diubah HRD Manager / Super Admin.')+'</div></form>';
}
const rolesFrom=d=>['pegawai',...(d.r_supervisor?['supervisor']:[])];
ACT.empEdit=t=>{
  const e=t.dataset.id?S.ctx.employees.find(x=>x.id===+t.dataset.id):null;
  modal({title:e?'Ubah Pegawai':'Tambah Pegawai',size:'wide',body:empForm(e),onMount:m=>pgPhotoInit(m,e),actions:[
    ...(e&&isSysAdmin()?[{t:'Hapus permanen',c:'bad',i:'trash',keep:true,fn:async m=>{if(!await confirmBox('Hapus '+e.full_name+' beserta SELURUH datanya (rekap, pengajuan, log)? Tindakan ini tidak dapat dibatalkan. Untuk pegawai yang sudah berhenti, gunakan Nonaktifkan.',{danger:true,ok:'Hapus permanen'}))return false;await DB.remove('employees',{id:e.id});await pgPhotoClean(e.photo_url,null);await audit('employee.delete','employees',e.id,e,null);toast('Pegawai dihapus','ok');m.close();Ctx.reset();render(true)}}]:[]),
    {t:'Batal',c:'sec',fn:()=>{}},
    {t:'Simpan',i:'check',fn:async m=>{
      const d=formData($('#ef',m.el));
      if(!normSp(d.full_name)){toast('Nama wajib diisi','warn');return false}
      for(const k of ['title_suffix','phone','email','division','position'])d[k]=String(d[k]||'').trim();
      normEmp(d);if(!d.join_date){toast('Tanggal bergabung wajib diisi','warn');return false}
      if(!d.email){toast('Email wajib diisi (dipakai untuk login)','warn');return false}
      {const v=empValidate(d);if(v){toast(v,'warn');return false}
       const o=S.ctx.employees.filter(x=>!e||x.id!==e.id);
       d.employee_no=String(d.employee_no||(e?e.employee_no:'')).trim();d.nip=d.employee_no;
       if(!e&&o.some(x=>String(x.employee_no).toLowerCase()===d.employee_no.toLowerCase())){toast('NIP sudah dipakai pegawai lain','warn');return false}
       if(o.some(x=>(x.email||'').toLowerCase()===d.email)){toast('Email sudah dipakai pegawai lain','warn');return false}}
      if(e&&!canRole()){d.r_supervisor=e.roles.includes('supervisor');d.is_active=e.is_active}
      if(!e&&!canRole()){d.r_supervisor=false}
      if(!e&&!String(d.employee_no).trim()){toast('NIP wajib diisi','warn');return false}
      if(S.legacyEmp&&EMP_X.some(k=>d[k])){toast(LEGACY_MSG,'bad');return false}
      const photo=await pgPhotoSave(m,e,e?e.employee_no:String(d.employee_no).trim());
      if(e){
        const up={full_name:d.full_name,title_suffix:d.title_suffix||null,nip:e.employee_no,division:d.division||null,department_id:d.department_id?+d.department_id:null,position:d.position||null,type:d.type,employment_status:d.employment_status,roles:rolesFrom(d),join_date:d.join_date,end_date:d.end_date||null,phone:d.phone||null,email:d.email,photo_url:photo,...empX(d),is_active:!!d.is_active,updated_at:new Date().toISOString()};
        if(S.legacyEmp)EMP_X.forEach(k=>delete up[k]);
        await DB.update('employees',up,{id:e.id});await audit('employee.update','employees',e.id,{full_name:e.full_name,roles:e.roles,is_active:e.is_active},{roles:up.roles,is_active:up.is_active});
        await pgPhotoClean(e.photo_url,photo);
      }else{
        const row={employee_no:d.employee_no.trim(),full_name:d.full_name,title_suffix:d.title_suffix,nip:d.employee_no.trim(),division:d.division,department_id:d.department_id?+d.department_id:null,position:d.position,type:d.type,employment_status:d.employment_status,roles:rolesFrom(d),join_date:d.join_date,phone:d.phone,email:d.email,photo_url:photo,password:d.password,...empX(d)};const r=await rpc('adm_create_employees',{p_rows:[row]});
        if(r.skipped&&r.skipped.length){await pgPhotoClean(photo,null);toast(r.skipped[0].reason,'bad');return false}
        if(await empXSave([row]))toast('Pegawai dibuat, tetapi sebagian kolom tambahan gagal disimpan. Pastikan migrasi-karyawan.sql sudah dijalankan.','warn');
      }
      toast('Data pegawai tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.empImport=()=>{
  modal({title:'Impor Pegawai (Excel / CSV)',size:'wide',body:'<div class="sk-info">Baris pertama harus berisi judul kolom. <b>Wajib:</b> <b>nip</b>, <b>nama_lengkap</b>, <b>email</b> (dipakai untuk login). <b>Opsional:</b> gelar, jenis_kelamin (L/P), tanggal_bergabung, tempat_lahir, tanggal_lahir, tanggal_akhir_kontrak, departemen, divisi, jabatan, pendidikan_terakhir (SD/SMP/SMA/D3/S1/S2/S3), nama_pendidikan, jurusan, bpjs_kesehatan, bpjs_ketenagakerjaan, status_karyawan (tetap/kontrak/honorer), no_handphone, status_pernikahan, url_pas_photo (boleh tautan Google Drive), tipe (guru/karyawan; bila kosong ditebak dari jabatan), sandi. Kolom <b>id</b> diabaikan (dibuat otomatis). Format kolom nip dan BPJS di Excel harus <i>Teks</i>; tanggal ditulis YYYY-MM-DD atau DD/MM/YYYY. Departemen yang belum ada dibuat otomatis. Sandi kosong = NIP; pegawai wajib menggantinya saat login pertama.</div><div class="sk-row" style="margin-bottom:10px"><button class="sk-b sec sm" data-act="empTpl">'+ic('download','sm')+'Unduh templat CSV</button></div><input type="file" id="ifile" accept=".xlsx,.xls,.csv,.txt" class="sk-in"><div id="iprev" style="margin-top:12px"></div>',
    onMount:m=>{$('#ifile',m.el).onchange=async ev=>{const f=ev.target.files[0];if(!f)return;try{const rows=await readTable(f);if(rows.length<2){$('#iprev',m.el).innerHTML='<div class="sk-err">Berkas kosong atau tidak terbaca</div>';return}
      const h=rows[0].map(normKey),ix=(...n)=>h.findIndex(x=>n.includes(x));
      const noI=(()=>{const i=ix('nip','noinduk','nomorinduk','nik','employeeno');return i>=0?i:ix('id')})();
      const col={no:noI,nm:ix('namalengkap','nama','name','fullname'),tt:ix('gelarbelakang','gelar','titlesuffix'),dv:ix('divisi','division'),dp:ix('departemen','department','bagian','unit'),ps:ix('jabatan','position'),ty:ix('tipe','type','jenis'),
        st:ix('statuskaryawan','statuskepegawaian','status'),jd:ix('tanggalbergabung','tanggalmasuk','tglmasuk','joindate','bergabung'),hp:ix('nohandphone','nohp','hp','telepon','phone'),em:ix('email'),pw:ix('sandi','password','katasandi'),
        gn:ix('jeniskelamin','gender','jk','kelamin'),bp:ix('tempatlahir','birthplace'),bd:ix('tanggallahir','tgllahir','birthdate'),ce:ix('tanggalakhirkontrak','akhirkontrak','kontrakberakhir','contractend'),
        el:ix('pendidikanterakhir','pendidikan','education'),ei:ix('namapendidikan','institusi','namainstitusi','sekolah','kampus'),mj:ix('jurusan','programstudi','prodi'),
        bk:ix('bpjskesehatan','bpjskes'),bt:ix('bpjsketenagakerjaan','bpjstk','bpjsnaker'),ms:ix('statuspernikahan','statusperkawinan','pernikahan'),ph:ix('urlpaspoto','urlpasphoto','urlfoto','paspoto','photourl','photo','foto')};
      if(col.no<0||col.nm<0){$('#iprev',m.el).innerHTML='<div class="sk-err">Kolom nip dan nama (nama_lengkap) tidak ditemukan pada baris judul.</div>';return}
      const g=(r,k)=>col[k]>=0?String(r[col[k]]==null?'':r[col[k]]).trim():'';
      const mapT=v=>/guru|teacher|pengajar/i.test(v)?'teacher':'staff',mapS=v=>/kontrak|contract/i.test(v)?'contract':/honor/i.test(v)?'honorary':'permanent';
      const mapG=v=>/^(l|lk|laki|pria|male|m)\b/i.test(v)?'L':/^(p|pr|perempuan|wanita|female|f)\b/i.test(v)?'P':'';
      const mapM=v=>{v=v.toLowerCase();if(!v)return'';if(/cerai\s*mati|janda|duda|widow/.test(v))return'cerai_mati';if(/cerai|divorc/.test(v))return'cerai_hidup';if(/belum|lajang|single|tidak\s*kawin/.test(v))return'belum_menikah';if(/menikah|kawin|married/.test(v))return'menikah';return''};
      const mapE=v=>{const u=String(v||'').toUpperCase().replace(/[\s.\-]/g,'');if(!u)return'';if(/^(S3|DOKTOR)/.test(u))return'S3';if(/^(S2|MAGISTER)/.test(u))return'S2';if(/^(S1|SARJANA|STRATA1)/.test(u))return'S1';if(/^D4/.test(u))return'D4';if(/^D3|^DIPLOMA3/.test(u))return'D3';if(/^D2/.test(u))return'D2';if(/^D1/.test(u))return'D1';if(/^(SMA|SMK|SMU|SLTA)/.test(u)||/^(MA|MAN)$/.test(u))return'SMA/SMK';if(/^(SMP|MTS|SLTP)/.test(u))return'SMP';if(/^(SD|SEKOLAHDASAR)/.test(u)||u==='MI')return'SD';return''};
      /* tanggal: YYYY-MM-DD, DD/MM/YYYY, m/d/yy (format bawaan Excel), atau nomor seri Excel */
      const impDate=v=>{v=String(v==null?'':v).trim();if(!v)return'';const p2=n=>String(n).padStart(2,'0');let x=v.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);if(x)return x[1]+'-'+p2(x[2])+'-'+p2(x[3]);
        x=v.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);if(x)return x[3]+'-'+p2(x[2])+'-'+p2(x[1]);
        x=v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);if(x)return(+x[3]>40?'19':'20')+x[3]+'-'+p2(x[1])+'-'+p2(x[2]);
        if(/^\d{5}(\.\d+)?$/.test(v)){const n=Math.floor(+v);if(n>20000&&n<80000)return new Date(Date.UTC(1899,11,30)+n*864e5).toISOString().slice(0,10)}return''};
      const digits=v=>String(v).replace(/[\s.\-]/g,'');
      const data=rows.slice(1).map(r=>({employee_no:g(r,'no'),full_name:normSp(g(r,'nm')),title_suffix:normSp(g(r,'tt')).replace(/^[,\s]+|[,\s]+$/g,''),nip:g(r,'no'),division:normSp(g(r,'dv')),department_name:normSp(g(r,'dp')),position:normSp(g(r,'ps')),
        type:g(r,'ty')?mapT(g(r,'ty')):(/guru|pengajar|ustadz|ustadzah|dosen/i.test(g(r,'ps'))?'teacher':'staff'),employment_status:mapS(g(r,'st')),join_date:impDate(g(r,'jd')),phone:g(r,'hp'),email:g(r,'em').toLowerCase(),password:g(r,'pw'),roles:['pegawai'],
        gender:mapG(g(r,'gn')),birth_place:normSp(g(r,'bp')),birth_date:impDate(g(r,'bd')),contract_end_date:impDate(g(r,'ce')),education_level:mapE(g(r,'el')),education_institution:normSp(g(r,'ei')),education_major:normSp(g(r,'mj')),
        bpjs_kesehatan:digits(g(r,'bk')),bpjs_ketenagakerjaan:digits(g(r,'bt')),marital_status:mapM(g(r,'ms')),photo_url:gdriveUrl(g(r,'ph')),
        _raw:{jd:g(r,'jd'),bd:g(r,'bd'),ce:g(r,'ce'),el:g(r,'el'),gn:g(r,'gn'),ms:g(r,'ms')}})).filter(r=>r.employee_no||r.full_name);
      m._data=data;
      const badDate=data.filter(r=>(r._raw.jd&&!r.join_date)||(r._raw.bd&&!r.birth_date)||(r._raw.ce&&!r.contract_end_date)).length,
        badOpt=data.filter(r=>(r._raw.el&&!r.education_level)||(r._raw.gn&&!r.gender)||(r._raw.ms&&!r.marital_status)).length;
      $('#iprev',m.el).innerHTML='<div class="sk-ok">'+data.length+' baris siap diimpor. Pratinjau 5 baris pertama:</div>'+(data.some(r=>!r.email)?'<div class="sk-warn">'+data.filter(r=>!r.email).length+' baris tanpa email akan dilewati (email dipakai untuk login).</div>':'')
        +(badDate?'<div class="sk-warn">'+badDate+' baris memiliki tanggal yang tidak terbaca (kolom dikosongkan). Gunakan format YYYY-MM-DD atau DD/MM/YYYY.</div>':'')
        +(badOpt?'<div class="sk-warn">'+badOpt+' baris memiliki nilai jenis kelamin / pendidikan / status pernikahan yang tidak dikenali (kolom dikosongkan).</div>':'')
        +(data.some(r=>/^\d(\.\d+)?e\+\d+$/i.test(r.employee_no)||/^\d(\.\d+)?e\+\d+$/i.test(r.bpjs_kesehatan)||/^\d(\.\d+)?e\+\d+$/i.test(r.bpjs_ketenagakerjaan))?'<div class="sk-err">NIP/BPJS terbaca dalam format ilmiah (mis. 1.98E+17). Ubah format kolomnya di Excel menjadi Teks, isi ulang, lalu impor kembali.</div>':'')
        +tbl(['NIP','Nama','Email','JK','Departemen','Jabatan','Status','Akhir Kontrak','Pendidikan'],data.slice(0,5).map(r=>[esc(r.employee_no),esc(r.full_name),esc(r.email)||'<span class="sk-sub">kosong</span>',esc(r.gender),esc(r.department_name),esc(r.position),EMPST[r.employment_status],esc(r.contract_end_date),esc([r.education_level,r.education_major].filter(Boolean).join(' '))]));
    }catch(e){$('#iprev',m.el).innerHTML='<div class="sk-err">'+esc(errMsg(e))+'</div>'}}},
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Impor',i:'upload',keep:true,fn:async m=>{
      if(!m._data||!m._data.length){toast('Pilih berkas terlebih dahulu','warn');return false}
      if(S.legacyEmp&&m._data.some(r=>EMP_X.some(k=>r[k]))){toast(LEGACY_MSG,'bad');return false}
      let ins=0,xfail=0;const skipped=[],good=[];
      for(const r of m._data){const v=empValidate(r);if(v)skipped.push({employee_no:r.employee_no,reason:v});else good.push(r)}
      for(const ch of chunk(good,150)){const send=ch.map(({_raw,...r})=>r);const r=await rpc('adm_create_employees',{p_rows:send});ins+=r.inserted;const sk=r.skipped||[];skipped.push(...sk);
        const skn=new Set(sk.map(x=>String(x.employee_no)));xfail+=await empXSave(send.filter(x=>!skn.has(String(x.employee_no))))}
      m.close();Ctx.reset();await render(true);
      modal({title:'Hasil Impor',body:'<div class="sk-ok">'+ins+' pegawai berhasil ditambahkan.</div>'+(xfail?'<div class="sk-warn">'+xfail+' baris: kolom tambahan (pendidikan, BPJS, dll.) gagal disimpan. Pastikan migrasi-karyawan.sql sudah dijalankan, lalu ubah datanya manual.</div>':'')+(skipped.length?'<div class="sk-warn">'+skipped.length+' baris dilewati:</div>'+tbl(['NIP','Alasan'],skipped.map(s=>[esc(s.employee_no||'-'),esc(s.reason)]),{maxh:'260px'}):''),actions:[{t:'Tutup',c:'sec',fn:()=>{}}]})}}]});
};
ACT.empTpl=()=>saveFile('templat-karyawan.csv',new Blob([toCSV(['nip','nama_lengkap','gelar','jenis_kelamin','tanggal_bergabung','tempat_lahir','tanggal_lahir','tanggal_akhir_kontrak','departemen','divisi','jabatan','pendidikan_terakhir','nama_pendidikan','jurusan','bpjs_kesehatan','bpjs_ketenagakerjaan','status_karyawan','no_handphone','email','status_pernikahan','url_pas_photo','tipe','sandi'],[['198001012005011001','Contoh Guru','S.Pd','L','2024-07-15','Bogor','1980-01-01','','Akademik','Kurikulum','Guru Matematika','S1','Universitas Pendidikan Indonesia','Pendidikan Matematika','0001234567890','12345678901','tetap','08123456789','guru@fiwa.sch.id','Menikah','https://lh3.googleusercontent.com/d/1Vv13nci81H8HJEPR-EphYk_r6RAUmw8o','guru',''],['199201152020011002','Contoh Staf','','P','2025-01-06','Jakarta','1992-01-15','2026-01-05','Tata Usaha','Administrasi','Staf Administrasi','D3','Politeknik Negeri Jakarta','Administrasi Bisnis','','','kontrak','','staf@fiwa.sch.id','Belum Menikah','','karyawan',''],])],{type:'text/csv;charset=utf-8'}));
ACT.deptEdit=t=>{
  const d=t.dataset.id?S.ctx.departments.find(x=>x.id===+t.dataset.id):null;
  modal({title:d?'Ubah Departemen':'Tambah Departemen',body:'<form id="df">'+fld('Nama departemen',inp('name',d?d.name:'',{req:1}))+fld('Kepala departemen',sel('head_employee_id',empOptions(),d?d.head_employee_id:'',{blank:'- belum ada -'}))+(d?chk('is_active','Departemen aktif',d.is_active):'')+'</form>',
   actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const f=formData($('#df',m.el));if(!f.name.trim()){toast('Nama wajib diisi','warn');return false}
     const row={name:f.name.trim(),head_employee_id:f.head_employee_id?+f.head_employee_id:null};if(d)row.is_active=!!f.is_active;
     if(d)await DB.update('departments',row,{id:d.id});else await DB.insert('departments',[{...row,is_active:true}]);await audit('department.save','departments',d?d.id:row.name,d,row);toast('Departemen tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.deptDel=async t=>{const d=S.ctx.departments.find(x=>x.id===+t.dataset.id);if(!await confirmBox('Hapus departemen "'+d.name+'"? Pegawai di dalamnya menjadi tanpa departemen.',{danger:true,ok:'Hapus'}))return;await DB.remove('departments',{id:d.id});await audit('department.delete','departments',d.id,d,null);toast('Departemen dihapus','ok');Ctx.reset();await render(true)};
ACT.subAdd=()=>{
  modal({title:'Tambah Pengganti Kepala Departemen',body:'<form id="sf">'+fld('Departemen',sel('department_id',deptOptions(),'',{req:1}))+fld('Pegawai pengganti',sel('substitute_employee_id',empOptions(),'',{req:1}))+'<div class="sk-grid sk-g2">'+fld('Mulai',inp('start_date',TODAY(),{type:'date',req:1}))+fld('Selesai',inp('end_date',D.add(TODAY(),6),{type:'date',req:1}))+'</div></form>',
   actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const f=formData($('#sf',m.el));if(f.end_date<f.start_date){toast('Tanggal selesai tidak boleh sebelum mulai','warn');return false}
     await DB.insert('department_substitutes',[{department_id:+f.department_id,substitute_employee_id:+f.substitute_employee_id,start_date:f.start_date,end_date:f.end_date}]);await audit('substitute.add','department_substitutes',f.department_id,null,f);toast('Pengganti tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.subDel=async t=>{if(!await confirmBox('Hapus penugasan pengganti ini?',{danger:true,ok:'Hapus'}))return;await DB.remove('department_substitutes',{id:+t.dataset.id});toast('Dihapus','ok');Ctx.reset();await render(true)};
ACT.fpEdit=t=>{
  const emp=+t.dataset.e,cur=(S.ctx.fingerprints||[]).find(f=>f.employee_id===emp);
  modal({title:'ID Fingerprint - '+nameOf(emp),body:'<form id="ff">'+fld('ID pegawai di mesin',inp('fp',cur?cur.fingerprint_id:(S.ctx.employees.find(e=>e.id===emp)||{}).employee_no||'',{req:1}),'Default sama dengan NIP. Ubah hanya bila ID di mesin berbeda.')+'</form>',
   actions:[...(cur?[{t:'Hapus',c:'bad',i:'trash',fn:async()=>{await DB.remove('employee_fingerprints',{id:cur.id});toast('Dihapus','ok');Ctx.reset();render(true)}}]:[]),{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{
     const fp=formData($('#ff',m.el)).fp.trim();if(!fp){toast('ID wajib diisi','warn');return false}
     if((S.ctx.fingerprints||[]).some(f=>f.fingerprint_id===fp&&f.employee_id!==emp)){toast('ID tersebut sudah dipakai pegawai lain','bad');return false}
     if(cur)await DB.update('employee_fingerprints',{fingerprint_id:fp},{id:cur.id});else await DB.insert('employee_fingerprints',[{employee_id:emp,fingerprint_id:fp}]);
     let n=0;try{const r=await DB.update('raw_logs',{employee_id:emp,match_status:'matched'},{fingerprint_id:fp,match_status:'unmatched'});n=r.length}catch(e){}
     toast('ID tersimpan'+(n?'. '+n+' log lama dicocokkan':''),'ok');Ctx.reset();render(true)}}]});
};

/* =====================================================================
   17. ADMIN: AKUN PENGGUNA
   ===================================================================== */
S.ak={tab:'pegawai',q:'',f:'all',lim:60};
const AKADM={admin_sistem:['shield','','b','Super Admin','Pengaturan sistem, akun, role, dan konfigurasi aplikasi.'],manajer_hrd:['check','v','v','HRD Manager','Approval akhir, monitoring seluruh karyawan, validasi kebijakan, dan laporan strategis HR.'],staf_hrd:['file','c','c','HRD Staff / Admin HR','Mengelola data karyawan, memeriksa absensi, memvalidasi pengajuan, dan memproses administrasi.']};
function akCard(e,now){
  const lk=!!(e.locked_until&&Date.parse(e.locked_until)>now),st=!e.is_active?'off':lk?'lock':'ok',roles=e.roles||[];
  const flags=(lk?bd('Terkunci','r'):'')+(e.must_change_password?bd('Wajib ganti sandi','a'):'')+(e.failed_login_count>0&&!lk?`<span class="ux-pill warn">${e.failed_login_count}x gagal login</span>`:'');
  const lg=e.last_login_at?`<span>${ic('clock','sm')} Login terakhir <b>${esc(uxAgo(e.last_login_at))}</b> &middot; ${esc(fmtDT(e.last_login_at))}</span>`:`<span>${ic('clock','sm')} Belum pernah login</span>`;
  return`<article class="ux-card ak-item st-${st}">`
    +`<div class="ux-h">${uxAv(e)}<div class="ux-id"><b title="${esc(e.full_name)}">${esc(e.full_name)}</b><small>${esc(deptName(e.department_id))}${e.position?' &middot; '+esc(e.position):''}</small></div></div>`
    +`<div class="ak-mail${e.email?'':' none'}">${ic('user','sm')}<span>${e.email?esc(e.email):'Email belum diisi'}</span></div>`
    +`<div class="ak-roles">${roles.map(r=>bd(ROLE[r]||r,r==='pegawai'?'s':'b')).join('')}${flags}</div>`
    +`<div class="ak-meta">${lg}</div>`
    +`<div class="ak-f"><span class="ak-sl">${e.is_active?'Akun aktif':'Akun nonaktif'}</span><button type="button" class="ak-sw${e.is_active?' on':''}" role="switch" aria-checked="${e.is_active?'true':'false'}" aria-label="${e.is_active?'Nonaktifkan':'Aktifkan'} akun ${esc(e.full_name)}" title="${e.is_active?'Nonaktifkan akun':'Aktifkan akun'}" data-act="akToggle" data-id="${e.id}"><i></i></button><span class="sp"></span>`
    +(lk||e.failed_login_count?`<button class="sk-b warn sm" data-act="akUnlock" data-id="${e.id}">${ic('lock','sm')}Buka kunci</button>`:'')
    +`<button class="sk-b sec sm" data-act="akReset" data-id="${e.id}">${ic('key','sm')}Reset sandi</button></div></article>`;
}
function akRow(e,now){
  const lk=!!(e.locked_until&&Date.parse(e.locked_until)>now),roles=e.roles||[];
  const flags=(lk?bd('Terkunci','r'):'')+(e.must_change_password?bd('Wajib ganti sandi','a'):'')+(e.failed_login_count>0&&!lk?bd(e.failed_login_count+'x gagal login','a'):'');
  return[
    `<div class="ap-emp">${uxAv(e)}<div><b>${esc(e.full_name)}</b><div class="sk-sub">${esc(deptName(e.department_id))}${e.position?' &middot; '+esc(e.position):''}</div></div></div>`,
    e.email?esc(e.email):'<span class="sk-sub">Belum diisi</span>',
    `<div class="ak-rl">${roles.map(r=>bd(ROLE[r]||r,r==='pegawai'?'s':'b')).join('')}</div>`,
    (flags||'<span class="sk-sub">Normal</span>')+(e.is_active?'':' '+bd('Nonaktif','s')),
    e.last_login_at?`${esc(uxAgo(e.last_login_at))}<div class="sk-sub">${esc(fmtDT(e.last_login_at))}</div>`:'<span class="sk-sub">Belum pernah</span>',
    `<button type="button" class="ak-sw${e.is_active?' on':''}" role="switch" aria-checked="${e.is_active?'true':'false'}" aria-label="${e.is_active?'Nonaktifkan':'Aktifkan'} akun ${esc(e.full_name)}" title="${e.is_active?'Nonaktifkan akun':'Aktifkan akun'}" data-act="akToggle" data-id="${e.id}"><i></i></button>`,
    `<div class="ap-btns">${lk||e.failed_login_count?`<button class="sk-b warn sm" data-act="akUnlock" data-id="${e.id}">Buka kunci</button>`:''}<button class="sk-b sec sm" data-act="akReset" data-id="${e.id}">Reset sandi</button></div>`
  ];
}
reg('akun',{
  range:()=>[TODAY(),TODAY()],
  async render(){
    const c=S.ctx,a=S.ak,tabs=stabs([['pegawai','Akun Pegawai'],['admin','Akun Admin / HRD']],a.tab,'akTab');
    if(a.tab==='admin'){
      const admins=await DB.select('admin_profiles');S._admins=admins;
      const ord={admin_sistem:0,manajer_hrd:1,staf_hrd:2},sorted=admins.slice().sort((x,y)=>(ord[x.role]??9)-(ord[y.role]??9)||String(x.full_name).localeCompare(String(y.full_name)));
      const act=admins.filter(x=>x.is_active).length;
      const hero=`<div class="rk-top ux-hero"><div class="rk-tl"><span class="rk-eye">${ic('key','sm')}Akun Pengguna</span><div class="ux-ht">Akun Admin / HRD</div><div class="ux-hs">Kelola siapa saja yang dapat mengakses panel admin beserta tingkat aksesnya.</div><div class="rk-meta"><span class="rk-mc">${ic('shield','sm')}${admins.length} akun admin</span><span class="rk-mc">${ic('check','sm')}${act} aktif</span></div></div>${isSysAdmin()?`<div class="ux-hr"><button class="sk-b sm" data-act="admEdit">${ic('plus','sm')}Tambah Admin</button></div>`:''}</div>`;
      const leg='<div class="ak-leg">'+Object.keys(AKADM).map(k=>{const m=AKADM[k];return`<div class="ak-lg ${m[1]}"><span class="li">${ic(m[0])}</span><div><b>${m[3]}</b><small>${m[4]}</small></div></div>`}).join('')+'</div>';
      const cards=sorted.map(x=>{const m=AKADM[x.role]||['shield','','s',ROLE[x.role]||x.role,''],me=S.admin&&x.id===S.admin.id;
        return`<article class="ux-card ak-item st-${x.is_active?'ok':'off'}"><div class="ux-h">${uxAv({id:x.id},x.full_name)}<div class="ux-id"><b title="${esc(x.full_name)}">${esc(x.full_name)}</b><small>${esc(ROLE[x.role]||x.role)}</small></div>${me?bd('Anda','g'):''}</div>`
          +(x.email?`<div class="ak-mail">${ic('user','sm')}<span>${esc(x.email)}</span></div>`:'')
          +`<div class="ak-roles">${bd(ROLE[x.role]||x.role,m[2])}${x.is_active?bd('Aktif','g'):bd('Nonaktif','s')}</div>`
          +`<div class="ak-meta"><span>${ic('users','sm')} ${x.employee_id?'Tertaut ke <b>'+esc(nameOf(x.employee_id))+'</b>':'Belum tertaut ke data pegawai'}</span></div>`
          +(isSysAdmin()?`<div class="ak-f"><span class="sp"></span><button class="sk-b sec sm" data-act="admEdit" data-id="${x.id}">${ic('edit','sm')}Ubah</button></div>`:'')+'</article>'}).join('');
      return hero+tabs+leg+(cards?'<div class="ux-grid">'+cards+'</div>':`<div class="ux-empty"><div class="ei">${ic('shield')}</div><b>Belum ada admin</b><p>Tambahkan akun admin pertama untuk mulai mengelola sistem.</p></div>`)
        +'<div class="sk-info">'+(IS_PREVIEW?'Pada mode Preview, email dan kata sandi admin diisi langsung saat menambah admin.':'Untuk menambah admin: buat pengguna di Supabase (Authentication &gt; Users), salin User UID-nya, lalu isi di tombol Tambah Admin.')+'</div>';
    }
    const now=Date.now(),isLk=e=>!!(e.locked_until&&Date.parse(e.locked_until)>now);
    const F={all:()=>true,active:e=>e.is_active,off:e=>!e.is_active,lock:isLk,pw:e=>!!e.must_change_password,never:e=>e.is_active&&!e.last_login_at};
    const FL=[['all','Semua'],['active','Aktif'],['off','Nonaktif'],['lock','Terkunci'],['pw','Wajib ganti sandi'],['never','Belum pernah login']];
    if(!F[a.f])a.f='all';
    const q=(a.q||'').trim().toLowerCase();
    efFix('ak',c.employees);const list=c.employees.filter(F[a.f]).filter(e=>efPass('ak',e)).filter(e=>!q||String(e.full_name||'').toLowerCase().includes(q)||String(e.email||'').toLowerCase().includes(q)||String(e.employee_no||'').toLowerCase().includes(q)||deptName(e.department_id).toLowerCase().includes(q));
    const shown=list,nAct=c.employees.filter(e=>e.is_active).length,nLock=c.employees.filter(isLk).length;
    const hero=`<div class="rk-top ux-hero"><div class="rk-tl"><span class="rk-eye">${ic('key','sm')}Akun Pengguna</span><div class="ux-ht">Akun Pegawai</div><div class="ux-hs">Atur akses login pegawai: reset kata sandi, buka kunci akun, dan aktifkan atau nonaktifkan akun.</div><div class="rk-meta"><span class="rk-mc">${ic('users','sm')}${c.employees.length} akun</span><span class="rk-mc">${ic('check','sm')}${nAct} aktif</span><span class="rk-mc">${ic('lock','sm')}${nLock} terkunci</span></div></div></div>`;
    const chips='<div class="ux-chips">'+FL.map(([k,l])=>`<button type="button" class="ux-chip${k===a.f?' on':''}" data-act="akFilter" data-k="${k}">${esc(l)}<em>${c.employees.filter(F[k]).length}</em></button>`).join('')+'</div>';
    const tool=`<div class="ux-tool"><label class="ux-sr">${ic('search','sm')}<input class="sk-in" type="search" placeholder="Cari nama, email, NIP, atau departemen" aria-label="Cari akun" value="${esc(a.q)}" data-inp="akQ"></label>${efBar('ak',c.employees,{inl:1})}<span class="ux-cnt">${list.length} akun</span></div>`;
    const body=shown.length?card('',tbl(['Pegawai','Email',{t:'Peran',c:''},'Status',{t:'Login terakhir',c:'nw'},{t:'Aktif',c:'ctr'},{t:'Aksi',c:'ap-act'}],shown.map(e=>akRow(e,now)),{cls:'ap-tbl ak-tbl',page:30,ef:'ak'}))
      :`<div class="ux-empty"><div class="ei">${ic('search')}</div><b>Akun tidak ditemukan</b><p>Ubah kata kunci atau filter yang dipilih.</p></div>`;
    return hero+tabs+chips+tool+body
      +'<div class="sk-info">Pegawai masuk memakai email dan kata sandi; email diubah di menu Pegawai &amp; Departemen. Reset sandi mengatur kata sandi sementara dan mewajibkan pegawai menggantinya. Peran (Supervisor dan Manager / Kepala Departemen) diubah di menu Pegawai &amp; Departemen.</div>';
  }
});
ACT.akTab=async t=>{S.ak.tab=t.dataset.k;await render()};
INP.akQ=debounce(async t=>{S.ak.q=t.value;S.ak.lim=60;const p=t.selectionStart;await render();const i=$('[data-inp=akQ]');if(i){i.focus();try{i.setSelectionRange(p,p)}catch(e){}}},350);
ACT.akFilter=async t=>{S.ak.f=t.dataset.k;S.ak.lim=60;await render()};
ACT.akMore=async()=>{S.ak.lim+=60;await render()};
ACT.akReset=async t=>{
  const e=S.ctx.employees.find(x=>x.id===+t.dataset.id);
  const pw=await promptBox('Reset kata sandi - '+e.full_name,'Kata sandi sementara baru',{multi:false,value:'Sk'+Math.random().toString(36).slice(2,8),hint:'Minimal 6 karakter. Pegawai wajib menggantinya saat login berikutnya dan sesi lamanya diakhiri.',ok:'Reset'});
  if(pw===null)return;if(pw.length<6){toast('Minimal 6 karakter','warn');return}
  await rpc('adm_set_employee_password',{p_id:e.id,p_password:pw});toast('Kata sandi direset','ok');Ctx.reset();await render(true);
};
ACT.akUnlock=async t=>{await DB.update('employees',{failed_login_count:0,locked_until:null},{id:+t.dataset.id});toast('Kunci dibuka','ok');Ctx.reset();await render(true)};
ACT.akToggle=async t=>{const e=S.ctx.employees.find(x=>x.id===+t.dataset.id);if(e.is_active&&!await confirmBox('Nonaktifkan akun '+e.full_name+'? Pegawai tidak dapat login dan tidak dihitung dalam rekap.',{ok:'Nonaktifkan',danger:true}))return;
  await DB.update('employees',{is_active:!e.is_active,updated_at:new Date().toISOString()},{id:e.id});await audit('employee.toggle','employees',e.id,{is_active:e.is_active},{is_active:!e.is_active});toast('Status akun diperbarui','ok');Ctx.reset();await render(true)};
ACT.admEdit=t=>{
  const x=t.dataset.id?S._admins.find(a=>a.id===+t.dataset.id):null;
  modal({title:x?'Ubah Admin':'Tambah Admin',body:'<form id="af">'+fld('Nama lengkap',inp('full_name',x?x.full_name:'',{req:1}))
    +(x?'':(IS_PREVIEW?fld('Email',inp('email','',{type:'email',req:1}))+fld('Kata sandi',inp('pw','',{req:1})):fld('User UID Supabase',inp('auth_user_id','',{req:1,ph:'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'}),'Authentication > Users > klik pengguna > salin User UID')))
    +fld('Peran',sel('role',[['admin_sistem','Super Admin'],['manajer_hrd','HRD Manager'],['staf_hrd','HRD Staff / Admin HR']],x?x.role:'staf_hrd'))
    +fld('Tertaut ke pegawai (opsional)',sel('employee_id',empOptions(),x?x.employee_id:'',{blank:'- tidak ada -'}),'Dipakai agar pengajuan milik admin tidak dapat diputuskan sendiri')+(x?chk('is_active','Aktif',x.is_active):'')+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{
      const f=formData($('#af',m.el));if(!f.full_name.trim()){toast('Nama wajib diisi','warn');return false}
      const row={full_name:f.full_name.trim(),role:f.role,employee_id:f.employee_id?+f.employee_id:null};
      if(x){if(x.id===(S.admin.id)&&(!f.is_active||f.role!=='admin_sistem')){toast('Anda tidak dapat menurunkan atau menonaktifkan akun sendiri','bad');return false}row.is_active=!!f.is_active;await DB.update('admin_profiles',row,{id:x.id})}
      else{if(IS_PREVIEW)Object.assign(row,{auth_user_id:'demo-'+uid(),email:f.email.trim().toLowerCase(),pw:f.pw,is_active:true});else{if(!/^[0-9a-f-]{36}$/i.test(f.auth_user_id.trim())){toast('User UID tidak valid','warn');return false}Object.assign(row,{auth_user_id:f.auth_user_id.trim(),is_active:true})}await DB.insert('admin_profiles',[row])}
      await audit('admin.save','admin_profiles',x?x.id:row.full_name,x,row);toast('Akun admin tersimpan','ok');Ctx.reset();render(true)}}]});
};


/* =====================================================================
   18. ADMIN: JADWAL NORMAL
   ===================================================================== */
const curAssign=(emp,ds)=>((Engine.ix(S.ctx).assign[emp]||[]).find(a=>String(a.effective_from).slice(0,10)<=ds&&(!a.effective_to||String(a.effective_to).slice(0,10)>=ds)))||null;
const hasShifts=emp=>(S.ctx.shift_assignments||[]).some(a=>a.employee_id===emp);
async function assignSchedule(ids,wsId,from){
  const ins=[];
  for(const id of ids){
    for(const a of (S.ctx.schedule_assignments||[]).filter(x=>x.employee_id===id)){
      const af=String(a.effective_from).slice(0,10),at=a.effective_to?String(a.effective_to).slice(0,10):null;
      if(af>=from)await DB.remove('schedule_assignments',{id:a.id});
      else if(!at||at>=from)await DB.update('schedule_assignments',{effective_to:D.add(from,-1)},{id:a.id});
    }
    ins.push({employee_id:id,work_schedule_id:wsId,effective_from:from,effective_to:null});
  }
  if(ins.length)await DB.insert('schedule_assignments',ins);
}
const dayLabels=['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'];
function scheduleSummary(ws){
  return(S.ctx.work_schedule_days||[]).filter(d=>d.work_schedule_id===ws.id).sort((a,b)=>a.weekday-b.weekday).map(d=>'<div class="sk-sub"><b>'+DAYS3[d.weekday-1]+'</b> '+(d.is_workday&&d.start_time?D.hm(d.start_time)+' - '+D.hm(d.end_time)+(d.break_minutes?' (istirahat '+d.break_minutes+' mnt)':''):'<i>libur</i>')+'</div>').join('');
}
/* =====================================================================
   EKSPOR / IMPOR JADWAL (Jadwal Normal, Penugasan, Master Shift, Pola, Penyusun)
   Impor selalu menampilkan pratinjau dulu; baris bermasalah dilewati dan dilaporkan.
   ===================================================================== */
const jxPad=n=>String(n).padStart(2,'0');
function jxDate(v){v=String(v==null?'':v).trim();if(!v)return'';let x=v.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);if(x)return x[1]+'-'+jxPad(x[2])+'-'+jxPad(x[3]);
  x=v.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);if(x)return x[3]+'-'+jxPad(x[2])+'-'+jxPad(x[1]);
  x=v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);if(x)return(+x[3]>40?'19':'20')+x[3]+'-'+jxPad(x[1])+'-'+jxPad(x[2]);
  if(/^\d{5}(\.\d+)?$/.test(v)){const n=Math.floor(+v);if(n>20000&&n<80000)return new Date(Date.UTC(1899,11,30)+n*864e5).toISOString().slice(0,10)}return''}
function jxTime(v){v=String(v==null?'':v).trim();const x=v.match(/^(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(am|pm)?$/i);if(!x)return'';let h=+x[1];if(x[3]){if(/pm/i.test(x[3])&&h<12)h+=12;if(/am/i.test(x[3])&&h===12)h=0}return h<24&&+x[2]<60?jxPad(h)+':'+x[2]:''}
function jxDayParse(v){v=String(v==null?'':v).trim();if(!v||/^(libur|off|-|x)$/i.test(v))return{w:false};
  const m=v.match(/^(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(?:-|–|—|s\/d|sd|sampai)\s*(\d{1,2})[:.](\d{2})(?::\d{2})?(?:\s*[\/|(]\s*(?:istirahat\s*)?(\d+)\s*\)?)?$/i);if(!m)return null;
  const s=jxTime(m[1]+':'+m[2]),e=jxTime(m[3]+':'+m[4]);return s&&e?{w:true,s,e,b:+m[5]||0}:null}
const jxDay=d=>d&&d.is_workday&&d.start_time?D.hm(d.start_time)+'-'+D.hm(d.end_time)+(d.break_minutes?'/'+d.break_minutes:''):'Libur';
const jxH=rows=>{const h=rows[0].map(normKey);return(...p)=>h.findIndex(x=>p.some(q=>x===q||x.startsWith(q)))};
const jxStat=v=>{v=String(v==null?'':v).trim();return v===''?null:!/^(nonaktif|tidak\s*aktif|tidak|0|no|false|n)$/i.test(v)};
const jxNum=(v,d)=>{const s=String(v==null?'':v).trim();if(s==='')return d;const n=+s.replace(',','.');return isNaN(n)||n<0?NaN:Math.round(n)};
const jxErrs=(a,t)=>a.length?'<div class="sk-warn"><b>'+a.length+' '+t+'</b><br>'+a.slice(0,8).map(esc).join('<br>')+(a.length>8?'<br>... dan '+(a.length-8)+' lainnya':'')+'</div>':'';
const jxBad=t=>({n:0,html:'<div class="sk-err">'+esc(t)+'</div>'});
const jxEmp=(nip,name)=>{const k=nipKey(nip).toLowerCase();return(k&&S.ctx.employees.find(e=>nipKey(e.employee_no).toLowerCase()===k))||(name?S.ctx.employees.find(e=>e.is_active&&String(e.full_name).trim().toLowerCase()===String(name).trim().toLowerCase()):null)||null};
function jxImport(o){
  modal({title:o.title,size:'wide',body:'<div class="sk-info">'+o.info+'</div><label class="sk-f"><span>Berkas Excel / CSV</span><input class="sk-in" type="file" id="jxf" accept=".xlsx,.xls,.csv"></label><div id="jxp"></div>',
    onMount:m=>{$('#jxf',m.el).onchange=async ev=>{const f=ev.target.files[0];m._r=null;if(!f)return;const box=$('#jxp',m.el);
      try{const rows=await readTable(f);if(rows.length<2){box.innerHTML='<div class="sk-err">Berkas kosong atau hanya berisi baris judul</div>';return}const r=o.parse(rows);m._r=r;box.innerHTML=r.html}
      catch(e){box.innerHTML='<div class="sk-err">'+esc(errMsg(e))+'</div>'}}},
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Impor',i:'upload',keep:true,fn:async m=>{
      if(!m._r||!m._r.n){toast('Pilih berkas yang berisi data valid terlebih dahulu','warn');return false}
      const msg=await m._r.apply();toast(msg,'ok');m.close()}}]});
}
/* ---- Jadwal Normal: daftar jadwal kerja ---- */
ACT.jnExport=()=>{const c=S.ctx;exportRows('jadwal-normal','Jadwal Normal',['Jadwal','Toleransi Terlambat (mnt)','Absen Masuk Dibuka (mnt)','Absen Pulang Diterima (mnt)','Status',...dayLabels],
  c.work_schedules.map(w=>{const by={};(c.work_schedule_days||[]).filter(d=>d.work_schedule_id===w.id).forEach(d=>by[d.weekday]=d);return[w.name,w.tolerance_minutes,w.early_clock_in_minutes,w.late_clock_out_minutes,w.is_active?'Aktif':'Nonaktif',...[1,2,3,4,5,6,7].map(n=>jxDay(by[n]))]}))};
ACT.jnImport=()=>jxImport({title:'Impor Jadwal Kerja',info:'Gunakan berkas hasil <b>Ekspor</b> sebagai acuan. Satu baris = satu jadwal, dicocokkan lewat <b>nama</b> (nama baru = jadwal baru, nama sama = diperbarui). Kolom hari ditulis <b>07:00-15:00</b> atau <b>07:00-15:00/30</b> (angka setelah garis miring = istirahat dalam menit), atau <b>Libur</b>. Kolom yang tidak ada atau dikosongkan (toleransi, batas absen, status) memakai nilai yang sudah tersimpan.',
  parse:rows=>{
    const hx=jxH(rows),c=S.ctx,ci={nm:hx('jadwal','namajadwal','nama'),tol:hx('toleransi'),er:hx('absenmasuk','bukamasuk'),lt:hx('absenpulang','tutuppulang'),st:hx('status')},dc=['senin','selasa','rabu','kamis','jumat','sabtu','minggu'].map(k=>hx(k));
    if(ci.nm<0)return jxBad('Kolom "Jadwal" tidak ditemukan pada baris judul.');if(dc.every(i=>i<0))return jxBad('Kolom hari (Senin sampai Minggu) tidak ditemukan pada baris judul.');
    const g=(r,i)=>i>=0?r[i]:'',errs=[],items=[],seen=new Set();
    rows.slice(1).forEach((r,k)=>{const ln='Baris '+(k+2)+': ',name=normSp(g(r,ci.nm));if(!name)return;
      if(seen.has(name.toLowerCase())){errs.push(ln+'nama jadwal ganda ('+name+')');return}seen.add(name.toLowerCase());
      const ex=c.work_schedules.find(w=>w.name.toLowerCase()===name.toLowerCase()),tol=jxNum(g(r,ci.tol),ex?ex.tolerance_minutes:10),er=jxNum(g(r,ci.er),ex?ex.early_clock_in_minutes:120),lt=jxNum(g(r,ci.lt),ex?ex.late_clock_out_minutes:240),st=jxStat(g(r,ci.st));
      if([tol,er,lt].some(isNaN)){errs.push(ln+name+': angka toleransi / batas absen tidak valid');return}
      const days=[];let bad='';
      for(let n=1;n<=7;n++){const i=dc[n-1];
        if(i<0){const d=ex&&(c.work_schedule_days||[]).find(x=>x.work_schedule_id===ex.id&&x.weekday===n);days.push(d?{weekday:n,is_workday:!!d.is_workday,start_time:d.start_time?D.hm(d.start_time):null,end_time:d.end_time?D.hm(d.end_time):null,break_minutes:d.break_minutes||0}:{weekday:n,is_workday:false,start_time:null,end_time:null,break_minutes:0});continue}
        const p=jxDayParse(r[i]);if(!p){bad=dayLabels[n-1]+' "'+(r[i]==null?'':r[i])+'"';break}
        days.push(p.w?{weekday:n,is_workday:true,start_time:p.s,end_time:p.e,break_minutes:p.b}:{weekday:n,is_workday:false,start_time:null,end_time:null,break_minutes:0})}
      if(bad){errs.push(ln+name+': jam tidak terbaca di '+bad+' (contoh 07:00-15:00/30)');return}
      items.push({ex,days,row:{name,tolerance_minutes:tol,early_clock_in_minutes:er,late_clock_out_minutes:lt,is_active:st==null?(ex?ex.is_active:true):st}})});
    const nNew=items.filter(x=>!x.ex).length;
    return{n:items.length,html:(items.length?'<div class="sk-ok">'+items.length+' jadwal siap diimpor: '+nNew+' baru, '+(items.length-nNew)+' diperbarui.</div>':'')+jxErrs(errs,'baris dilewati')
      +(items.length?tbl(['Jadwal','Aksi','Hari kerja'],items.slice(0,8).map(x=>['<b>'+esc(x.row.name)+'</b>',x.ex?bd('Perbarui','b'):bd('Baru','g'),esc(x.days.map(d=>d.is_workday?dayLabels[d.weekday-1].slice(0,3)+' '+d.start_time+'-'+d.end_time:'').filter(Boolean).join(', ')||'Tidak ada hari kerja')]),{page:8}):''),
      apply:async()=>{for(const it of items){let id=it.ex?it.ex.id:null;if(id)await DB.update('work_schedules',it.row,{id});else{const r=await DB.insert('work_schedules',[it.row]);id=r[0].id}
        await DB.upsert('work_schedule_days',it.days.map(x=>({...x,work_schedule_id:id})),'work_schedule_id,weekday')}
        await audit('schedule.import','work_schedules','import',null,{jumlah:items.length});Ctx.reset();await render(true);return items.length+' jadwal kerja diimpor'}}}});
/* ---- Jadwal Normal: penugasan ke pegawai ---- */
ACT.jnaExport=()=>{const c=S.ctx,t=TODAY();efFix('jn',c.employees.filter(e=>e.is_active));
  exportRows('penugasan-jadwal-normal','Penugasan',['NIP','Nama','Departemen','Divisi','Jadwal Normal','Berlaku Mulai'],c.employees.filter(e=>e.is_active&&efPass('jn',e)).sort((a,b)=>a.full_name.localeCompare(b.full_name,'id')).map(e=>{const a=curAssign(e.id,t),ws=a?c.work_schedules.find(w=>w.id===a.work_schedule_id):null;return[String(e.employee_no),e.full_name,deptName(e.department_id)==='-'?'':deptName(e.department_id),e.division||'',ws?ws.name:'',a?String(a.effective_from).slice(0,10):'']}))};
ACT.jnaImport=()=>jxImport({title:'Impor Penugasan Jadwal Normal',info:'Kolom <b>NIP</b> (atau Nama) untuk mencocokkan pegawai, <b>Jadwal Normal</b> (nama persis seperti di Daftar Jadwal Kerja), dan <b>Berlaku Mulai</b> (YYYY-MM-DD atau DD/MM/YYYY; kosong = hari ini). Baris dengan jadwal kosong atau yang sama dengan jadwal pegawai saat ini dilewati. Riwayat sebelum tanggal berlaku tidak berubah. Format kolom NIP di Excel sebaiknya Teks.',
  parse:rows=>{
    const hx=jxH(rows),ci={no:hx('nip','nomorinduk','noinduk'),nm:hx('nama','pegawai','namalengkap'),ws:hx('jadwalnormal','jadwal'),fr:hx('berlakumulai','berlaku','mulai','efektif')};
    if(ci.no<0&&ci.nm<0)return jxBad('Kolom NIP atau Nama tidak ditemukan pada baris judul.');if(ci.ws<0)return jxBad('Kolom "Jadwal Normal" tidak ditemukan pada baris judul.');
    const g=(r,i)=>i>=0?r[i]:'',c=S.ctx,t=TODAY(),errs=[],items=[];let same=0,empty=0;const seen=new Set();
    rows.slice(1).forEach((r,k)=>{const ln='Baris '+(k+2)+': ',nip=String(g(r,ci.no)||'').trim(),nm=String(g(r,ci.nm)||'').trim(),wn=normSp(g(r,ci.ws));if(!nip&&!nm)return;
      const e=jxEmp(nip,nm);if(!e){errs.push(ln+'pegawai tidak ditemukan ('+(nip||nm)+')');return}
      if(!wn){empty++;return}const w=c.work_schedules.find(x=>x.name.toLowerCase()===wn.toLowerCase());if(!w){errs.push(ln+e.full_name+': jadwal "'+wn+'" tidak ada');return}if(!w.is_active){errs.push(ln+e.full_name+': jadwal "'+w.name+'" nonaktif');return}
      const raw=String(g(r,ci.fr)||'').trim(),from=raw?jxDate(raw):t;if(!from){errs.push(ln+e.full_name+': tanggal berlaku tidak terbaca ('+raw+')');return}
      if(seen.has(e.id)){errs.push(ln+e.full_name+': pegawai muncul lebih dari sekali');return}seen.add(e.id);
      const cur=curAssign(e.id,from);if(cur&&cur.work_schedule_id===w.id){same++;return}
      items.push({e,w,from})});
    return{n:items.length,html:(items.length?'<div class="sk-ok">'+items.length+' penugasan akan diubah.</div>':'<div class="sk-warn">Tidak ada perubahan yang perlu diimpor.</div>')
      +(same||empty?'<div class="sk-sub" style="margin-bottom:8px">'+same+' baris sudah sama dengan jadwal saat ini, '+empty+' baris tanpa jadwal (dilewati).</div>':'')+jxErrs(errs,'baris dilewati')
      +(items.length?tbl(['Pegawai','Jadwal Normal','Berlaku mulai'],items.slice(0,8).map(x=>[esc(x.e.full_name),bd(x.w.name,'b'),esc(fmtD(x.from))]),{page:8}):''),
      apply:async()=>{const grp=new Map();items.forEach(x=>{const k=x.w.id+'|'+x.from;if(!grp.has(k))grp.set(k,{w:x.w,from:x.from,ids:[]});grp.get(k).ids.push(x.e.id)});
        for(const gp of grp.values()){await assignSchedule(gp.ids,gp.w.id,gp.from);if(gp.from<=TODAY())await reprocess(gp.from,TODAY(),gp.ids)}
        await audit('schedule.import_assign','schedule_assignments','import',null,{jumlah:items.length});Ctx.reset();await render(true);return'Penugasan jadwal diubah untuk '+items.length+' pegawai'}}}});
/* ---- Jadwal Shift: master shift ---- */
ACT.shxExport=()=>exportRows('master-shift','Master Shift',['Nama Shift','Jam Masuk','Jam Pulang','Istirahat (mnt)','Toleransi Terlambat (mnt)','Absen Masuk Dibuka (mnt)','Absen Pulang Diterima (mnt)','Warna','Status'],
  (S.ctx.shifts||[]).map(s=>[s.name,D.hm(s.start_time),D.hm(s.end_time),s.break_minutes,s.tolerance_minutes,s.early_clock_in_minutes,s.late_clock_out_minutes,s.color||'#2563eb',s.is_active?'Aktif':'Nonaktif']));
ACT.shxImport=()=>jxImport({title:'Impor Master Shift',info:'Gunakan berkas hasil <b>Ekspor</b> sebagai acuan. Shift dicocokkan lewat <b>nama</b> (nama baru = shift baru). Jam ditulis <b>HH:MM</b>; bila jam pulang lebih kecil atau sama dengan jam masuk, shift dianggap melewati tengah malam. Warna berupa kode heks (mis. #2563eb). Kolom angka, warna, dan status yang kosong memakai nilai yang sudah tersimpan.',
  parse:rows=>{
    const hx=jxH(rows),ci={nm:hx('namashift','shift','nama'),s:hx('jammasuk','masuk'),e:hx('jampulang','pulang'),br:hx('istirahat'),tol:hx('toleransi'),er:hx('absenmasuk','bukamasuk'),lt:hx('absenpulang','tutuppulang'),co:hx('warna'),st:hx('status')};
    if(ci.nm<0||ci.s<0||ci.e<0)return jxBad('Kolom "Nama Shift", "Jam Masuk", dan "Jam Pulang" wajib ada pada baris judul.');
    const g=(r,i)=>i>=0?r[i]:'',errs=[],items=[],seen=new Set();
    rows.slice(1).forEach((r,k)=>{const ln='Baris '+(k+2)+': ',name=normSp(g(r,ci.nm));if(!name)return;
      if(seen.has(name.toLowerCase())){errs.push(ln+'nama shift ganda ('+name+')');return}seen.add(name.toLowerCase());
      const ex=(S.ctx.shifts||[]).find(s=>s.name.toLowerCase()===name.toLowerCase()),st=jxTime(g(r,ci.s)),en=jxTime(g(r,ci.e));
      if(!st||!en){errs.push(ln+name+': jam masuk / pulang tidak terbaca (format HH:MM)');return}
      const br=jxNum(g(r,ci.br),ex?ex.break_minutes:30),tol=jxNum(g(r,ci.tol),ex?ex.tolerance_minutes:10),er=jxNum(g(r,ci.er),ex?ex.early_clock_in_minutes:120),lt=jxNum(g(r,ci.lt),ex?ex.late_clock_out_minutes:240);
      if([br,tol,er,lt].some(isNaN)){errs.push(ln+name+': angka istirahat / toleransi / batas absen tidak valid');return}
      let co=String(g(r,ci.co)||'').trim();if(co&&!/^#[0-9a-f]{6}$/i.test(co)){errs.push(ln+name+': warna harus berupa kode heks seperti #2563eb');return}
      if(!co)co=ex?ex.color||'#2563eb':'#2563eb';const sa=jxStat(g(r,ci.st));
      items.push({ex,row:{name,start_time:st,end_time:en,crosses_midnight:en<=st,break_minutes:br,tolerance_minutes:tol,early_clock_in_minutes:er,late_clock_out_minutes:lt,color:co,is_active:sa==null?(ex?ex.is_active:true):sa}})});
    const nNew=items.filter(x=>!x.ex).length;
    return{n:items.length,html:(items.length?'<div class="sk-ok">'+items.length+' shift siap diimpor: '+nNew+' baru, '+(items.length-nNew)+' diperbarui.</div>':'')+jxErrs(errs,'baris dilewati')
      +(items.length?tbl(['Shift','Aksi','Jam','Status'],items.slice(0,8).map(x=>['<span style="display:inline-block;width:12px;height:12px;border-radius:4px;background:'+esc(x.row.color)+';margin-right:6px"></span><b>'+esc(x.row.name)+'</b>',x.ex?bd('Perbarui','b'):bd('Baru','g'),esc(x.row.start_time+' - '+x.row.end_time+(x.row.crosses_midnight?' (lewat tengah malam)':'')),x.row.is_active?bd('Aktif','g'):bd('Nonaktif','s')]),{page:8}):''),
      apply:async()=>{for(const it of items){if(it.ex)await DB.update('shifts',it.row,{id:it.ex.id});else await DB.insert('shifts',[it.row])}
        await audit('shift.import','shifts','import',null,{jumlah:items.length});Ctx.reset();await render(true);return items.length+' shift diimpor'}}}});
/* ---- Jadwal Shift: pola rotasi ---- */
ACT.plxExport=()=>exportRows('pola-rotasi','Pola Rotasi',['Nama Pola','Siklus'],(S._patterns||[]).map(p=>[p.name,(p.cycle||[]).map(x=>x==null?'Libur':x).join(', ')]));
ACT.plxImport=()=>jxImport({title:'Impor Pola Rotasi',info:'Kolom <b>Nama Pola</b> dan <b>Siklus</b>. Siklus berisi nama shift persis seperti di Master Shift atau <b>Libur</b>, dipisah koma (mis. <i>Pagi, Pagi, Siang, Libur</i>). Pola dicocokkan lewat nama (nama sama = diperbarui).',
  parse:rows=>{
    const hx=jxH(rows),ci={nm:hx('namapola','pola','nama'),cy:hx('siklus','urutan','cycle')};if(ci.nm<0||ci.cy<0)return jxBad('Kolom "Nama Pola" dan "Siklus" wajib ada pada baris judul.');
    const shifts=S.ctx.shifts||[],errs=[],items=[],seen=new Set();
    rows.slice(1).forEach((r,k)=>{const ln='Baris '+(k+2)+': ',name=normSp(r[ci.nm]);if(!name)return;
      if(seen.has(name.toLowerCase())){errs.push(ln+'nama pola ganda ('+name+')');return}seen.add(name.toLowerCase());
      const parts=String(r[ci.cy]==null?'':r[ci.cy]).split(/[,;\n]/).map(x=>x.trim()).filter(Boolean);if(!parts.length){errs.push(ln+name+': siklus kosong');return}
      const unk=parts.filter(x=>!/^libur$/i.test(x)&&!shifts.some(s=>s.name.toLowerCase()===x.toLowerCase()));if(unk.length){errs.push(ln+name+': nama shift tidak dikenal ('+[...new Set(unk)].join(', ')+')');return}
      items.push({ex:(S._patterns||[]).find(p=>p.name.toLowerCase()===name.toLowerCase()),row:{name,cycle:parts.map(x=>/^libur$/i.test(x)?null:shifts.find(s=>s.name.toLowerCase()===x.toLowerCase()).name)}})});
    const nNew=items.filter(x=>!x.ex).length;
    return{n:items.length,html:(items.length?'<div class="sk-ok">'+items.length+' pola siap diimpor: '+nNew+' baru, '+(items.length-nNew)+' diperbarui.</div>':'')+jxErrs(errs,'baris dilewati')
      +(items.length?tbl(['Pola','Aksi','Siklus'],items.slice(0,8).map(x=>['<b>'+esc(x.row.name)+'</b>',x.ex?bd('Perbarui','b'):bd('Baru','g'),esc(x.row.cycle.map(c=>c==null?'Libur':c).join(' → '))]),{page:8}):''),
      apply:async()=>{for(const it of items){if(it.ex)await DB.update('shift_patterns',it.row,{id:it.ex.id});else await DB.insert('shift_patterns',[it.row])}
        Ctx.reset();await render(true);return items.length+' pola diimpor'}}}});
/* ---- Jadwal Shift: penyusun bulanan (hasil impor masuk sebagai perubahan, belum tersimpan) ---- */
ACT.spExport=()=>{const y=ymStart(),days=D.range(y,D.mEnd(y)),mem=SP.members().map(id=>empIx()[id]).filter(Boolean).filter(e=>efPass('sp',e)).sort((a,b)=>a.full_name.localeCompare(b.full_name));
  exportRows('jadwal-shift-'+y.slice(0,7),'Jadwal Shift',['NIP','Pegawai','Departemen',...days],mem.map(e=>[String(e.employee_no),e.full_name,deptName(e.department_id)==='-'?'':deptName(e.department_id),...days.map(ds=>{const v=SP.val(e.id,ds);return v==='off'?'Libur':v?(shiftById(v)||{name:'?'}).name:''})]))};
ACT.spImport=()=>{const y=ymStart(),from=y,to=D.mEnd(y);
  jxImport({title:'Impor Jadwal Shift - '+fmtM(y),info:'Gunakan berkas hasil <b>Ekspor</b> sebagai acuan: kolom <b>NIP</b> (atau Pegawai) dan satu kolom per tanggal (YYYY-MM-DD) pada bulan yang sedang dibuka. Isi sel dengan nama shift persis seperti di Master Shift, <b>Libur</b>, atau <b>-</b> untuk menghapus jadwal; sel kosong = tidak diubah. Hasil impor masuk sebagai perubahan (bingkai oranye) dan baru tersimpan setelah Anda klik <b>Simpan draf</b> atau <b>Terbitkan</b>.',
  parse:rows=>{
    const h=rows[0].map(normKey),ni=h.findIndex(x=>x==='nip'||x==='nomorinduk'||x==='noinduk'),pi=h.findIndex(x=>x==='pegawai'||x==='nama'||x==='namalengkap');
    if(ni<0&&pi<0)return jxBad('Kolom NIP atau Pegawai tidak ditemukan pada baris judul.');
    const cols=[];let outside=0;rows[0].forEach((x,i)=>{if(i===ni||i===pi)return;const raw=String(x==null?'':x).trim();if(!raw||normKey(raw)==='departemen')return;let d=jxDate(raw);if(!d){const m=raw.match(/^(\d{1,2})(\s|$)/);if(m&&+m[1]>=1&&+m[1]<=31)d=y.slice(0,8)+jxPad(m[1])}if(!d)return;if(d<from||d>to){outside++;return}cols.push({i,d})});
    if(!cols.length)return jxBad('Tidak ada kolom tanggal pada bulan '+fmtM(y)+'. Pastikan judul kolom berupa tanggal (YYYY-MM-DD) dan bulannya sama dengan yang sedang dibuka.');
    const shifts=(S.ctx.shifts||[]).filter(s=>s.is_active),errs=[],ch=[],emps=new Set();let locked=0,same=0;
    rows.slice(1).forEach((r,k)=>{const ln='Baris '+(k+2)+': ',nip=ni>=0?String(r[ni]||'').trim():'',nm=pi>=0?String(r[pi]||'').trim():'';if(!nip&&!nm)return;
      const e=jxEmp(nip,nm);if(!e||!e.is_active){errs.push(ln+'pegawai aktif tidak ditemukan ('+(nip||nm)+')');return}
      cols.forEach(({i,d})=>{const raw=String(r[i]==null?'':r[i]).trim();if(raw==='')return;let v;
        if(/^(libur|off)$/i.test(raw))v='off';else if(raw==='-')v='';else{const s=shifts.find(x=>x.name.toLowerCase()===raw.toLowerCase());if(!s){errs.push(ln+e.full_name+' '+d+': shift "'+raw+'" tidak dikenal');return}v=String(s.id)}
        if(isLockedDate(d)){locked++;return}
        const b=SP.base(e.id,d),bv=b?(b.shift_id?String(b.shift_id):'off'):'';
        if(v===bv||(v===''&&!b)){same++;return}
        ch.push({id:e.id,d,v});emps.add(e.id)})});
    return{n:ch.length,html:(ch.length?'<div class="sk-ok">'+ch.length+' sel jadwal akan diubah untuk '+emps.size+' pegawai.</div>':'<div class="sk-warn">Tidak ada perubahan yang perlu diimpor.</div>')
      +'<div class="sk-sub" style="margin-bottom:8px">'+same+' sel sudah sama dengan jadwal saat ini'+(outside?', '+outside+' kolom di luar bulan ini dilewati':'')+(locked?', '+locked+' sel pada bulan terkunci dilewati':'')+'.</div>'+jxErrs(errs,'masalah ditemukan'),
      apply:async()=>{ch.forEach(x=>{S.sp.dirty[x.id+'|'+x.d]=x.v});S.sp.members=[...new Set([...SP.members(),...emps])];await render();return ch.length+' perubahan dimasukkan. Periksa lalu klik Simpan draf atau Terbitkan.'}}}})};
S.jn={tab:'jadwal'};
reg('jadwal',{
  range:()=>[TODAY(),TODAY()],
  render(){
    const c=S.ctx,tabs=stabs([['jadwal','Daftar Jadwal Kerja'],['tugas','Penugasan ke Pegawai']],S.jn.tab,'jnTab'),t=TODAY();
    if(S.jn.tab==='tugas'){
      efFix('jn',c.employees.filter(e=>e.is_active));const rows=c.employees.filter(e=>e.is_active&&efPass('jn',e)).map(e=>{const a=curAssign(e.id,t),ws=a?c.work_schedules.find(w=>w.id===a.work_schedule_id):null;
        return[esc(e.full_name),esc(deptName(e.department_id)),ws?bd(ws.name,'b'):(hasShifts(e.id)?bd('Jadwal shift','v'):bd('Belum ada','a')),a?fmtD(a.effective_from):'-','<button class="sk-b sec sm" data-act="jnAssign" data-e="'+e.id+'">'+ic('edit','sm')+'Atur</button>']});
      return tabs+card('Jadwal kerja tiap pegawai (per '+fmtD(t)+')','<div id="jn-tbl">'+tbl(['Pegawai','Departemen','Jadwal Normal','Berlaku sejak',''],rows,{maxh:'64vh',cls:'sk-fit',ef:'jn'})+'</div>','<div class="sk-tool">'+efBar('jn',c.employees.filter(e=>e.is_active),{inl:1})+'<input class="sk-in" type="search" placeholder="Cari pegawai atau departemen" aria-label="Cari pegawai" data-inp="jnQ"><button class="sk-b sec" data-act="jnaExport">'+ic('download','sm')+'Ekspor</button><button class="sk-b sec" data-act="jnaImport">'+ic('upload','sm')+'Impor</button><button class="sk-b" data-act="jnBulk">'+ic('users','sm')+'Terapkan ke banyak pegawai</button></div>')+'<div class="sk-info">Pegawai yang punya jadwal shift terbit pada suatu tanggal otomatis memakai jadwal shift pada tanggal itu. Selain itu berlaku jadwal normal. Perubahan jadwal mulai berlaku pada tanggal efektif; riwayat sebelumnya tetap.</div>';
    }
    const act=c.employees.filter(e=>e.is_active),normal=act.filter(e=>curAssign(e.id,t)).length,onShift=act.filter(e=>!curAssign(e.id,t)&&hasShifts(e.id)).length,noSch=act.length-normal-onShift;
    const toMin=x=>{const h=D.hm(x).split(':');return(+h[0])*60+(+h[1])};
    const hrs=w=>{const by={};(c.work_schedule_days||[]).filter(d=>d.work_schedule_id===w.id).forEach(d=>by[d.weekday]=d);const groups=[];let wk=0,nw=0;
      for(let i=1;i<=7;i++){const d=by[i];if(!(d&&d.is_workday&&d.start_time))continue;nw++;wk+=Math.max(0,toMin(d.end_time)-toMin(d.start_time)-(d.break_minutes||0));
        const key=D.hm(d.start_time)+' - '+D.hm(d.end_time)+(d.break_minutes?' (istirahat '+d.break_minutes+' mnt)':''),g=groups[groups.length-1];
        if(g&&g.key===key&&g.to===i-1)g.to=i;else groups.push({key,from:i,to:i})}
      return{groups,wk,nw}};
    const wsRows=c.work_schedules.map(w=>{const n=new Set(c.schedule_assignments.filter(a=>a.work_schedule_id===w.id&&String(a.effective_from).slice(0,10)<=t&&(!a.effective_to||String(a.effective_to).slice(0,10)>=t)).map(a=>a.employee_id)).size,h=hrs(w);
      return['<b>'+esc(w.name)+'</b>'+(w.is_active?'':' '+bd('Nonaktif','s')),
        h.groups.length?h.groups.map(g=>'<div><b>'+(g.from===g.to?DAYS3[g.from-1]:DAYS3[g.from-1]+' - '+DAYS3[g.to-1])+'</b> '+esc(g.key)+'</div>').join(''):'<span class="sk-sub">Tidak ada hari kerja</span>',
        nf(Math.round(h.wk/6)/10)+' jam',w.tolerance_minutes+' mnt',
        '<span class="sk-sub">Masuk dibuka '+w.early_clock_in_minutes+' mnt sebelum, pulang diterima '+w.late_clock_out_minutes+' mnt sesudah</span>',
        n+' pegawai',
        '<div class="sk-row" style="gap:5px"><button class="sk-b sec sm" data-act="jnEdit" data-id="'+w.id+'">'+ic('edit','sm')+'Ubah</button><button class="sk-ibtn sk-xs" data-act="jnDel" data-id="'+w.id+'" title="Hapus jadwal" aria-label="Hapus jadwal">'+ic('trash','sm')+'</button></div>']});
    const wsTable=tbl(['Jadwal','Hari & jam kerja',{t:'Jam/minggu',c:'num'},{t:'Toleransi',c:'num'},'Batas absen',{t:'Pegawai',c:'num'},''],wsRows,{empty:'Belum ada jadwal kerja. Klik Tambah Jadwal untuk membuat yang pertama.',cls:'sk-fit'});
    return tabs+'<div class="sk-kpis">'+kpi('clock','#1d4ed8',c.work_schedules.filter(w=>w.is_active).length,'Jadwal aktif','dari '+c.work_schedules.length+' jadwal')+kpi('users','#059669',normal,'Pegawai jadwal normal')+kpi('swap','#7c3aed',onShift,'Pegawai jadwal shift')+kpi('alert',noSch?'#d97706':'#64748b',noSch,'Belum punya jadwal','','')+'</div>'
      +'<div class="sk-row sp" style="margin-bottom:14px"><span class="sk-sub">Jadwal kerja tetap mingguan (mis. Guru, Tata Usaha). Ubah jam lewat tombol Ubah, atur siapa yang memakainya di tab Penugasan.</span><div class="sk-row"><button class="sk-b sec" data-act="jnExport">'+ic('download','sm')+'Ekspor</button><button class="sk-b sec" data-act="jnImport">'+ic('upload','sm')+'Impor</button><button class="sk-b" data-act="jnEdit">'+ic('plus','sm')+'Tambah Jadwal</button></div></div>'+card('Daftar Jadwal Kerja ('+c.work_schedules.length+')',wsTable);
  }
});
INP.jnQ=t=>{const q=t.value.trim().toLowerCase();if(q)$$('#jn-tbl .sk-more button').forEach(b=>{const st=TBL_REG.get(b.dataset.id);if(st)st.per=st.r.length,tblMore(b)});$$('#jn-tbl tbody tr').forEach(r=>{r.style.display=!q||r.textContent.toLowerCase().includes(q)?'':'none'})};
ACT.jnTab=async t=>{S.jn.tab=t.dataset.k;await render()};
ACT.jnEdit=t=>{
  const w=t.dataset.id?S.ctx.work_schedules.find(x=>x.id===+t.dataset.id):null,days={};(S.ctx.work_schedule_days||[]).filter(d=>w&&d.work_schedule_id===w.id).forEach(d=>days[d.weekday]=d);
  const rows=[1,2,3,4,5,6,7].map(n=>{const d=days[n]||{is_workday:n<=5,start_time:'07:00',end_time:'15:00',break_minutes:30};return'<tr><td><b>'+dayLabels[n-1]+'</b></td><td class="ctr"><input type="checkbox" name="w'+n+'"'+(d.is_workday?' checked':'')+'></td><td><input class="sk-in" type="time" name="s'+n+'" value="'+D.hm(d.start_time||'07:00')+'"></td><td><input class="sk-in" type="time" name="e'+n+'" value="'+D.hm(d.end_time||'15:00')+'"></td><td><input class="sk-in" type="number" min="0" name="b'+n+'" value="'+(d.break_minutes||0)+'" style="width:80px"></td></tr>'}).join('');
  modal({title:w?'Ubah Jadwal Kerja':'Tambah Jadwal Kerja',size:'wide',body:'<form id="jf"><div class="sk-grid sk-g2">'+fld('Nama jadwal',inp('name',w?w.name:'',{req:1,ph:'mis. Guru'}))+fld('Toleransi terlambat (menit)',inp('tol',w?w.tolerance_minutes:10,{type:'number',min:0}))+fld('Absen masuk dibuka (menit sebelum jam masuk)',inp('early',w?w.early_clock_in_minutes:120,{type:'number',min:0}))+fld('Absen pulang diterima (menit setelah jam pulang)',inp('late',w?w.late_clock_out_minutes:240,{type:'number',min:0}))+'</div>'
    +'<div class="sk-tw"><table class="sk-t"><thead><tr><th>Hari</th><th class="ctr">Kerja</th><th>Masuk</th><th>Pulang</th><th>Istirahat (mnt)</th></tr></thead><tbody>'+rows+'</tbody></table></div>'+(w?chk('is_active','Jadwal aktif',w.is_active):'')+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{
      const d=formData($('#jf',m.el));if(!d.name.trim()){toast('Nama wajib diisi','warn');return false}
      const row={name:d.name.trim(),tolerance_minutes:+d.tol||0,early_clock_in_minutes:+d.early||0,late_clock_out_minutes:+d.late||0,is_active:w?!!d.is_active:true};
      const dr=[];for(let n=1;n<=7;n++){if(d['w'+n]&&(!d['s'+n]||!d['e'+n])){toast('Jam masuk dan pulang '+dayLabels[n-1]+' wajib diisi','warn');return false}dr.push({weekday:n,is_workday:!!d['w'+n],start_time:d['w'+n]?d['s'+n]:null,end_time:d['w'+n]?d['e'+n]:null,break_minutes:+d['b'+n]||0})}
      let id=w?w.id:null;if(id)await DB.update('work_schedules',row,{id});else{const r=await DB.insert('work_schedules',[row]);id=r[0].id}
      await DB.upsert('work_schedule_days',dr.map(x=>({...x,work_schedule_id:id})),'work_schedule_id,weekday');await audit('schedule.save','work_schedules',id,w,row);toast('Jadwal kerja tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.jnDel=async t=>{const w=S.ctx.work_schedules.find(x=>x.id===+t.dataset.id);if(S.ctx.schedule_assignments.some(a=>a.work_schedule_id===w.id)){toast('Jadwal masih dipakai penugasan pegawai. Nonaktifkan saja atau pindahkan pegawainya terlebih dahulu.','bad');return}
  if(!await confirmBox('Hapus jadwal "'+w.name+'"?',{danger:true,ok:'Hapus'}))return;await DB.remove('work_schedules',{id:w.id});await audit('schedule.delete','work_schedules',w.id,w,null);toast('Jadwal dihapus','ok');Ctx.reset();await render(true)};
ACT.jnAssign=t=>{
  const e=S.ctx.employees.find(x=>x.id===+t.dataset.e),a=curAssign(e.id,TODAY());
  modal({title:'Jadwal kerja - '+e.full_name,body:'<form id="jaf">'+fld('Jadwal normal',sel('ws',S.ctx.work_schedules.filter(w=>w.is_active).map(w=>[w.id,w.name]),a?a.work_schedule_id:''))+fld('Berlaku mulai tanggal',inp('from',TODAY(),{type:'date',req:1}),'Riwayat sebelum tanggal ini tidak berubah')+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const d=formData($('#jaf',m.el));await assignSchedule([e.id],+d.ws,d.from);await audit('schedule.assign','schedule_assignments',e.id,null,d);if(d.from<=TODAY())await reprocess(d.from,TODAY(),[e.id]);toast('Penugasan tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.jnBulk=()=>{
  const emps=S.ctx.employees.filter(e=>e.is_active);
  modal({title:'Terapkan Jadwal ke Banyak Pegawai',size:'wide',body:'<form id="jbf"><div class="sk-grid sk-g3">'+fld('Jadwal normal',sel('ws',S.ctx.work_schedules.filter(w=>w.is_active).map(w=>[w.id,w.name]),''))+fld('Berlaku mulai',inp('from',TODAY(),{type:'date',req:1}))+fld('Pilih cepat',sel('quick',[['','- pilih -'],['all','Semua pegawai'],['teacher','Semua guru'],['staff','Semua karyawan umum'],['none','Kosongkan'],...deptOptions().map(d=>['d'+d[0],'Departemen: '+d[1]])],'',{attr:'data-chg="jbQuick"'}))+'</div>'+efModBar(emps)+'<div class="sk-tw" style="max-height:300px"><table class="sk-t"><tbody>'+emps.map(e=>'<tr'+efTr(e)+'><td style="width:34px"><input type="checkbox" class="jbc" value="'+e.id+'" data-t="'+e.type+'" data-d="'+e.department_id+'"></td><td>'+esc(e.full_name)+'</td><td>'+esc(deptName(e.department_id))+'</td><td>'+esc(TYPE[e.type])+'</td></tr>').join('')+'</tbody></table></div></form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Terapkan',i:'check',fn:async m=>{const d=formData($('#jbf',m.el)),ids=$$('.jbc:checked',m.el).map(x=>+x.value);if(!ids.length){toast('Pilih minimal satu pegawai','warn');return false}
      await assignSchedule(ids,+d.ws,d.from);await audit('schedule.bulk','schedule_assignments',ids.length+' pegawai',null,d);if(d.from<=TODAY())await reprocess(d.from,TODAY(),ids);toast('Jadwal diterapkan ke '+ids.length+' pegawai','ok');Ctx.reset();render(true)}}]});
};
CHG.jbQuick=t=>{const v=t.value;$$('.jbc',t.closest('.sk-md')).forEach(c=>{c.checked=v==='all'||(v==='teacher'&&c.dataset.t==='teacher')||(v==='staff'&&c.dataset.t==='staff')||(v[0]==='d'&&c.dataset.d===v.slice(1))})};

/* =====================================================================
   19. JADWAL SHIFT: master, pola, penyusun (dipakai Admin / HRD)
   ===================================================================== */
S.sp={members:null,dirty:{},tab:'grid',lastYm:null};
const shiftById=id=>(S.ctx.shifts||[]).find(s=>s.id===+id);
const SP={
  base(emp,ds){return(S.ctx.shift_assignments||[]).find(a=>a.employee_id===emp&&String(a.work_date).slice(0,10)===ds)},
  members(){if(!S.sp.members){S.sp.members=[...new Set((S.ctx.shift_assignments||[]).map(a=>a.employee_id))].filter(id=>empIx()[id])}return S.sp.members},
  val(emp,ds){const k=emp+'|'+ds;if(k in S.sp.dirty)return S.sp.dirty[k];const a=this.base(emp,ds);return a?(a.shift_id?String(a.shift_id):'off'):''},
  plan(from,to){const p={};for(const e of this.members())for(const ds of D.range(from,to)){const v=this.val(e,ds);if(v!=='')p[e+'|'+ds]=v==='off'?null:+v}return p},
  warnings(from,to){return Engine.shiftWarnings(S.ctx,this.plan(from,to),from,to)},
  warnHtml(from,to){const w=this.warnings(from,to);return w.length?'<div class="sk-warn"><b>'+w.length+' peringatan jadwal:</b><br>'+w.slice(0,25).map(esc).join('<br>')+(w.length>25?'<br>... dan '+(w.length-25)+' lainnya':'')+'</div>':'<div class="sk-ok">Tidak ada peringatan: jeda antar shift, hari berturut-turut, dan minimal petugas terpenuhi.</div>'},
  dirtyCount(){return Object.keys(S.sp.dirty).length},
  grid(from,to,canEdit){
    efFix('sp',this.members().map(id=>empIx()[id]).filter(Boolean));const mem=this.members().map(id=>empIx()[id]).filter(Boolean).filter(e=>efPass('sp',e)).sort((a,b)=>a.full_name.localeCompare(b.full_name)),days=D.range(from,to),shifts=(S.ctx.shifts||[]).filter(s=>s.is_active);
    const opts=(v)=>'<option value=""'+(v===''?' selected':'')+'>-</option><option value="off"'+(v==='off'?' selected':'')+'>Libur</option>'+shifts.map(s=>'<option value="'+s.id+'"'+(String(s.id)===v?' selected':'')+'>'+esc(s.name)+'</option>').join('');
    if(!mem.length&&efOn('sp')&&this.members().length)return'<div class="sk-empty">Tidak ada pegawai shift yang sesuai filter departemen / divisi.</div>';
    if(!mem.length)return'<div class="sk-empty">Belum ada pegawai dalam jadwal shift. Klik "Tambah pegawai" untuk memilih petugas yang bekerja bergiliran (mis. keamanan).</div>';
    return'<div class="sk-tw" style="max-height:62vh"><table class="sk-t sk-mat"><thead><tr><th>Pegawai</th>'+days.map(ds=>'<th'+(D.iso(ds)>=6?' style="background:#fee2e2"':'')+'>'+D.parse(ds).d+'<br><small>'+dayName(ds)+'</small></th>').join('')+'</tr></thead><tbody>'
     +mem.map(e=>'<tr><td class="nm">'+esc(e.full_name)+'<div class="sk-sub" style="font-weight:400">'+esc(deptName(e.department_id))+'</div></td>'+days.map(ds=>{const v=this.val(e.id,ds),sh=v&&v!=='off'?shiftById(v):null,b=this.base(e.id,ds),k=e.id+'|'+ds,lock=isLockedDate(ds);
        return'<td style="'+(sh?'background:'+(sh.color||'#2563eb')+'22;':'')+(k in S.sp.dirty?'outline:2px solid #f59e0b;outline-offset:-2px;':(b&&b.status==='draft'?'outline:2px dashed #94a3b8;outline-offset:-2px;':''))+'">'+(canEdit&&!lock?'<select data-chg="spCell" data-k="'+k+'">'+opts(v)+'</select>':(sh?esc(sh.name):(v==='off'?'Libur':'-')))+'</td>'}).join('')+'</tr>').join('')+'</tbody></table></div>'
     +'<div class="sk-sub" style="margin-top:8px">Sel berbingkai oranye = perubahan belum disimpan. Berbingkai putus-putus = draf (belum terbit, belum terlihat pegawai). "-" = ikut jadwal normal. Bulan terkunci tidak dapat diubah.</div>';
  }
};
CHG.spCell=t=>{
  const k=t.dataset.k,[e,ds]=k.split('|'),b=SP.base(+e,ds),bv=b?(b.shift_id?String(b.shift_id):'off'):'';
  if(t.value===bv&&!(b&&b.status==='draft'&&false))delete S.sp.dirty[k];else S.sp.dirty[k]=t.value;
  const sh=t.value&&t.value!=='off'?shiftById(t.value):null,td=t.closest('td');td.style.background=sh?(sh.color||'#2563eb')+'22':'';td.style.outline=(k in S.sp.dirty)?'2px solid #f59e0b':(b&&b.status==='draft'?'2px dashed #94a3b8':'');td.style.outlineOffset='-2px';
  const y=ymStart(),w=$('#sp-warn');if(w)w.innerHTML=SP.warnHtml(y,D.mEnd(y));const n=$('#sp-dirty');if(n)n.textContent=SP.dirtyCount();
};
function spPanel(canEdit,saveFn,publishFn){
  const y=ymStart(),from=y,to=D.mEnd(y);
  if(S.sp.lastYm!==y){S.sp.lastYm=y}
  const drafts=(S.ctx.shift_assignments||[]).filter(a=>a.status==='draft'&&String(a.work_date).slice(0,10)>=from&&String(a.work_date).slice(0,10)<=to).length;
  return'<div class="sk-row sp" style="margin-bottom:12px">'+monthBar()+'<div class="sk-row">'+(canEdit?'<button class="sk-b sec" data-act="spAddMembers">'+ic('plus','sm')+'Tambah pegawai</button>'+(S.mode==='admin'?'<button class="sk-b sec" data-act="spPattern">'+ic('swap','sm')+'Terapkan pola</button>':'')+'<button class="sk-b sec" data-act="spReset">Batalkan perubahan</button><button class="sk-b sec" data-act="spSave">'+ic('check','sm')+'Simpan draf (<span id="sp-dirty">'+SP.dirtyCount()+'</span>)</button><button class="sk-b" data-act="spPublish">'+ic('upload','sm')+'Terbitkan'+(drafts?' ('+drafts+' draf)':'')+'</button>':'')+(canEdit?'<button class="sk-b sec" data-act="spImport">'+ic('upload','sm')+'Impor</button>':'')+'<button class="sk-b sec" data-act="spExport">'+ic('download','sm')+'Ekspor</button></div></div>'
   +'<div class="ef-top">'+efBar('sp',SP.members().map(id=>empIx()[id]).filter(Boolean),{nolim:1})+'</div>'+'<div id="sp-warn">'+SP.warnHtml(from,to)+'</div>'+'<section class="sk-card">'+SP.grid(from,to,canEdit)+'</section>';
}
ACT.spReset=async()=>{S.sp.dirty={};await render()};
ACT.spAddMembers=()=>{
  const cur=new Set(SP.members());
  modal({title:'Pegawai dalam Jadwal Shift',body:'<div class="sk-sub" style="margin-bottom:8px">Centang pegawai yang bekerja bergiliran. Pegawai yang dicentang akan muncul di tabel dan dapat diberi shift.</div>'+efModBar(S.ctx.employees.filter(e=>e.is_active))+'<div class="sk-tw" style="max-height:340px"><table class="sk-t"><tbody>'+S.ctx.employees.filter(e=>e.is_active).map(e=>'<tr'+efTr(e)+'><td style="width:34px"><input type="checkbox" class="smc" value="'+e.id+'"'+(cur.has(e.id)?' checked':'')+'></td><td>'+esc(e.full_name)+'</td><td>'+esc(deptName(e.department_id))+'</td></tr>').join('')+'</tbody></table></div>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Terapkan',i:'check',fn:async m=>{S.sp.members=$$('.smc:checked',m.el).map(x=>+x.value);await render()}}]});
};
ACT.spPattern=()=>{
  const pats=S.ctx.shift_patterns||S._patterns||[];
  modal({title:'Terapkan Pola Rotasi',size:'wide',body:'<form id="pf"><div class="sk-grid sk-g2">'+fld('Pola',sel('pat',(S._patterns||[]).map(p=>[p.id,p.name+' ('+(p.cycle||[]).length+' hari)']),'',{req:1}))+fld('Mulai tanggal',inp('from',ymStart(),{type:'date',req:1}))+fld('Sampai tanggal',inp('to',D.mEnd(ymStart()),{type:'date',req:1}))+fld('Geser awal antar pegawai (hari)',inp('shift',0,{type:'number'}),'Mis. 2 = pegawai berikutnya mulai 2 hari lebih jauh dalam siklus')+'</div>'+efModBar(S.ctx.employees.filter(e=>e.is_active))+'<div class="sk-tw" style="max-height:260px"><table class="sk-t"><tbody>'+S.ctx.employees.filter(e=>e.is_active).map(e=>'<tr'+efTr(e)+'><td style="width:34px"><input type="checkbox" class="pmc" value="'+e.id+'"'+(SP.members().includes(e.id)?' checked':'')+'></td><td>'+esc(e.full_name)+'</td><td>'+esc(deptName(e.department_id))+'</td></tr>').join('')+'</tbody></table></div><div class="sk-hint">Hasil dimasukkan sebagai perubahan (bingkai oranye); periksa lalu klik Simpan draf atau Terbitkan.</div></form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Terapkan',i:'check',fn:async m=>{
      const d=formData($('#pf',m.el)),pat=(S._patterns||[]).find(p=>p.id===+d.pat),ids=$$('.pmc:checked',m.el).map(x=>+x.value);
      if(!pat||!(pat.cycle||[]).length){toast('Pilih pola yang valid','warn');return false}if(!ids.length){toast('Pilih pegawai','warn');return false}if(d.to<d.from){toast('Rentang tanggal tidak valid','warn');return false}
      const byName={};(S.ctx.shifts||[]).forEach(s=>byName[s.name.toLowerCase()]=s.id);const days=D.range(d.from,d.to),n=pat.cycle.length,off=+d.shift||0;
      ids.forEach((id,k)=>days.forEach((ds,i)=>{if(isLockedDate(ds))return;const tk=pat.cycle[((i+k*off)%n+n)%n],v=tk==null||/^libur$/i.test(tk)?'off':(byName[String(tk).toLowerCase()]!=null?String(byName[String(tk).toLowerCase()]):'');const key=id+'|'+ds,b=SP.base(id,ds),bv=b?(b.shift_id?String(b.shift_id):'off'):'';if(v===bv)delete S.sp.dirty[key];else S.sp.dirty[key]=v}));
      S.sp.members=[...new Set([...SP.members(),...ids])];toast('Pola diterapkan. Periksa lalu simpan.','ok');await render()}}]});
};

/* penyimpanan: admin lewat tabel */
async function spPersist(publish){
  const y=ymStart(),from=y,to=D.mEnd(y),rows=[];
  for(const k in S.sp.dirty){const[e,ds]=k.split('|'),v=S.sp.dirty[k];rows.push({employee_id:+e,work_date:ds,shift_id:v&&v!=='off'?+v:null,remove:v===''})}
  if(publish){const w=SP.warnings(from,to);if(w.length&&!await confirmBox('Ada '+w.length+' peringatan jadwal (mis. jeda antar shift terlalu pendek). Tetap terbitkan?',{ok:'Tetap terbitkan'}))return false}
  if(S.mode==='user'){
    if(rows.length)await rpc('u_save_shifts',{p_token:S.token,p_rows:rows,p_publish:publish});
    if(publish)await rpc('u_publish_shifts',{p_token:S.token,p_from:from,p_to:to});
  }else{
    if(rows.some(r=>isLockedDate(r.work_date)))throw new Error('Ada tanggal pada periode yang sudah dikunci');
    const ts=new Date().toISOString(),up=rows.filter(r=>!r.remove).map(r=>({employee_id:r.employee_id,work_date:r.work_date,shift_id:r.shift_id,status:publish?'published':'draft',published_at:publish?ts:null}));
    if(up.length)await DB.upsert('shift_assignments',up,'employee_id,work_date');
    for(const r of rows.filter(x=>x.remove))await DB.remove('shift_assignments',{employee_id:r.employee_id,work_date:r.work_date});
    if(publish){const dr=(S.ctx.shift_assignments||[]).filter(a=>a.status==='draft'&&String(a.work_date).slice(0,10)>=from&&String(a.work_date).slice(0,10)<=to&&!(a.employee_id+'|'+String(a.work_date).slice(0,10) in S.sp.dirty));
      for(const a of dr)await DB.update('shift_assignments',{status:'published',published_at:ts},{id:a.id});
      const emps=new Set([...rows.map(r=>r.employee_id),...dr.map(a=>a.employee_id)]);for(const id of emps)await notifyEmp(id,'Jadwal shift diterbitkan','Jadwal shift '+fmtM(y)+' telah diterbitkan atau diubah. Silakan cek menu Jadwal Saya.','jadwal');
      const past=rows.filter(r=>r.work_date<=TODAY());if(past.length)await reprocess(past.map(r=>r.work_date).sort()[0],TODAY(),[...new Set(past.map(r=>r.employee_id))]);
    }
  }
  S.sp.dirty={};return true;
}
ACT.spSave=async t=>{if(!SP.dirtyCount()){toast('Tidak ada perubahan','warn');return}await busy(t,async()=>{if(await spPersist(false)!==false){toast('Draf jadwal tersimpan','ok');Ctx.reset();await render(true)}})};
ACT.spPublish=async t=>{await busy(t,async()=>{if(await spPersist(true)!==false){toast('Jadwal shift diterbitkan','ok');Ctx.reset();await render(true)}})};

/* ---------- Admin: Jadwal Shift ---------- */
S.sh={tab:'grid'};
reg('shift',{
  async render(){
    const tab=S.sh.tab;S._patterns=await DB.select('shift_patterns');
    const tabs=stabs([['grid','Penyusun Jadwal'],['master','Master Shift'],['pola','Pola Rotasi'],['tukar','Tukar Shift']],tab,'shTab');
    if(tab==='master'){
      const rows=S.ctx.shifts.map(s=>['<span style="display:inline-block;width:12px;height:12px;border-radius:4px;background:'+esc(s.color||'#2563eb')+';margin-right:6px"></span><b>'+esc(s.name)+'</b>',D.hm(s.start_time)+' - '+D.hm(s.end_time)+(s.crosses_midnight?' (lewat tengah malam)':''),s.break_minutes+' mnt',s.tolerance_minutes+' mnt',s.early_clock_in_minutes+' / '+s.late_clock_out_minutes+' mnt',s.is_active?bd('Aktif','g'):bd('Nonaktif','s'),'<div class="sk-row" style="gap:5px"><button class="sk-b sec sm" data-act="shEdit" data-id="'+s.id+'">'+ic('edit','sm')+'Ubah</button><button class="sk-b sec sm" data-act="shDel" data-id="'+s.id+'">'+ic('trash','sm')+'</button></div>']);
      return tabs+card('Master Shift',tbl(['Shift','Jam','Istirahat','Toleransi','Buka masuk / tutup pulang',' Status',''],rows,{empty:'Belum ada shift'}),'<div class="sk-row"><button class="sk-b sec" data-act="shxExport">'+ic('download','sm')+'Ekspor</button><button class="sk-b sec" data-act="shxImport">'+ic('upload','sm')+'Impor</button><button class="sk-b" data-act="shEdit">'+ic('plus','sm')+'Tambah Shift</button></div>');
    }
    if(tab==='pola'){
      const rows=S._patterns.map(p=>['<b>'+esc(p.name)+'</b>',esc((p.cycle||[]).map(x=>x==null?'Libur':x).join(' → ')),(p.cycle||[]).length+' hari','<div class="sk-row" style="gap:5px"><button class="sk-b sec sm" data-act="plEdit" data-id="'+p.id+'">'+ic('edit','sm')+'Ubah</button><button class="sk-b sec sm" data-act="plDel" data-id="'+p.id+'">'+ic('trash','sm')+'</button></div>']);
      return tabs+card('Pola Rotasi Shift',tbl(['Nama','Urutan siklus','Panjang',''],rows,{empty:'Belum ada pola'}),'<div class="sk-row"><button class="sk-b sec" data-act="plxExport">'+ic('download','sm')+'Ekspor</button><button class="sk-b sec" data-act="plxImport">'+ic('upload','sm')+'Impor</button><button class="sk-b" data-act="plEdit">'+ic('plus','sm')+'Tambah Pola</button></div>')+'<div class="sk-info">Pola dipakai di Penyusun Jadwal lewat tombol "Terapkan pola" untuk mengisi satu bulan sekaligus.</div>';
    }
    if(tab==='tukar')return tabs+swapAdminPanel();
    return tabs+spPanel(true);
  }
});
ACT.shTab=async t=>{S.sh.tab=t.dataset.k;await render()};
ACT.shEdit=t=>{
  const s=t.dataset.id?S.ctx.shifts.find(x=>x.id===+t.dataset.id):null;
  modal({title:s?'Ubah Shift':'Tambah Shift',body:'<form id="sf2"><div class="sk-grid sk-g2">'+fld('Nama shift',inp('name',s?s.name:'',{req:1}))+fld('Warna',inp('color',s?s.color||'#2563eb':'#2563eb',{type:'color'}))+fld('Jam masuk',inp('start_time',s?D.hm(s.start_time):'07:00',{type:'time',req:1}))+fld('Jam pulang',inp('end_time',s?D.hm(s.end_time):'15:00',{type:'time',req:1}))+fld('Istirahat (menit)',inp('break_minutes',s?s.break_minutes:30,{type:'number',min:0}))+fld('Toleransi terlambat (menit)',inp('tolerance_minutes',s?s.tolerance_minutes:10,{type:'number',min:0}))+fld('Absen masuk dibuka (menit sebelum)',inp('early',s?s.early_clock_in_minutes:120,{type:'number',min:0}))+fld('Absen pulang diterima (menit sesudah)',inp('late',s?s.late_clock_out_minutes:240,{type:'number',min:0}))+'</div><div class="sk-hint">Bila jam pulang lebih kecil atau sama dengan jam masuk, shift otomatis dianggap melewati tengah malam.</div>'+(s?chk('is_active','Shift aktif',s.is_active):'')+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const d=formData($('#sf2',m.el));if(!d.name.trim()){toast('Nama wajib diisi','warn');return false}
      const row={name:d.name.trim(),start_time:d.start_time,end_time:d.end_time,crosses_midnight:d.end_time<=d.start_time,break_minutes:+d.break_minutes||0,tolerance_minutes:+d.tolerance_minutes||0,early_clock_in_minutes:+d.early||0,late_clock_out_minutes:+d.late||0,color:d.color,is_active:s?!!d.is_active:true};
      if(s)await DB.update('shifts',row,{id:s.id});else await DB.insert('shifts',[row]);await audit('shift.save','shifts',s?s.id:row.name,s,row);toast('Shift tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.shDel=async t=>{const s=S.ctx.shifts.find(x=>x.id===+t.dataset.id);if(!await confirmBox('Hapus shift "'+s.name+'"? Shift yang sudah dipakai pada jadwal tidak dapat dihapus (nonaktifkan saja).',{danger:true,ok:'Hapus'}))return;
  try{await DB.remove('shifts',{id:s.id});await audit('shift.delete','shifts',s.id,s,null);toast('Shift dihapus','ok');Ctx.reset();await render(true)}catch(e){toast('Shift sedang dipakai pada jadwal. Nonaktifkan saja.','bad')}};
ACT.plEdit=t=>{
  const p=t.dataset.id?S._patterns.find(x=>x.id===+t.dataset.id):null;
  modal({title:p?'Ubah Pola':'Tambah Pola',body:'<form id="plf">'+fld('Nama pola',inp('name',p?p.name:'',{req:1}))+fld('Urutan siklus',inp('cycle',p?(p.cycle||[]).map(x=>x==null?'Libur':x).join(', '):'Pagi, Pagi, Siang, Siang, Malam, Malam, Libur',{req:1}),'Pisahkan dengan koma. Gunakan nama shift persis seperti di Master Shift, atau "Libur". Contoh: Pagi, Pagi, Siang, Malam, Libur')+'<div class="sk-sub">Shift tersedia: '+esc(S.ctx.shifts.map(s=>s.name).join(', '))+'</div></form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const d=formData($('#plf',m.el)),names=S.ctx.shifts.map(s=>s.name.toLowerCase());
      const cyc=d.cycle.split(',').map(x=>x.trim()).filter(Boolean).map(x=>/^libur$/i.test(x)?null:(S.ctx.shifts.find(s=>s.name.toLowerCase()===x.toLowerCase())||{name:x}).name);
      const bad=d.cycle.split(',').map(x=>x.trim()).filter(x=>x&&!/^libur$/i.test(x)&&!names.includes(x.toLowerCase()));if(bad.length){toast('Nama shift tidak dikenal: '+bad.join(', '),'bad');return false}if(!cyc.length){toast('Siklus kosong','warn');return false}
      const row={name:d.name.trim(),cycle:cyc};if(p)await DB.update('shift_patterns',row,{id:p.id});else await DB.insert('shift_patterns',[row]);toast('Pola tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.plDel=async t=>{if(!await confirmBox('Hapus pola ini?',{danger:true,ok:'Hapus'}))return;await DB.remove('shift_patterns',{id:+t.dataset.id});toast('Pola dihapus','ok');await render(true)};

/* ---------- tukar shift: tampilan keputusan (admin / HRD) ---------- */
function swapRow(s,actions){
  const sh=(emp,ds)=>{const a=(S.ctx.shift_assignments||[]).find(x=>x.employee_id===emp&&String(x.work_date).slice(0,10)===ds&&x.shift_id);const x=a?shiftById(a.shift_id):null;return x?x.name:'-'};
  return[esc(nameOf(s.requester_employee_id))+'<div class="sk-sub">'+fmtD(s.requester_date)+' ('+esc(sh(s.requester_employee_id,s.requester_date))+')</div>',esc(nameOf(s.partner_employee_id))+'<div class="sk-sub">'+fmtD(s.partner_date)+' ('+esc(sh(s.partner_employee_id,s.partner_date))+')</div>',esc(s.reason||'-'),(SWST[s.status]?bd(SWST[s.status][0],SWST[s.status][1]):esc(s.status))+(s.decided_by?'<div class="sk-sub">'+esc(s.decided_by)+'</div>':''),actions||''];
}
function swapAdminPanel(){
  const swp=S.ctx.employees.filter(e=>S.ctx.swaps.some(s=>s.requester_employee_id===e.id||s.partner_employee_id===e.id));efFix('sw',swp);
  const swl=S.ctx.swaps.filter(s=>efPassId('sw',s.requester_employee_id)||efPassId('sw',s.partner_employee_id)),wait=swl.filter(s=>s.status==='waiting_coordinator'),rest=swl.filter(s=>s.status!=='waiting_coordinator');
  const btn=s=>'<div class="sk-row" style="gap:5px"><button class="sk-b ok sm" data-act="swDecide" data-id="'+s.id+'" data-d="approved">Setujui</button><button class="sk-b bad sm" data-act="swDecide" data-id="'+s.id+'" data-d="rejected">Tolak</button></div>';
  return'<div class="ef-top">'+efBar('sw',swp)+'</div>'+card('Menunggu keputusan ('+wait.length+')',tbl(['Pengaju','Pengganti','Alasan','Status',''],wait.map(s=>swapRow(s,btn(s))),{empty:'Tidak ada permintaan menunggu',ef:'sw'}))+card('Riwayat tukar shift',tbl(['Pengaju','Pengganti','Alasan','Status',''],rest.slice(0,60).map(s=>swapRow(s,'')),{empty:'Belum ada riwayat',ef:'sw'}))
   +'<div class="sk-info">Alur: pegawai mengajukan tukar → pegawai pengganti menyetujui → HRD memutuskan. Saat disetujui, jadwal kedua pegawai langsung ditukar.</div>';
}
ACT.swDecide=async t=>{
  const id=+t.dataset.id,d=t.dataset.d;if(!await confirmBox(d==='approved'?'Setujui tukar shift ini? Jadwal kedua pegawai akan langsung ditukar.':'Tolak tukar shift ini?',{ok:d==='approved'?'Setujui':'Tolak',danger:d!=='approved'}))return;
  await busy(t,async()=>{
    const r=S.mode==='admin'?await rpc('adm_swap_decide',{p_id:id,p_decision:d}):await rpc('u_swap_decide',{p_token:S.token,p_id:id,p_decision:d});
    if(S.mode==='admin'&&d==='approved'&&r.dates&&r.employees){const lo=r.dates.slice().sort()[0];if(lo<=TODAY())for(const e of r.employees)await applyRange(e,lo,r.dates.slice().sort()[1])}
    toast('Keputusan tersimpan','ok');Ctx.reset();await render(true)});
};

/* =====================================================================
   20. ADMIN: KALENDER LIBUR
   ===================================================================== */
S.hl={year:null};
reg('libur',{
  range:()=>[TODAY(),TODAY()],
  async render(){
    const y=S.hl.year||(S.hl.year=D.parse(TODAY()).y),all=await DB.select('holidays'),list=all.filter(h=>String(h.holiday_date).slice(0,4)===String(y)).sort((a,b)=>String(a.holiday_date).localeCompare(String(b.holiday_date)));
    const rows=list.map(h=>[fmtDL(String(h.holiday_date).slice(0,10)),esc(h.name),bd(HOLT[h.type]||h.type,h.type==='national'?'r':h.type==='collective_leave'?'a':'b'),{all:'Semua',teacher:'Guru saja',staff:'Karyawan umum saja'}[h.applies_to]||h.applies_to,'<div class="sk-row" style="gap:5px"><button class="sk-b sec sm" data-act="hlEdit" data-id="'+h.id+'">'+ic('edit','sm')+'Ubah</button><button class="sk-b sec sm" data-act="hlDel" data-id="'+h.id+'">'+ic('trash','sm')+'</button></div>']);
    S._hol=all;
    const byM={};list.forEach(h=>{const m=D.parse(String(h.holiday_date).slice(0,10)).m;byM[m]=(byM[m]||0)+1});
    return'<div class="sk-row sp" style="margin-bottom:12px"><div class="sk-row" style="gap:6px"><button class="sk-ibtn" data-act="hlYear" data-d="-1">'+ic('left')+'</button><b style="min-width:70px;text-align:center;font-size:calc(18px*var(--fs,1))">'+y+'</b><button class="sk-ibtn" data-act="hlYear" data-d="1">'+ic('right')+'</button></div><div class="sk-row"><button class="sk-b sec" data-act="hlSeed">'+ic('download','sm')+'Isi libur tetap '+y+'</button><button class="sk-b sec" data-act="hlRange">'+ic('calendar','sm')+'Tambah rentang libur lembaga</button><button class="sk-b" data-act="hlEdit">'+ic('plus','sm')+'Tambah Libur</button></div></div>'
     +card('Hari libur '+y+' ('+list.length+')',tbl(['Tanggal','Keterangan','Jenis','Berlaku untuk',''],rows,{empty:'Belum ada hari libur tahun ini'}))
     +'<div class="sk-info">Hari libur nasional dan cuti bersama berubah tiap tahun (terutama hari raya keagamaan), jadi tanggalnya perlu diisi sesuai SKB resmi. Tombol "Isi libur tetap" hanya mengisi yang tanggalnya tidak berubah (1 Jan, 1 Mei, 1 Jun, 17 Agu, 25 Des). Libur menghilangkan kewajiban hadir dan tidak memotong saldo cuti.</div>';
  }
});
ACT.hlYear=async t=>{S.hl.year+=+t.dataset.d;await render()};
async function holidaysChanged(d1,d2){const t=TODAY();if(d1<=t){try{await reprocess(d1,d2>t?t:d2)}catch(e){console.warn(e)}}}
ACT.hlEdit=t=>{
  const h=t.dataset.id?S._hol.find(x=>x.id===+t.dataset.id):null;
  modal({title:h?'Ubah Hari Libur':'Tambah Hari Libur',body:'<form id="hf">'+fld('Tanggal',inp('holiday_date',h?String(h.holiday_date).slice(0,10):'',{type:'date',req:1}))+fld('Keterangan',inp('name',h?h.name:'',{req:1}))+fld('Jenis',sel('type',Object.entries(HOLT),h?h.type:'national'))+fld('Berlaku untuk',sel('applies_to',[['all','Semua pegawai'],['teacher','Guru saja'],['staff','Karyawan umum saja']],h?h.applies_to:'all'))+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const d=formData($('#hf',m.el));if(!d.holiday_date||!d.name.trim()){toast('Tanggal dan keterangan wajib diisi','warn');return false}
      const row={holiday_date:d.holiday_date,name:d.name.trim(),type:d.type,applies_to:d.applies_to};
      try{if(h)await DB.update('holidays',row,{id:h.id});else await DB.insert('holidays',[row])}catch(e){if(/duplicate|unique/i.test(e.message)){toast('Sudah ada libur pada tanggal dan cakupan yang sama','bad');return false}throw e}
      await audit('holiday.save','holidays',d.holiday_date,h,row);await holidaysChanged(d.holiday_date,d.holiday_date);toast('Hari libur tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.hlDel=async t=>{const h=S._hol.find(x=>x.id===+t.dataset.id);if(!await confirmBox('Hapus libur "'+h.name+'"?',{danger:true,ok:'Hapus'}))return;await DB.remove('holidays',{id:h.id});await audit('holiday.delete','holidays',h.id,h,null);await holidaysChanged(String(h.holiday_date).slice(0,10),String(h.holiday_date).slice(0,10));toast('Dihapus','ok');Ctx.reset();await render(true)};
ACT.hlSeed=async()=>{const y=S.hl.year;const have=new Set(S._hol.filter(h=>h.applies_to==='all').map(h=>String(h.holiday_date).slice(0,10)));const rows=FIXED_HOLIDAYS.map(([md,n])=>({holiday_date:y+'-'+md,name:n,type:'national',applies_to:'all'})).filter(r=>!have.has(r.holiday_date));
  if(!rows.length){toast('Libur tetap sudah lengkap','warn');return}await DB.insert('holidays',rows);toast(rows.length+' hari libur ditambahkan','ok');Ctx.reset();await render(true)};
ACT.hlRange=()=>{
  modal({title:'Tambah Rentang Libur Lembaga',body:'<form id="hrf">'+fld('Nama libur',inp('name','Libur Semester',{req:1}))+'<div class="sk-grid sk-g2">'+fld('Mulai',inp('from',TODAY(),{type:'date',req:1}))+fld('Sampai',inp('to',D.add(TODAY(),13),{type:'date',req:1}))+'</div>'+fld('Jenis',sel('type',[['school','Libur Lembaga'],['academic_break','Libur Akademik'],['collective_leave','Cuti Bersama']],'academic_break'))+fld('Berlaku untuk',sel('applies_to',[['all','Semua pegawai'],['teacher','Guru saja'],['staff','Karyawan umum saja']],'teacher'),'Libur semester biasanya hanya untuk guru; karyawan umum tetap bekerja')+'<div class="sk-hint">Setiap tanggal pada rentang dibuat sebagai satu hari libur. Maksimal 120 hari.</div></form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const d=formData($('#hrf',m.el));if(d.to<d.from||D.diff(d.from,d.to)>120){toast('Rentang tidak valid (maks. 120 hari)','warn');return false}
      const have=new Set(S._hol.filter(h=>h.applies_to===d.applies_to).map(h=>String(h.holiday_date).slice(0,10)));const rows=D.range(d.from,d.to).filter(x=>!have.has(x)).map(x=>({holiday_date:x,name:d.name.trim(),type:d.type,applies_to:d.applies_to}));
      if(rows.length)await DB.insert('holidays',rows);await holidaysChanged(d.from,d.to);toast(rows.length+' tanggal libur ditambahkan','ok');Ctx.reset();render(true)}}]});
};


/* =====================================================================
   21. ADMIN: IMPOR FINGERPRINT
   ===================================================================== */
S.im={rows:null,file:'',hasHeader:true,fmt:'dmy',map:{id:-1,dt:-1,date:-1,time:-1},byNo:true};
const normFp=x=>String(x==null?'':x).trim().replace(/^0+(?=\d)/,'').toLowerCase();
function parseStamp(dv,tv,fmt){
  let s=String(dv==null?'':dv).trim();if(tv!=null&&String(tv).trim()!=='')s+=' '+String(tv).trim();
  const m=s.match(/(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})(?:[ T]+(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?\s*([AaPp][Mm])?)?/);
  if(!m)return null;let a=+m[1],b=+m[2],c=+m[3],y,mo,d;
  if(m[1].length>=4){y=a;mo=b;d=c}else if(fmt==='mdy'){mo=a;d=b;y=c}else{d=a;mo=b;y=c}
  if(y<100)y+=2000;let h=m[4]==null?0:+m[4];const mi=m[5]==null?0:+m[5],se=m[6]==null?0:+m[6];
  if(m[7]){const pm=/p/i.test(m[7]);if(pm&&h<12)h+=12;if(!pm&&h===12)h=0}
  if(mo<1||mo>12||d<1||d>31||h>23||mi>59||se>59)return null;
  if(m[4]==null&&(tv==null||String(tv).trim()===''))return null;
  const date=y+'-'+pad(mo)+'-'+pad(d);return{date,ms:D.epoch(date,pad(h)+':'+pad(mi)+':'+pad(se))};
}
function imGuess(h){
  const f=(re)=>h.findIndex(x=>re.test(x));
  const id=f(/^(id|pin|nik|badge|userid|no\.?\s?id|nomor|nip|enrollno|acno|ac-no|no\.?$)/i),dt=f(/(date\s?time|waktu|datetime|scan|tanggal.*jam|time)/i),dd=f(/^(tanggal|date|tgl)/i),tt=f(/^(jam|time|waktu)/i);
  return{id:id>=0?id:0,dt:dt>=0&&dt!==id?dt:-1,date:dt<0&&dd>=0?dd:-1,time:dt<0&&tt>=0&&tt!==dd?tt:-1};
}
reg('impor',{
  range:()=>[TODAY(),TODAY()],
  async render(){
    const im=S.im,hdr=im.rows?(im.hasHeader?im.rows[0].map((x,i)=>String(x||'Kolom '+(i+1))):im.rows[0].map((_,i)=>'Kolom '+(i+1))):[];
    const batches=await DB.select('import_batches',{order:['id',false],limit:15});
    let up='<div class="sk-info">Impor log absensi dari mesin fingerprint (berkas Excel .xlsx atau CSV hasil ekspor mesin). Karena format ekspor tiap merek berbeda, tentukan sendiri kolom mana yang berisi ID pegawai dan waktu absen. Log yang sudah pernah diimpor otomatis dilewati (tidak ganda).<br><b>Format yang langsung dikenali:</b> kolom <code>No.</code> (ID pegawai), <code>Name</code>, <code>Time</code> (tanggal dan jam, contoh 25/09/2026 7:05), <code>State</code>, <code>Tanggal</code>, <code>Jam</code>. Klik <b>Unduh Template</b> untuk contoh berkas yang siap diisi.</div>'
      +'<div class="sk-row" style="margin-bottom:10px"><input type="file" id="imfile" class="sk-in" accept=".xlsx,.xls,.csv,.txt" style="max-width:380px" data-chg="imFile"><button class="sk-b sec" type="button" data-act="fpTemplate">'+ic('download','sm')+'Unduh Template</button>'+(im.file?'<span class="sk-sub">Berkas: <b>'+esc(im.file)+'</b> ('+(im.rows?im.rows.length:0)+' baris)</span>':'')+'</div>';
    if(im.rows){
      const co=hdr.map((h,i)=>[i,h]),pv=im.rows.slice(im.hasHeader?1:0,(im.hasHeader?1:0)+8);
      const parse=r=>{const d=im.map.dt>=0?r[im.map.dt]:(im.map.date>=0?r[im.map.date]:null),t=im.map.dt<0&&im.map.time>=0?r[im.map.time]:null;return parseStamp(d,t,im.fmt)};
      up+='<form id="imform"><div class="sk-grid sk-g3">'+fld('Kolom ID pegawai',sel('id',co,im.map.id,{attr:'data-chg="imMap"'}))+fld('Kolom tanggal+jam (satu kolom)',sel('dt',co,im.map.dt,{blank:'- terpisah / tidak ada -',attr:'data-chg="imMap"'}))+fld('Format tanggal',sel('fmt',[['dmy','Hari/Bulan/Tahun (31/12/2026)'],['mdy','Bulan/Hari/Tahun (12/31/2026)']],im.fmt,{attr:'data-chg="imMap"'}),'Format tahun-bulan-hari (2026-12-31) dikenali otomatis')
        +(im.map.dt<0?fld('Kolom tanggal',sel('date',co,im.map.date,{blank:'- pilih -',attr:'data-chg="imMap"'}))+fld('Kolom jam',sel('time',co,im.map.time,{blank:'- pilih -',attr:'data-chg="imMap"'})):'')+'</div>'
        +'<div class="sk-row" style="margin-bottom:10px">'+chk('hasHeader','Baris pertama adalah judul kolom',im.hasHeader).replace('<input','<input data-chg="imMap"')+chk('byNo','Cocokkan juga dengan NIP bila ID mesin belum dipetakan',im.byNo).replace('<input','<input data-chg="imMap"')+'</div></form>'
        +card('Pratinjau (8 baris pertama)',tbl(['Baris','ID','Nilai waktu','Terbaca sebagai'],pv.map((r,i)=>{const p=parse(r);return[i+1,esc(r[im.map.id]),esc(im.map.dt>=0?r[im.map.dt]:(r[im.map.date]||'')+' '+(r[im.map.time]||'')),p?bd(fmtDT(p.ms),'g'):bd('Tidak terbaca','r')]})),'<button class="sk-b" data-act="imRun">'+ic('upload','sm')+'Impor '+(im.rows.length-(im.hasHeader?1:0))+' baris</button>');
    }
    return card('Impor Log Fingerprint',up)+card('Riwayat impor',tbl(['Waktu','Berkas','Oleh',{t:'Total',c:'num'},{t:'Masuk',c:'num'},{t:'Ganda',c:'num'},{t:'Gagal baca',c:'num'},{t:'ID tak dikenal',c:'num'},''],batches.map(b=>[esc(fmtDT(b.created_at)),esc(b.file_name||'-'),esc(b.uploaded_by||'-'),b.total_rows,b.inserted_rows,b.duplicate_rows,b.failed_rows,b.unmatched_rows,'<button class="sk-b sec sm" data-act="imDel" data-id="'+b.id+'">'+ic('trash','sm')+'Batalkan</button>']),{empty:'Belum ada impor'}));
  }
});
CHG.imFile=async t=>{const f=t.files[0];if(!f)return;try{toast('Membaca berkas...');const rows=await readTable(f);if(rows.length<1)throw new Error('Berkas kosong');S.im.rows=rows;S.im.file=f.name;S.im.hasHeader=true;S.im.map=imGuess(rows[0].map(x=>String(x||'')));await render()}catch(e){fail(e)}};
CHG.imMap=async t=>{const im=S.im;
  if(t.name==='hasHeader'){im.hasHeader=t.checked;if(im.rows)im.map=imGuess(im.hasHeader?im.rows[0].map(x=>String(x||'')):im.rows[0].map(()=>''))}
  else if(t.name==='byNo')im.byNo=t.checked;else if(t.name==='fmt')im.fmt=t.value;else im.map[t.name]=t.value===''?-1:+t.value;
  if(t.name==='dt'&&im.map.dt>=0){im.map.date=-1;im.map.time=-1}await render()};
ACT.fpTemplate=async()=>{
  const H=['No.','Name','Time','State','Tanggal','Jam'];
  const R=[['0000.00.00.001','CONTOH PEGAWAI 1','25/09/2026 07:00','C/In','25/09/2026','07:00'],['0000.00.00.001','CONTOH PEGAWAI 1','25/09/2026 15:05','C/Out','25/09/2026','15:05'],['0000.00.00.002','CONTOH PEGAWAI 2','25/09/2026 07:12','C/In','25/09/2026','07:12']];
  const P=[['PETUNJUK PENGISIAN TEMPLATE IMPOR FINGERPRINT'],[''],
   ['1. Isi data pada sheet "Template". Hapus 3 baris CONTOH sebelum mengimpor data asli.'],
   ['2. No. : ID pegawai pada mesin fingerprint, yaitu NIP pegawai. Harus sama dengan NIP di aplikasi (atau ID Mesin di menu Pegawai > ID Mesin Fingerprint bila dibuat berbeda).'],
   ['3. Name : nama pegawai (hanya informasi, tidak dipakai untuk pencocokan).'],
   ['4. Time : tanggal dan jam absen dalam satu kolom, format dd/mm/yyyy jam:menit, contoh 25/09/2026 7:05.'],
   ['5. State : C/In atau C/Out (informasi saja; masuk dan pulang ditentukan dari jam absen).'],
   ['6. Tanggal dan Jam : boleh diisi sebagai pelengkap. Bila kolom Time terisi, kolom Time yang dipakai.'],
   ['7. Format sel kolom Time, Tanggal, dan Jam sebaiknya "Teks" agar Excel tidak mengubah tampilan tanggal.'],
   ['8. Simpan sebagai .xlsx atau .csv, lalu unggah pada halaman Impor Fingerprint. Log yang sama tidak akan tercatat dua kali.']];
  try{
    await loadScript(XLSX_URL);
    const ws=XLSX.utils.aoa_to_sheet([H,...R]);ws['!cols']=[{wch:18},{wch:28},{wch:19},{wch:9},{wch:12},{wch:8}];
    const wp=XLSX.utils.aoa_to_sheet(P);wp['!cols']=[{wch:130}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Template');XLSX.utils.book_append_sheet(wb,wp,'Petunjuk');
    XLSX.writeFile(wb,'template-impor-fingerprint.xlsx');toast('Template Excel diunduh','ok');
  }catch(e){
    saveFile('template-impor-fingerprint.csv',new Blob([toCSV(H,R)],{type:'text/csv;charset=utf-8'}));toast('Pustaka Excel tidak tersedia, template diunduh sebagai CSV','warn');
  }
};
ACT.imRun=async t=>{
  const im=S.im;if(!im.rows)return;if(im.map.dt<0&&(im.map.date<0||im.map.time<0)){toast('Pilih kolom tanggal+jam, atau kolom tanggal dan jam terpisah','warn');return}
  await busy(t,async()=>{
    const fps={};(await DB.select('employee_fingerprints')).forEach(f=>fps[normFp(f.fingerprint_id)]=f.employee_id);
    const byNo={};S.ctx.employees.forEach(e=>byNo[normFp(e.employee_no)]=e.id);
    const body=im.rows.slice(im.hasHeader?1:0);let failed=0;const seen=new Set(),logs=[],unmatchedIds=new Set();let fileDup=0,unm=0;
    for(const r of body){
      const raw=r[im.map.id],id=String(raw==null?'':raw).trim();if(!id){failed++;continue}
      const p=im.map.dt>=0?parseStamp(r[im.map.dt],null,im.fmt):parseStamp(r[im.map.date],r[im.map.time],im.fmt);if(!p){failed++;continue}
      const iso=new Date(p.ms).toISOString(),key=id+'|'+iso;if(seen.has(key)){fileDup++;continue}seen.add(key);
      let emp=fps[normFp(id)]||(im.byNo?byNo[normFp(id)]:null)||null;if(!emp){unm++;unmatchedIds.add(id)}
      logs.push({fingerprint_id:id,employee_id:emp,logged_at:iso,source:'fingerprint_import',match_status:emp?'matched':'unmatched'});
    }
    if(!logs.length){toast('Tidak ada baris yang dapat dibaca. Periksa pemetaan kolom dan format tanggal.','bad');return}
    if(!await confirmBox('Impor '+logs.length+' log ('+unm+' dengan ID yang belum dikenal, '+failed+' baris tidak terbaca)? Log ganda akan dilewati otomatis.',{ok:'Impor'}))return;
    const p=progUI('Mengimpor log fingerprint'),nf=v=>v.toLocaleString('id-ID');
    let bid,inserted=0,rec=0;const ids=new Set();let minD=null,maxD=null;
    try{
      p.set(1,'Menyiapkan batch impor','');
      const bt=await DB.insert('import_batches',[{file_name:im.file,uploaded_by:actor(),total_rows:body.length,inserted_rows:0,duplicate_rows:0,failed_rows:failed,unmatched_rows:unm}]);bid=bt[0].id;
      const t1=Date.now();
      const res=await DB.upsertP('raw_logs',logs.map(x=>({...x,import_batch_id:bid})),'fingerprint_id,logged_at,source',true,{cols:'employee_id,logged_at',size:1000,conc:4,
        onProgress:(d,n)=>p.set(3+47*d/n,'Menyimpan log absensi',nf(d)+' dari '+nf(n)+' log'+etaTxt(t1,d,n))});
      inserted=res.length;
      for(const x of res){if(x.employee_id)ids.add(x.employee_id);const d=D.local(Date.parse(x.logged_at)).date;if(!minD||d<minD)minD=d;if(!maxD||d>maxD)maxD=d}
      await DB.update('import_batches',{inserted_rows:inserted,duplicate_rows:logs.length-inserted+fileDup},{id:bid});
      if(ids.size&&minD){let t2=0;
        rec=await reprocess(D.add(minD,-1),maxD,[...ids],(ph,d,n)=>{
          if(ph==='load')p.set(52,'Memuat jadwal dan data kehadiran','Tahap ini lebih lama bila datanya banyak');
          else if(ph==='calc')p.set(62,'Menghitung rekap kehadiran','');
          else{if(!t2)t2=Date.now();p.set(66+32*d/n,'Menyimpan rekap kehadiran',nf(d)+' dari '+nf(n)+' rekap harian'+etaTxt(t2,d,n))}})}
      p.set(100,'Selesai','')
    }catch(e){p.close();throw e}
    p.close();
    await audit('import.fingerprint','import_batches',bid,null,{file:im.file,inserted,failed,unmatched:unm});
    S.im.rows=null;S.im.file='';Ctx.reset();await render(true);
    modal({title:'Hasil Impor',body:'<div class="sk-ok">'+inserted+' log baru berhasil diimpor, '+(logs.length-inserted+fileDup)+' ganda dilewati, rekap '+rec+' hari dihitung ulang.</div>'+(failed?'<div class="sk-warn">'+failed+' baris tidak terbaca (ID kosong atau tanggal tidak dikenali).</div>':'')+(unmatchedIds.size?'<div class="sk-warn"><b>'+unmatchedIds.size+' ID mesin belum dikenal</b> dan disimpan sebagai "tidak cocok": '+esc([...unmatchedIds].slice(0,30).join(', '))+(unmatchedIds.size>30?', ...':'')+'.<br>Isi ID tersebut di Pegawai &amp; Departemen &gt; ID Mesin Fingerprint; log lama otomatis dicocokkan.</div>':''),actions:[{t:'Tutup',c:'sec',fn:()=>{}}]});
  });
};
ACT.imDel=async t=>{
  const id=+t.dataset.id;if(!await confirmBox('Batalkan impor ini? Seluruh log dari berkas ini dihapus dan rekap dihitung ulang.',{danger:true,ok:'Batalkan impor'}))return;
  await busy(t,async()=>{
    const logs=await DB.select('raw_logs',{cols:'id,employee_id,logged_at',eq:{import_batch_id:id}});const ids=new Set(logs.filter(l=>l.employee_id).map(l=>l.employee_id));let minD=null,maxD=null;logs.forEach(l=>{const d=D.local(Date.parse(l.logged_at)).date;if(!minD||d<minD)minD=d;if(!maxD||d>maxD)maxD=d});
    await DB.remove('raw_logs',{import_batch_id:id});await DB.remove('import_batches',{id});await audit('import.revert','import_batches',id,null,{logs:logs.length});
    if(ids.size&&minD)await reprocess(D.add(minD,-1),maxD,[...ids]);toast('Impor dibatalkan','ok');Ctx.reset();await render(true)});
};

/* =====================================================================
   22. ADMIN: SALDO CUTI
   ===================================================================== */
S.sd={type:null,period:null};
const balTypes=()=>(S.ctx.request_types||[]).filter(t=>t.deducts_balance);
function periodChoices(){
  const set=new Set([Engine.periodKey(S.ctx,TODAY())]);(S.ctx.ledger||[]).forEach(l=>set.add(l.period_key));(S.ctx.academic_years||[]).forEach(a=>set.add(a.name));const y=D.parse(TODAY()).y;if(Engine.set(S.ctx,'leave_period_mode','calendar')==='jun25'){const k=+Engine.leaveStart(TODAY()).slice(0,4);[k-1,k,k+1].forEach(x=>set.add(x+'/'+(x+1)))}else[y-1,y,y+1].forEach(x=>set.add(String(x)));return[...set].sort();
}
reg('saldo',{
  range:()=>[TODAY(),TODAY()],
  render(){
    const c=S.ctx,types=balTypes();if(!types.length)return'<div class="sk-empty">Belum ada jenis cuti yang memotong saldo. Atur di Pengaturan &gt; Jenis Izin &amp; Cuti.</div>';
    const sd=S.sd;if(!sd.type||!types.some(t=>t.id===sd.type))sd.type=types[0].id;sd.period=sd.period||Engine.periodKey(c,TODAY());
    const rt=types.find(t=>t.id===sd.type),rows=(efFix('sd',c.employees.filter(e=>e.is_active)),c.employees.filter(e=>e.is_active&&efPass('sd',e))).map(e=>{
      const led=Engine.ledger(c).filter(l=>l.employee_id===e.id&&l.request_type_id===rt.id&&l.period_key===sd.period),tot=sum(led.filter(l=>+l.days>0),l=>+l.days),used=0-sum(led.filter(l=>+l.days<0),l=>+l.days),res=Engine.reserved(c,e.id,rt.id,sd.period);
      return{e,tot,used,res,left:tot-used-res};});
    S._sdRows=rows;
    return'<div class="sk-row sp" style="margin-bottom:14px"><div class="sk-row">'+sel('type',types.map(t=>[t.id,t.name]),sd.type,{attr:'data-chg="sdF" style="min-width:190px"'})+sel('period',periodChoices().map(p=>[p,'Periode '+p]),sd.period,{attr:'data-chg="sdF" style="min-width:160px"'})+efBar('sd',c.employees.filter(e=>e.is_active),{inl:1})+'</div><div class="sk-row"><button class="sk-b sec" data-act="sdExport">'+ic('download','sm')+'Excel</button><button class="sk-b sec" data-act="sdExpire">Hanguskan carry-over kedaluwarsa</button><button class="sk-b sec" data-act="sdCarry">Proses akhir periode</button><button class="sk-b" data-act="sdAccrual">'+ic('plus','sm')+'Beri jatah periode</button></div></div>'
     +card('Saldo '+esc(rt.name)+' - periode '+esc(sd.period),tbl(['Pegawai','Departemen',{t:'Jatah / masuk',c:'num'},{t:'Terpakai',c:'num'},{t:'Menunggu',c:'num'},{t:'Sisa tersedia',c:'num'},''],rows.map(r=>[esc(r.e.full_name),esc(deptName(r.e.department_id)),nf(r.tot),nf(r.used),nf(r.res),'<b style="color:'+(r.left<=0?'#dc2626':'#166534')+'">'+nf(r.left)+'</b>','<div class="sk-row" style="gap:5px"><button class="sk-b sec sm" data-act="sdHist" data-e="'+r.e.id+'">Riwayat</button><button class="sk-b sec sm" data-act="sdAdj" data-e="'+r.e.id+'">Penyesuaian</button></div>']),{maxh:'62vh',ef:'sd'}))
     +'<div class="sk-info">Saldo adalah buku besar (ledger): jatah masuk sebagai entri +, pemakaian saat pengajuan disetujui sebagai entri -, pembatalan mengembalikan saldo. Pengajuan yang masih menunggu ditahan di kolom "Menunggu" sehingga tidak dapat dipakai dua kali.</div>';
  }
});
CHG.sdF=async t=>{S.sd[t.name]=t.name==='type'?+t.value:t.value;await render()};
ACT.sdExport=()=>{const rt=rtOf(S.sd.type);exportRows('saldo-'+rt.code+'-'+S.sd.period.replace(/\W/g,'-'),'Saldo',['Pegawai','NIP','Departemen','Jatah/Masuk','Terpakai','Menunggu','Sisa Tersedia'],S._sdRows.map(r=>[r.e.full_name,r.e.employee_no,deptName(r.e.department_id),r.tot,r.used,r.res,r.left]))};
ACT.sdHist=t=>{
  const e=+t.dataset.e,l=Engine.ledger(S.ctx).filter(x=>x.employee_id===e&&x.request_type_id===S.sd.type&&x.period_key===S.sd.period).sort((a,b)=>String(a.created_at).localeCompare(String(b.created_at)));
  const lab={accrual:'Jatah',usage:'Pemakaian',refund:'Pengembalian',adjustment:'Penyesuaian',carry_over:'Carry-over',expire:'Hangus'};let run=0;
  modal({title:'Riwayat saldo - '+nameOf(e),size:'wide',body:tbl(['Waktu','Jenis',{t:'Hari',c:'num'},{t:'Saldo',c:'num'},'Catatan','Oleh'],l.map(x=>{run+=+x.days;return[esc(fmtDT(x.created_at)),esc(lab[x.entry_type]||x.entry_type),'<span style="color:'+(+x.days<0?'#dc2626':'#166534')+'">'+(+x.days>0?'+':'')+nf(x.days)+'</span>',nf(run),esc(x.note||''),esc(x.created_by||'')]}),{empty:'Belum ada catatan'}),actions:[{t:'Tutup',c:'sec',fn:()=>{}}]});
};
ACT.sdAdj=t=>{
  const e=+t.dataset.e,rt=rtOf(S.sd.type);
  modal({title:'Penyesuaian saldo - '+nameOf(e),body:'<div class="sk-info">'+esc(rt.name)+', periode '+esc(S.sd.period)+'. Isi angka positif untuk menambah, negatif untuk mengurangi.</div><form id="adf">'+fld('Jumlah hari',inp('days','',{type:'number',step:'0.5',req:1}))+fld('Alasan (wajib)',inp('note','',{req:1}))+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const d=formData($('#adf',m.el));if(!+d.days||!d.note.trim()){toast('Jumlah dan alasan wajib diisi','warn');return false}
      await DB.insert('leave_balance_entries',[{employee_id:e,request_type_id:S.sd.type,period_key:S.sd.period,entry_type:'adjustment',days:+d.days,note:d.note.trim(),created_by:actor()}]);await audit('balance.adjust','leave_balance_entries',e,null,d);toast('Penyesuaian tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.sdAccrual=()=>{
  const types=balTypes();
  modal({title:'Beri Jatah Cuti Periode',body:'<form id="acf">'+fld('Jenis cuti',sel('type',types.map(t=>[t.id,t.name+' ('+t.annual_quota_days+' hari/periode)']),S.sd.type))+fld('Periode',inp('period',S.sd.period,{req:1}),'Tahun (mis. 2026) atau nama tahun ajaran (mis. 2026/2027) sesuai pengaturan periode saldo')+chk('prorate','Prorata untuk pegawai yang bergabung di tengah tahun (hanya periode tahun kalender)',true)+chk('svc','Lewati pegawai yang belum memenuhi masa kerja minimal',true)+'<div class="sk-hint">Pegawai yang sudah menerima jatah untuk jenis dan periode ini dilewati, sehingga aman dijalankan ulang.</div></form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Proses',i:'check',fn:async m=>{
      const d=formData($('#acf',m.el)),rt=types.find(x=>x.id===+d.type),have=new Set(Engine.ledger(S.ctx).filter(l=>l.request_type_id===rt.id&&l.period_key===d.period&&l.entry_type==='accrual').map(l=>l.employee_id)),rows=[],isYear=/^\d{4}$/.test(d.period);let skipSvc=0;
      for(const e of S.ctx.employees.filter(x=>x.is_active)){
        if(have.has(e.id))continue;
        if(d.svc&&rt.min_service_months>0){const a=D.parse(TODAY()),b=D.parse(String(e.join_date).slice(0,10));let mo=(a.y-b.y)*12+(a.m-b.m);if(a.d<b.d)mo--;if(mo<rt.min_service_months){skipSvc++;continue}}
        let q=+rt.annual_quota_days;if(d.prorate&&isYear){const j=D.parse(String(e.join_date).slice(0,10));if(j.y===+d.period)q=Math.round(q*(13-j.m)/12*2)/2;else if(j.y>+d.period)continue}
        if(q>0)rows.push({employee_id:e.id,request_type_id:rt.id,period_key:d.period,entry_type:'accrual',days:q,note:'Jatah '+rt.name+' '+d.period,created_by:actor()});
      }
      if(!rows.length){toast('Tidak ada pegawai yang perlu diberi jatah'+(skipSvc?' ('+skipSvc+' belum memenuhi masa kerja)':''),'warn');return false}
      await DB.insert('leave_balance_entries',rows);await audit('balance.accrual','leave_balance_entries',d.period,null,{type:rt.code,count:rows.length});toast('Jatah diberikan kepada '+rows.length+' pegawai','ok');S.sd.type=rt.id;S.sd.period=d.period;Ctx.reset();render(true)}}]});
};
ACT.sdCarry=()=>{
  const types=balTypes(),cur=S.sd.period,nx=/^\d{4}$/.test(cur)?String(+cur+1):((S.ctx.academic_years||[]).filter(a=>a.name>cur).sort((a,b)=>a.name.localeCompare(b.name))[0]||{}).name||'';
  modal({title:'Proses Akhir Periode (Carry-over)',body:'<div class="sk-info">Sisa saldo periode lama dipindahkan ke periode baru sesuai kebijakan jenis cuti (batas maksimal carry-over). Sisa yang tidak dipindahkan dihanguskan. Dijalankan sekali per jenis dan periode.</div><form id="cof">'+fld('Jenis cuti',sel('type',types.map(t=>[t.id,t.name+' - '+(t.carry_over_policy==='capped'?'carry-over maks. '+t.carry_over_max_days+' hari':'tanpa carry-over')]),S.sd.type))+'<div class="sk-grid sk-g2">'+fld('Dari periode',inp('from',cur,{req:1}))+fld('Ke periode',inp('to',nx,{req:1}))+'</div></form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Proses',i:'check',fn:async m=>{
      const d=formData($('#cof',m.el)),rt=types.find(x=>x.id===+d.type);if(!d.from||!d.to||d.from===d.to){toast('Periode asal dan tujuan harus berbeda','warn');return false}
      const done=Engine.ledger(S.ctx).some(l=>l.request_type_id===rt.id&&l.period_key===d.from&&l.entry_type==='expire'&&/akhir periode/i.test(l.note||''));if(done){toast('Proses akhir periode untuk jenis dan periode ini sudah pernah dijalankan','bad');return false}
      const rows=[];let carried=0,expired=0;
      for(const e of S.ctx.employees.filter(x=>x.is_active)){const bal=Engine.balance(S.ctx,e.id,rt.id,d.from);if(bal<=0)continue;
        const cap=rt.carry_over_policy==='capped'?Math.min(bal,+rt.carry_over_max_days):0,gone=bal-cap;
        if(cap>0){let vu=null;if(rt.carry_over_valid_months>0){const p=D.parse(D.mStart(TODAY()));const dt=new Date(Date.UTC(p.y,p.m-1+(+rt.carry_over_valid_months),1));vu=dt.toISOString().slice(0,10)}
          rows.push({employee_id:e.id,request_type_id:rt.id,period_key:d.to,entry_type:'carry_over',days:cap,note:'Carry-over dari '+d.from,created_by:actor(),valid_until:vu});carried+=cap}
        rows.push({employee_id:e.id,request_type_id:rt.id,period_key:d.from,entry_type:'expire',days:-bal,note:'Akhir periode '+d.from+(cap>0?' (dipindahkan '+cap+', hangus '+gone+')':' (hangus '+gone+')'),created_by:actor()});expired+=gone}
      if(!rows.length){toast('Tidak ada sisa saldo untuk diproses','warn');return false}
      await DB.insert('leave_balance_entries',rows);await audit('balance.rollover','leave_balance_entries',d.from+'>'+d.to,null,{type:rt.code,carried,expired});toast('Selesai: '+nf(carried)+' hari dipindahkan, '+nf(expired)+' hari hangus','ok');S.sd.type=rt.id;S.sd.period=d.to;Ctx.reset();render(true)}}]});
};
ACT.sdExpire=async t=>{
  const today=TODAY(),led=Engine.ledger(S.ctx),cos=led.filter(l=>l.entry_type==='carry_over'&&l.valid_until&&String(l.valid_until).slice(0,10)<today&&!led.some(x=>x.entry_type==='expire'&&String(x.note||'').includes('[co#'+l.id+']')));
  if(!cos.length){toast('Tidak ada carry-over yang kedaluwarsa','warn');return}
  if(!await confirmBox(cos.length+' carry-over sudah melewati masa berlaku. Hanguskan sisa yang belum terpakai?',{ok:'Hanguskan'}))return;
  await busy(t,async()=>{const rows=[];for(const l of cos){const bal=Engine.balance(S.ctx,l.employee_id,l.request_type_id,l.period_key),g=Math.min(+l.days,Math.max(0,bal));if(g>0)rows.push({employee_id:l.employee_id,request_type_id:l.request_type_id,period_key:l.period_key,entry_type:'expire',days:-g,note:'Carry-over kedaluwarsa [co#'+l.id+']',created_by:actor()});else rows.push({employee_id:l.employee_id,request_type_id:l.request_type_id,period_key:l.period_key,entry_type:'expire',days:0,note:'Carry-over sudah terpakai [co#'+l.id+']',created_by:actor()})}
    await DB.insert('leave_balance_entries',rows);toast('Carry-over kedaluwarsa dihanguskan','ok');Ctx.reset();await render(true)});
};

/* =====================================================================
   23. LAPORAN (Admin dan Kepala Dept.)
   ===================================================================== */
const REPORTS={
  rekap:{n:'Rekap Kehadiran',d:'Ringkasan kehadiran per pegawai',build(o){const h=['Nama','NIP','Departemen','Hari Kerja','Hadir','Terlambat','Menit Terlambat','Pulang Cepat (mnt)','Alpha','Belum Lengkap','Cuti','Sakit','Izin','Dinas Luar','Jam Kerja','Lembur (jam)','% Hadir'];
    return{h,r:o.emps.map(e=>{const s=summarize(e.id,o.from,o.to);return[e.full_name,e.employee_no,deptName(e.department_id),s.sched,s.hadir,s.late,s.late_min,s.early_min,s.absent,s.incomplete,s.leave,s.sick,s.permit,s.official_duty,Math.round(s.worked/6)/10,Math.round(s.ot/6)/10,s.pct+'%']})}}},
  late:{n:'Keterlambatan',d:'Frekuensi dan total menit terlambat',build(o){const rows=o.emps.map(e=>{let n=0,m=0,mx=0;D.range(o.from,o.to).forEach(ds=>{const r=recFor(e.id,ds);if(r.late_minutes>0){n++;m+=r.late_minutes;mx=Math.max(mx,r.late_minutes)}});return[e.full_name,e.employee_no,deptName(e.department_id),n,m,n?Math.round(m/n):0,mx]}).filter(r=>r[3]>0).sort((a,b)=>b[4]-a[4]);return{h:['Nama','NIP','Departemen','Kali Terlambat','Total Menit','Rata-rata (mnt)','Terlama (mnt)'],r:rows}}},
  absent:{n:'Ketidakhadiran (Alpha)',d:'Daftar tanggal alpha atau absen tidak lengkap',build(o){const rows=[];o.emps.forEach(e=>D.range(o.from,o.to).forEach(ds=>{const r=recFor(e.id,ds);if(r.status==='absent'||r.status==='incomplete')rows.push([e.full_name,e.employee_no,deptName(e.department_id),ds,dayName(ds),(ST[r.status]||[r.status])[0],r.note||''])}));return{h:['Nama','NIP','Departemen','Tanggal','Hari','Status','Catatan'],r:rows.sort((a,b)=>a[3].localeCompare(b[3]))}}},
  leave:{n:'Cuti, Izin & Sakit',d:'Pengajuan yang disetujui pada rentang tanggal',build(o){const ids=new Set(o.emps.map(e=>e.id));const rows=(S.ctx.requests||[]).filter(r=>r.status==='approved'&&ids.has(r.employee_id)&&r.start_date<=o.to&&r.end_date>=o.from&&['leave','permit','sick','official_duty','time_permit'].includes(rtOf(r.request_type_id).category)).map(r=>[nameOf(r.employee_id),deptOf(r.employee_id),rtOf(r.request_type_id).name,String(r.start_date).slice(0,10),String(r.end_date).slice(0,10),+r.total_days,r.reason||'(tidak ditampilkan)']).sort((a,b)=>a[3].localeCompare(b[3]));return{h:['Nama','Departemen','Jenis','Mulai','Selesai','Hari','Alasan'],r:rows}}},
  balance:{n:'Saldo Cuti',d:'Jatah, pemakaian, dan sisa semua jenis cuti berpotong saldo',build(o){const rows=[];o.emps.forEach(e=>Engine.balances(S.ctx,e.id,o.to).forEach(b=>rows.push([e.full_name,e.employee_no,deptName(e.department_id),b.type.name,b.period,b.total,b.used,b.reserved,b.left])));return{h:['Nama','NIP','Departemen','Jenis','Periode','Jatah','Terpakai','Menunggu','Sisa'],r:rows}}},
  ot:{n:'Lembur',d:'Menit lembur yang disetujui dan terhitung',build(o){const rows=o.emps.map(e=>{let n=0,m=0;D.range(o.from,o.to).forEach(ds=>{const r=recFor(e.id,ds);if(r.overtime_minutes>0){n++;m+=r.overtime_minutes}});return[e.full_name,e.employee_no,deptName(e.department_id),n,m,Math.round(m/6)/10]}).filter(r=>r[3]>0).sort((a,b)=>b[4]-a[4]);return{h:['Nama','NIP','Departemen','Hari Lembur','Total Menit','Total Jam'],r:rows}}},
  dept:{n:'Kehadiran per Departemen',d:'Perbandingan antar departemen',build(o){const by=groupBy(o.emps,e=>e.department_id||0);const rows=Object.keys(by).map(k=>{const t={sched:0,hadir:0,late:0,absent:0,izin:0,dinas:0};by[k].forEach(e=>{const s=summarize(e.id,o.from,o.to);t.sched+=s.sched;t.hadir+=s.hadir;t.late+=s.late;t.absent+=s.absent;t.izin+=s.leave+s.sick+s.permit;t.dinas+=s.official_duty});return[+k?deptName(+k):'Tanpa departemen',by[k].length,t.sched,t.hadir,t.late,t.absent,t.izin,t.dinas,(t.sched?Math.round((t.hadir+t.dinas)*100/t.sched):0)+'%']});return{h:['Departemen','Pegawai','Hari Terjadwal','Hadir','Terlambat','Alpha','Cuti/Sakit/Izin','Dinas Luar','% Hadir'],r:rows}}},
  shift:{n:'Jadwal Shift',d:'Matriks shift terbit per pegawai (maks. 31 hari)',build(o){if(D.diff(o.from,o.to)>30)return{h:['Info'],r:[['Laporan jadwal shift dibatasi 31 hari. Perpendek rentang tanggal.']]};const days=D.range(o.from,o.to),ix=Engine.ix(S.ctx);const rows=o.emps.filter(e=>days.some(ds=>ix.sa[e.id+'|'+ds])).map(e=>[e.full_name,deptName(e.department_id),...days.map(ds=>{const a=ix.sa[e.id+'|'+ds];if(!a||a.status!=='published')return'';if(!a.shift_id)return'Libur';const s=ix.shift[a.shift_id];return s?s.name:'?'})]);return{h:['Nama','Departemen',...days.map(d=>D.parse(d).d+' '+dayName(d))],r:rows}}},
  raw:{n:'Log Absensi Mentah',d:'Seluruh catatan absen (fingerprint dan mandiri)',build(o){const ids=new Set(o.emps.map(e=>e.id)),all=(S.ctx.raw_logs||[]).filter(l=>(l.employee_id?ids.has(l.employee_id):S.mode==='admin'&&!o.dept)).sort((a,b)=>Date.parse(a.logged_at)-Date.parse(b.logged_at));const src={fingerprint_import:'Fingerprint',web:'Absen mandiri',mobile:'Absen mandiri',manual:'Manual'};
    return{h:['Waktu','Pegawai','ID Mesin','Sumber'],r:all.slice(0,5000).map(l=>[fmtDT(l.logged_at),l.employee_id?nameOf(l.employee_id):'(ID tidak dikenal)',l.fingerprint_id||'-',src[l.source]||l.source]),note:all.length>5000?'Hanya 5.000 baris pertama ditampilkan. Persempit rentang tanggal.':''}}}
};
S.rp={rep:'rekap',from:null,to:null,dept:'',div:'',type:'',q:'',lim:100};
function initRp(){if(!S.rp.from){const t=TODAY();S.rp.from=payRange(payLabel(t))[0];S.rp.to=t}}
function rpPayOpts(){const o=[];let l=payLabel(TODAY());for(let i=0;i<10;i++){const r=payRange(l);o.push([l,'Payroll '+fmtM(l)+' ('+r[0].slice(8)+' '+MONS[D.parse(r[0]).m-1].slice(0,3)+' - 25 '+MONS[D.parse(r[1]).m-1].slice(0,3)+')']);l=D.mStart(D.add(l,-1))}return o}
CHG.rpPay=async t=>{if(!t.value)return;const r=payRange(t.value);S.rp.from=r[0];S.rp.to=r[1];S.rp.lim=100;await render()};
const RP_ICON={rekap:'list',late:'clock',absent:'alert',leave:'calendar',balance:'brief',ot:'moon',dept:'building',shift:'swap',raw:'file'};
const RP_COL={rekap:'#1d4ed8',late:'#d97706',absent:'#dc2626',leave:'#7c3aed',balance:'#0891b2',ot:'#4f46e5',dept:'#059669',shift:'#0ea5e9',raw:'#64748b'};
const RP_PILL={'Hadir':'g','Terlambat':'a','Alpha':'r','Belum Lengkap':'o','Cuti':'b','Sakit':'v','Izin':'c','Dinas Luar':'i','Dinas':'i','Cuti/Sakit/Izin':'b','Kali Terlambat':'a','Hari Lembur':'i'};
const RP_STEP=100;
const rpSum=(r,i)=>r.reduce((a,x)=>a+(+x[i]||0),0);
const rpU=(r,i)=>new Set(r.map(x=>x[i])).size;
function rpTop(title,ico,items,cls,fmt){
  if(!items.length)return'';const mx=Math.max(1,...items.map(x=>x.v));
  return`<section class="rk-panel"><h4>${ic(ico,'sm')}${esc(title)}</h4>${items.map(x=>`<div class="rk-dr"><div class="rk-drh"><span>${esc(x.n)}${x.s?` <small>${esc(x.s)}</small>`:''}</span><b class="${cls.t||''}">${esc(fmt(x.v))}</b></div><div class="rk-bar"><i class="${cls.b}" style="width:${Math.max(3,Math.round(x.v*100/mx))}%"></i></div></div>`).join('')}</section>`}
function rpKpis(rep,out){
  const r=out.r,n=r.length;if(!n||out.h[0]==='Info')return'';let k=[];
  if(rep==='rekap'){const sc=rpSum(r,3),h=rpSum(r,4),l=rpSum(r,5),a=rpSum(r,8),d=rpSum(r,13);
    k=[kpi('chart','#1d4ed8',sc?Math.round((h+d)*100/sc)+'%':'-','Tingkat kehadiran',n+' pegawai'),kpi('check','#059669',h,'Total hadir',(h-l)+' tepat waktu'),kpi('clock','#d97706',l,'Terlambat',fmtMin(rpSum(r,6))+' total'),kpi('alert','#dc2626',a,'Alpha',rpSum(r,9)+' belum lengkap')]}
  else if(rep==='late'){const c=rpSum(r,3),m=rpSum(r,4);k=[kpi('users','#d97706',n,'Pegawai terlambat'),kpi('clock','#b45309',c,'Total kejadian'),kpi('chart','#1d4ed8',fmtMin(m),'Total waktu terlambat'),kpi('alert','#dc2626',c?Math.round(m/c)+' mnt':'-','Rata-rata per kejadian','Terlama '+Math.max(...r.map(x=>x[6]))+' mnt')]}
  else if(rep==='absent'){const al=r.filter(x=>x[5]==='Alpha').length;k=[kpi('alert','#dc2626',n,'Total kejadian'),kpi('users','#7c3aed',rpU(r,0),'Pegawai terkait'),kpi('x','#dc2626',al,'Alpha'),kpi('clock','#ea580c',n-al,'Belum lengkap')]}
  else if(rep==='leave'){const by={};r.forEach(x=>{by[x[2]]=(by[x[2]]||0)+1});const top=Object.entries(by).sort((a,b)=>b[1]-a[1])[0];k=[kpi('inbox','#7c3aed',n,'Pengajuan disetujui'),kpi('calendar','#1d4ed8',nf(rpSum(r,5)),'Total hari'),kpi('users','#059669',rpU(r,0),'Pegawai'),kpi('brief','#0891b2',top?esc(top[0]):'-','Jenis terbanyak',top?top[1]+' pengajuan':'')]}
  else if(rep==='balance'){k=[kpi('users','#1d4ed8',rpU(r,0),'Pegawai'),kpi('brief','#0891b2',nf(rpSum(r,5)),'Total jatah (hari)'),kpi('check','#059669',nf(rpSum(r,6)),'Terpakai (hari)','Menunggu '+nf(rpSum(r,7))),kpi('calendar','#7c3aed',nf(rpSum(r,8)),'Sisa (hari)')]}
  else if(rep==='ot'){k=[kpi('users','#4f46e5',n,'Pegawai lembur'),kpi('calendar','#1d4ed8',rpSum(r,3),'Total hari lembur'),kpi('clock','#d97706',nf(Math.round(rpSum(r,5)*10)/10),'Total jam lembur'),kpi('chart','#059669',esc(r[0][0]),'Lembur terbanyak',r[0][5]+' jam')]}
  else if(rep==='dept'){const p=r.map(x=>({n:x[0],v:parseInt(x[8],10)||0})).sort((a,b)=>b.v-a.v);k=[kpi('building','#059669',n,'Departemen'),kpi('check','#16a34a',p[0].v+'%','Kehadiran tertinggi',p[0].n),kpi('alert','#d97706',p[p.length-1].v+'%','Kehadiran terendah',p[p.length-1].n),kpi('x','#dc2626',rpSum(r,5),'Total alpha')]}
  else if(rep==='shift'){k=[kpi('users','#0ea5e9',n,'Pegawai berjadwal'),kpi('calendar','#1d4ed8',out.h.length-2,'Hari ditampilkan')]}
  else if(rep==='raw'){const un=r.filter(x=>x[1].charAt(0)==='(').length;k=[kpi('file','#64748b',n,'Catatan absen'),kpi('users','#1d4ed8',rpU(r,1),'Pegawai/ID unik'),kpi('alert',un?'#d97706':'#64748b',un,'ID tidak dikenal')]}
  else k=[kpi('list','#1d4ed8',n,'Baris data')];
  return'<div class="sk-kpis rk-kpis">'+k.join('')+'</div>'}
function rpViz(rep,out){
  const r=out.r;if(!r.length||out.h[0]==='Info')return'';
  if(rep==='rekap'){const l=rpSum(r,5);return rkMix({present:Math.max(0,rpSum(r,4)-l),late:l,official_duty:rpSum(r,13),leave:rpSum(r,10),sick:rpSum(r,11),permit:rpSum(r,12),incomplete:rpSum(r,9),absent:rpSum(r,8)})}
  if(rep==='dept')return rpTop('Perbandingan kehadiran antar departemen','building',r.map(x=>({n:x[0],s:x[1]+' pegawai',v:parseInt(x[8],10)||0})).sort((a,b)=>b.v-a.v),{b:'g'},v=>v+'%');
  if(rep==='late')return rpTop('5 pegawai paling lama terlambat','clock',r.slice(0,5).map(x=>({n:x[0],s:x[3]+'x',v:x[4]})),{b:'a',t:'a'},v=>v+' mnt');
  if(rep==='ot')return rpTop('5 pegawai lembur terbanyak','moon',r.slice(0,5).map(x=>({n:x[0],s:x[3]+' hari',v:x[5]})),{b:'i'},v=>v+' jam');
  if(rep==='absent'){const by={};r.forEach(x=>{by[x[0]]=(by[x[0]]||0)+1});return rpTop('5 pegawai paling sering tidak hadir','alert',Object.entries(by).sort((a,b)=>b[1]-a[1]).slice(0,5).map(x=>({n:x[0],v:x[1]})),{b:'r',t:'r'},v=>v+' kali')}
  return''}
function rpCell(h,v,i,rep){
  if(v==null||v==='')return'';
  if((h==='Nama'||h==='Pegawai')&&typeof v==='string'&&v.charAt(0)!=='(')return`<div class="rk-who"><span class="rk-av">${esc(initials(v))}</span><span class="rk-nm"><b>${esc(v)}</b></span></div>`;
  if(h==='% Hadir'){const p=parseInt(v,10);return isNaN(p)?esc(v):rkPct({sched:1,pct:p})}
  if(h==='Status')return bd(String(v),/alpha/i.test(v)?'r':'o');
  if(RP_PILL[h]&&typeof v==='number')return rkN(v,RP_PILL[h]);
  if(rep==='shift'&&i>=2)return v==='Libur'?bd('Libur','s'):`<span class="rk-n b">${esc(v)}</span>`;
  return esc(v)}
function rpQuick(){const t=TODAY(),s=payRange(payLabel(t)),pv=payRange(D.mStart(D.add(s[0],-1)));return{cur:[s[0],t],prev:pv,d7:[D.add(t,-6),t],d30:[D.add(t,-29),t]}}
function reportView(){
  return{
    range:()=>{initRp();let{from,to}=S.rp;if(to<from)to=from;if(D.diff(from,to)>92)to=D.add(from,92);return[from,to]},
    render(){
      initRp();const rp=S.rp;const reps=S.mode==='user'?Object.keys(REPORTS).filter(k=>k!=='raw'):Object.keys(REPORTS);if(!reps.includes(rp.rep))rp.rep=reps[0];
      const rpPool=S.ctx.employees.filter(e=>e.is_active&&(S.mode==='admin'||((S.me&&S.me.head_dept_ids)||[]).includes(e.department_id)));efValid(rp,rpPool);
      const def=REPORTS[rp.rep],emps=S.ctx.employees.filter(e=>e.is_active&&efDeptOk(e,rp.dept)&&efDivOk(e,rp.div)&&(!rp.type||e.type===rp.type));
      const out=def.build({emps,from:rp.from,to:rp.to,dept:rp.dept||rp.div}),q=(rp.q||'').toLowerCase();
      const fr=q?out.r.filter(r=>r.some(x=>String(x==null?'':x).toLowerCase().includes(q))):out.r;
      S._rep={h:out.h,r:fr};
      const cap=D.diff(rp.from,rp.to)>92?'<div class="sk-warn" style="margin-bottom:14px">Rentang dibatasi 93 hari. Data setelah '+esc(fmtD(D.add(rp.from,92)))+' tidak ikut dihitung.</div>':'';
      const depts=(S.ctx.departments||[]).filter(d=>d.is_active).filter(d=>S.mode==='admin'||S.me.head_dept_ids.includes(d.id));
      const dn=[rp.dept?((efOpts(rpPool,'').depts.find(x=>x[0]===String(rp.dept))||[])[1]||''):'',rp.div&&rp.div!==EF_NONE?rp.div:(rp.div===EF_NONE?'Tanpa divisi':'')].filter(Boolean).join(' · ');
      const head=`<div class="rk-top"><div class="rk-tl"><span class="rk-eye">${ic('chart','sm')}Laporan</span><div class="lp-h">${esc(def.n)}</div><div class="lp-sub">${esc(def.d)}</div><div class="rk-meta"><span class="rk-mc">${ic('calendar','sm')}${esc(fmtD(rp.from))} s.d. ${esc(fmtD(rp.to))}</span><span class="rk-mc">${ic('users','sm')}${emps.length} pegawai</span><span class="rk-mc">${ic('list','sm')}${out.r.length} baris</span>${dn?`<span class="rk-mc">${ic('building','sm')}${esc(dn)}</span>`:''}</div></div><div class="sk-rkact"><button class="sk-b ghost sm rk-pr" data-act="rpExport" title="Unduh Excel">${ic('download','sm')}<span class="lbl">Excel</span></button><button class="sk-b sec sm" data-act="rpPrint" title="Cetak atau simpan PDF">${ic('print','sm')}<span class="lbl">Cetak / PDF</span></button></div></div>`;
      const tiles=`<div class="lp-tiles" role="tablist" aria-label="Jenis laporan">${reps.map(k=>`<button type="button" class="lp-tile${k===rp.rep?' on':''}" data-act="rpSel" data-k="${k}" role="tab" aria-selected="${k===rp.rep}"><span class="lp-ti" style="background:${RP_COL[k]||'#1d4ed8'}">${ic(RP_ICON[k]||'list')}</span><span class="lp-tt"><b>${esc(REPORTS[k].n)}</b><small>${esc(REPORTS[k].d)}</small></span></button>`).join('')}</div>`;
      const qk=rpQuick(),chip=(k,l)=>`<button type="button" class="rk-chip${rp.from===qk[k][0]&&rp.to===qk[k][1]?' on':''}" data-act="rpQuick" data-k="${k}">${l}</button>`;
      const payVal=(rpPayOpts().find(o=>{const r=payRange(o[0]);return r[0]===rp.from&&r[1]===rp.to})||[''])[0];
      const filt=`<section class="lp-fl"><div class="lp-qc"><span>Rentang cepat</span>${chip('cur','Periode payroll ini')}${chip('prev','Periode lalu')}${chip('d7','7 hari terakhir')}${chip('d30','30 hari terakhir')}</div><div class="lp-fg">${fld('Periode payroll (26 - 25)',sel('pay',rpPayOpts(),payVal,{blank:'Rentang kustom',attr:'data-chg="rpPay"'}))}${fld('Dari tanggal',inp('from',rp.from,{type:'date',attr:'data-chg="rpF"'}))}${fld('Sampai tanggal',inp('to',rp.to,{type:'date',attr:'data-chg="rpF"'}))}${fld('Departemen',sel('dept',[['','Semua'],...efOpts(rpPool,'').depts],rp.dept,{attr:'data-chg="rpF"'}))}${fld('Divisi',sel('div',[['','Semua'],...efOpts(rpPool,rp.dept).divs],rp.div,{attr:'data-chg="rpF"'}))}${fld('Tipe pegawai',sel('type',[['','Semua'],['teacher','Guru'],['staff','Karyawan Umum']],rp.type,{attr:'data-chg="rpF"'}))}</div></section>`;
      const shown=fr.slice(0,rp.lim);
      const body=(out.note?'<div class="sk-warn" style="margin-bottom:10px">'+esc(out.note)+'</div>':'')+tbl(out.h.map((h,i)=>({t:esc(h),c:i>0&&fr.length&&!/induk|nip/i.test(h)&&fr.every(r=>typeof r[i]==='number'||/^-?\d+([.,]\d+)?%?$/.test(String(r[i])))?'num':''})),shown.map(r=>r.map((x,i)=>rpCell(out.h[i],x,i,rp.rep))),{empty:q?'Tidak ada baris yang cocok dengan pencarian':'Tidak ada data pada rentang ini',maxh:'64vh',cls:'sk-fit'})
        +(fr.length>shown.length?`<div class="lp-more"><button class="sk-b sec" data-act="rpMore">Tampilkan ${Math.min(RP_STEP,fr.length-shown.length)} lagi</button><span class="sk-sub">${shown.length} dari ${fr.length} baris (Excel dan PDF memuat semua baris)</span></div>`:'');
      const search=`<div class="lp-q">${ic('search','sm')}<input class="sk-in" type="search" placeholder="Cari di tabel" aria-label="Cari di tabel" value="${esc(rp.q)}" data-inp="rpQ"></div>`;
      return head+tiles+filt+cap+rpKpis(rp.rep,out)+rpViz(rp.rep,out)+card('Data laporan <span class="rk-cnt">'+fr.length+'</span>',body,search);
    }
  };
}
reg('laporan',reportView());reg('laporan-u',reportView());
CHG.rpF=async t=>{S.rp[t.name]=t.value;if(S.rp.to<S.rp.from)S.rp.to=S.rp.from;S.rp.lim=100;await render()};
ACT.rpSel=async t=>{S.rp.rep=t.dataset.k;S.rp.q='';S.rp.lim=RP_STEP;await render()};
ACT.rpQuick=async t=>{const r=rpQuick()[t.dataset.k];if(!r)return;S.rp.from=r[0];S.rp.to=r[1];S.rp.lim=RP_STEP;await render()};
ACT.rpMore=async()=>{S.rp.lim+=RP_STEP;await render()};
INP.rpQ=debounce(async t=>{S.rp.q=t.value;S.rp.lim=RP_STEP;const p=t.selectionStart;await render();const i=$('[data-inp=rpQ]');if(i){i.focus();try{i.setSelectionRange(p,p)}catch(e){}}},350);

ACT.rpExport=()=>exportRows('laporan-'+S.rp.rep+'-'+S.rp.from,REPORTS[S.rp.rep].n,S._rep.h,S._rep.r);
ACT.rpPrint=()=>printTable(REPORTS[S.rp.rep].n,fmtD(S.rp.from)+' s.d. '+fmtD(S.rp.to),S._rep.h,S._rep.r);

/* =====================================================================
   24. ADMIN: PENUTUPAN PERIODE
   ===================================================================== */
reg('periode',{
  render(){
    const c=S.ctx,t=TODAY(),cur=D.parse(t),list=[];for(let i=0;i<14;i++){let m=cur.m-i,y=cur.y;while(m<1){m+=12;y--}list.push({y,m})}
    const rows=list.map(p=>{const pr=(c.periods||[]).find(x=>x.year===p.y&&x.month===p.m),locked=pr&&pr.status==='locked',isCur=p.y===cur.y&&p.m===cur.m;
      return['<b>'+fmtMon(p.y,p.m)+'</b>'+(isCur?' '+bd('Berjalan','b'):''),locked?bd('Terkunci','r'):bd('Terbuka','g'),locked?esc((pr.locked_by||'-')+' &middot; '+fmtDT(pr.locked_at)):(pr&&pr.reopened_at?'<span class="sk-sub">Dibuka kembali oleh '+esc(pr.reopened_by||'-')+' &middot; '+esc(fmtDT(pr.reopened_at))+'</span>':'-'),isMgr()?(locked?'<button class="sk-b sec sm" data-act="prReopen" data-y="'+p.y+'" data-m="'+p.m+'">'+ic('lock','sm')+'Buka kembali</button>':'<button class="sk-b sm" data-act="prLock" data-y="'+p.y+'" data-m="'+p.m+'">'+ic('lock','sm')+'Kunci periode</button>'):'']});
    return card('Penutupan Periode Kehadiran',tbl(['Periode','Status','Keterangan',''],rows))
      +'<div class="sk-info"><b>Mengunci periode</b> membekukan rekap kehadiran bulan tersebut (untuk penggajian atau laporan resmi): rekap tidak dapat diubah, proses otomatis tidak menimpa, koreksi absen untuk bulan itu ditolak, dan jadwal shift pada bulan itu tidak dapat diubah. Sebelum dikunci aplikasi menghitung ulang rekap bulan tersebut dan memeriksa pengajuan yang belum diputuskan.'+(isMgr()?'':' Hanya HRD Manager / Super Admin yang dapat mengunci atau membuka periode.')+'</div>';
  }
});
ACT.prLock=async t=>{
  const y=+t.dataset.y,m=+t.dataset.m,a=y+'-'+pad(m)+'-01',b=D.mEnd(a);
  await busy(t,async()=>{
    const pend=(S.ctx.requests||[]).filter(r=>['waiting_l1','waiting_l2'].includes(r.status)&&r.start_date<=b&&r.end_date>=a),unm=(await DB.select('raw_logs',{cols:'id',eq:{match_status:'unmatched'},limit:1})).length;
    const msg='Kunci periode '+fmtMon(y,m)+'?'+(pend.length?'\n\nPERHATIAN: masih ada '+pend.length+' pengajuan yang belum diputuskan pada bulan ini. Setelah dikunci, keputusan baru tidak mengubah rekap bulan ini.':'')+(unm?'\n\nAda log fingerprint dengan ID yang belum dikenal.':'')+'\n\nRekap bulan ini akan dihitung ulang terlebih dahulu.';
    if(!await confirmBox(msg,{ok:'Hitung ulang dan kunci',danger:!!pend.length}))return;
    toast('Menghitung ulang rekap...');await reprocess(a,b);await rpc('adm_lock_period',{p_year:y,p_month:m});toast('Periode '+fmtMon(y,m)+' dikunci','ok');Ctx.reset();await render(true)});
};
ACT.prReopen=async t=>{
  const y=+t.dataset.y,m=+t.dataset.m,r=await promptBox('Buka kembali periode '+fmtMon(y,m),'Alasan pembukaan kembali (wajib, tercatat di log audit)',{ok:'Buka kembali',danger:true});if(r===null)return;
  await busy(t,async()=>{await rpc('adm_reopen_period',{p_year:y,p_month:m,p_reason:r});toast('Periode dibuka kembali','ok');Ctx.reset();await render(true)});
};

/* =====================================================================
   25. ADMIN: PENGATURAN
   ===================================================================== */
S.pn={tab:'umum'};
const PN_NAV=[['umum','Umum & Kebijakan','Lembaga, kehadiran, persetujuan','settings'],['jenis','Jenis Izin & Cuti','Aturan setiap pengajuan','inbox'],['tahun','Tahun Ajaran','Periode akademik','school'],['sistem','Info Sistem','Status dan pemeliharaan','shield']];
const PN_CATC={leave:'#2563eb',permit:'#0891b2',sick:'#7c3aed',official_duty:'#4f46e5',time_permit:'#d97706',correction:'#64748b',overtime:'#059669'};
const pnSec=(ico,col,title,desc,body,cls)=>`<section class="pn-card ${cls||''}"><div class="pn-ch"><span class="pn-ci" style="background:${col}">${ic(ico)}</span><div><b>${esc(title)}</b><small>${esc(desc)}</small></div></div><div class="pn-cb">${body}</div></section>`;
const pnTg=(name,title,desc,on)=>`<label class="pn-tg"><span class="pn-tx"><b>${esc(title)}</b>${desc?`<small>${esc(desc)}</small>`:''}</span><input type="checkbox" name="${name}"${on?' checked':''}><i class="pn-sl" aria-hidden="true"></i></label>`;
const pnG=(inner,c)=>`<div class="pn-g${c?' g'+c:''}">${inner}</div>`;
function pnType(t,ro){
  const col=PN_CATC[t.category]||'#64748b',chips=[];
  if(t.deducts_balance)chips.push(bd('Potong saldo','b'));if(t.allow_half_day)chips.push(bd('Setengah hari','c'));if(t.counts_as_present)chips.push(bd('Dihitung hadir','g'));if(t.one_time_per_employment)chips.push(bd('Sekali seumur kerja','v'));
  if(t.requires_attachment_after_days!=null)chips.push(bd('Lampiran > '+nf(t.requires_attachment_after_days)+' hr','a'));if(t.carry_over_policy==='capped')chips.push(bd('Carry-over maks. '+nf(t.carry_over_max_days)+' hr','i'));
  return`<article class="pn-rt${t.is_active?'':' off'}" style="--c:${col}"><div class="pn-rth"><div><b>${esc(t.name)}</b><small>kode: ${esc(t.code)}</small></div>${t.is_active?bd('Aktif','g'):bd('Nonaktif','s')}</div><div class="pn-rtm"><div><span>Jatah</span><b>${t.deducts_balance?nf(t.annual_quota_days)+' hari':'Tanpa potong saldo'}</b></div><div><span>Hitungan</span><b>${t.count_mode==='calendar_days'?'Hari kalender':'Hari kerja'}</b></div><div><span>Maks. berturut</span><b>${t.max_consecutive_days!=null?t.max_consecutive_days+' hari':'Bebas'}</b></div><div><span>Min. pengajuan</span><b>${t.min_notice_days?'H-'+t.min_notice_days:'Bebas'}</b></div></div><div class="pn-rtf">${chips.join('')}<span class="sp"></span>${ro?'':`<button class="sk-b sec sm" data-act="rtEdit" data-id="${t.id}">${ic('edit','sm')}Ubah</button>`}</div></article>`}
function pnYear(a,ro){
  const s=String(a.start_date).slice(0,10),e=String(a.end_date).slice(0,10),t=TODAY(),tot=Math.max(1,D.diff(s,e)),el=Math.min(tot,Math.max(0,D.diff(s,t))),pct=Math.round(el*100/tot);
  const st=t<s?'Belum mulai':t>e?'Sudah selesai':D.diff(t,e)+' hari lagi';
  return`<article class="pn-ay${a.is_current?' cur':''}"><div class="pn-ayh"><span class="pn-ayi">${ic('school')}</span><div class="pn-ayt"><b>${esc(a.name)}</b><small>${esc(fmtD(s))} s.d. ${esc(fmtD(e))}</small></div>${a.is_current?bd('Berjalan','g'):''}</div>${a.is_current?`<div class="pn-ayp"><div class="sk-prog"><i style="width:${pct}%"></i></div><small>${pct}% berjalan · ${esc(st)}</small></div>`:`<div class="pn-aym"><span>${Math.round(tot/30.4)} bulan</span><span>${esc(st)}</span></div>`}${ro?'':`<div class="pn-rtf"><button class="sk-b sec sm" data-act="ayEdit" data-id="${a.id}">${ic('edit','sm')}Ubah</button><button class="sk-b sec sm" data-act="ayDel" data-id="${a.id}" aria-label="Hapus ${esc(a.name)}">${ic('trash','sm')}</button></div>`}</article>`}
reg('pengaturan',{
  range:()=>[TODAY(),TODAY()],
  render(){
    const c=S.ctx,st=c.settings,tab=S.pn.tab,ro=!isMgr(),dis=ro?' disabled':'',v=(k,d)=>st[k]==null?d:st[k];
    const rn=ro?'<div class="sk-warn" style="margin-bottom:14px">Anda login sebagai HRD Staff: pengaturan hanya dapat dilihat. Perubahan dilakukan oleh HRD Manager atau Super Admin.</div>':'';
    const hero=`<div class="rk-top"><div class="rk-tl"><span class="rk-eye">${ic('settings','sm')}Pengaturan</span><div class="lp-h">Pengaturan Sistem</div><div class="lp-sub">Atur kebijakan kehadiran, jenis pengajuan, tahun ajaran, dan pemeliharaan sistem.</div><div class="rk-meta"><span class="rk-mc">${ic('building','sm')}${esc(v('school.name','Lembaga'))}</span><span class="rk-mc">${ic('user','sm')}${esc(ROLE[S.admin.role]||S.admin.role)}</span>${ro?`<span class="rk-mc">${ic('eye','sm')}Hanya lihat</span>`:''}</div></div></div>`;
    const nav=`<nav class="pn-nav" aria-label="Kategori pengaturan">${PN_NAV.map(([k,l,d,i])=>`<button type="button" class="pn-ni${k===tab?' on':''}" data-act="pnTab" data-k="${k}"><span class="pn-nii">${ic(i)}</span><span class="pn-nt"><b>${esc(l)}</b><small>${esc(d)}</small></span></button>`).join('')}</nav>`;
    const wrap=body=>hero+rn+`<div class="pn-wrap">${nav}<div class="pn-main">${body}</div></div>`;
    if(tab==='jenis'){
      const types=c.request_types.slice().sort((a,b)=>(b.is_active-a.is_active)||String(a.name).localeCompare(b.name,'id')),keys=Object.keys(CAT).concat(types.map(t=>t.category).filter(k=>!CAT[k])).filter((k,i,a)=>a.indexOf(k)===i&&types.some(t=>t.category===k));
      const kp='<div class="sk-kpis rk-kpis">'+kpi('inbox','#1d4ed8',types.length,'Jenis pengajuan')+kpi('check','#059669',types.filter(t=>t.is_active).length,'Aktif')+kpi('brief','#0891b2',types.filter(t=>t.deducts_balance).length,'Memotong saldo')+kpi('calendar','#7c3aed',keys.length,'Kategori')+'</div>';
      const bar=`<div class="pn-bar"><div><b>Jenis Izin, Cuti & Pengajuan Lain</b><small>Aturan kuota, batas, dan lampiran untuk tiap jenis pengajuan.</small></div>${ro?'':`<button class="sk-b" data-act="rtEdit">${ic('plus','sm')}Tambah Jenis</button>`}</div>`;
      const grp=keys.map(k=>{const m=types.filter(t=>t.category===k);return`<div class="pn-grp" style="--c:${PN_CATC[k]||'#64748b'}"><i></i>${esc(CAT[k]||k)}<span>${m.length} jenis</span></div><div class="pn-rts">${m.map(t=>pnType(t,ro)).join('')}</div>`}).join('');
      return wrap(bar+kp+(grp||'<div class="sk-empty">Belum ada jenis pengajuan</div>'));
    }
    if(tab==='tahun'){
      const ays=(c.academic_years||[]).slice().sort((a,b)=>String(b.start_date).localeCompare(String(a.start_date)));
      const bar=`<div class="pn-bar"><div><b>Tahun Ajaran</b><small>Dipakai bila "Periode saldo cuti" diatur mengikuti tahun ajaran.</small></div>${ro?'':`<button class="sk-b" data-act="ayEdit">${ic('plus','sm')}Tambah Tahun Ajaran</button>`}</div>`;
      return wrap(bar+(ays.length?`<div class="pn-ays">${ays.map(a=>pnYear(a,ro)).join('')}</div>`:'<div class="sk-empty">Belum ada tahun ajaran</div>'));
    }
    if(tab==='sistem'){
      const it=(ico,col,l,val)=>`<div class="pn-it"><span class="pn-ci" style="background:${col}">${ic(ico)}</span><div><small>${esc(l)}</small><b>${val}</b></div></div>`;
      const act=(ico,col,t,d,btn)=>`<div class="pn-act"><span class="pn-ci" style="background:${col}">${ic(ico)}</span><div class="pn-actx"><b>${esc(t)}</b><small>${esc(d)}</small></div>${btn}</div>`;
      return wrap('<div class="pn-its">'+it('shield','#059669','Mode data',IS_PREVIEW?bd('Preview (data contoh)','a'):bd('Terhubung ke Supabase','g'))+it('user','#1d4ed8','Login sebagai',esc(S.admin.full_name)+' <span class="pn-r">'+esc(ROLE[S.admin.role]||'')+'</span>')+it('clock','#d97706','Zona waktu',esc(v('timezone','Asia/Jakarta')))+it('refresh','#7c3aed','Rekap terakhir diproses',st.last_processed_date?esc(fmtD(st.last_processed_date)):'-')+'</div>'
        +pnSec('settings','#1d4ed8','Pemeliharaan','Tindakan cepat untuk menjaga data tetap akurat','<div class="pn-acts">'+act('refresh','#7c3aed','Hitung ulang rekap','Hitung ulang rekap kehadiran 31 hari terakhir bila ada perubahan jadwal atau kebijakan.','<button class="sk-b sec" data-act="sysReproc">'+ic('refresh','sm')+'Proses</button>')+act('download','#0891b2','Pasang aplikasi','Pasang aplikasi di layar utama perangkat agar bisa dibuka seperti aplikasi biasa.','<button class="sk-b sec'+(pwaStandalone()?' sk-hide':'')+'" id="sk-install" data-act="install">'+ic('download','sm')+'Pasang</button>')+'</div>')
        +'<div class="pn-sec2"><span class="pn-ci" style="background:#059669">'+ic('lock')+'</span><div><b>Keamanan data</b><p>Aplikasi ini aman dipakai dengan kunci anon Supabase karena semua tabel dilindungi Row Level Security: pengunjung tanpa login tidak dapat membaca atau menulis apa pun, dan akses pegawai hanya lewat fungsi database yang memeriksa sesi dan peran.</p></div></div>');
    }
    const shs=c.shifts||[],mins=v('shift_min_staff',{})||{};
    const f=(name,val,o)=>inp(name,val,o||{});
    const form='<form data-form="saveSettings" id="stf"><fieldset'+dis+' style="border:0;padding:0;margin:0;min-width:0">'
      +pnSec('building','#1d4ed8','Umum','Identitas lembaga dan pengaturan dasar',pnG(fld('Nama lembaga',inp('school.name',v('school.name',''),{}))+fld('Zona waktu',sel('timezone',Object.keys(TZMAP).map(k=>[k,k+' (UTC+'+TZMAP[k]+')']),v('timezone','Asia/Jakarta')))+fld('Periode saldo cuti',sel('leave_period_mode',[['jun25','26 Juni - 25 Juni (hangus tiap 25 Juni)'],['calendar','Tahun kalender (Jan-Des)'],['academic','Tahun ajaran']],v('leave_period_mode','jun25')))+fld('Lama sesi login pegawai (jam)',inp('session_hours',v('session_hours',8),{type:'number',min:1}))))
      +pnSec('clock','#d97706','Kebijakan Kehadiran','Aturan keterlambatan, koreksi, dan lembur',pnG(fld('Perhitungan menit terlambat',sel('late_calc_from',[['start','Dari jam masuk'],['tolerance_end','Dari akhir toleransi']],v('late_calc_from','start')))+fld('Jeda minimal masuk-pulang (menit)',inp('min_in_out_gap_minutes',v('min_in_out_gap_minutes',30),{type:'number',min:1}),'Dua absen berdekatan dianggap absen yang sama')+fld('Batas koreksi absen (hari kerja ke belakang)',inp('correction_days_limit',v('correction_days_limit',3),{type:'number',min:0}))+fld('Maks. jam lembur per hari',inp('overtime_max_hours_per_day',v('overtime_max_hours_per_day',4),{type:'number',min:1,step:'0.5'}))+fld('Pembulatan lembur ke bawah (menit)',sel('overtime_round_minutes',[1,5,10,15,30].map(x=>[x,x+' menit']),v('overtime_round_minutes',15)))))
      +pnSec('check','#059669','Persetujuan','Alur persetujuan pengajuan pegawai',pnTg('approval.single_day_final_l1','Cuti 1 hari cukup disetujui Kepala Departemen','Pengajuan cuti satu hari tidak perlu menunggu HRD',v('approval.single_day_final_l1',false))+pnTg('allow_cancel_waiting_l2','Pegawai boleh membatalkan pengajuan yang menunggu HRD','Pembatalan sendiri selama belum diputuskan HRD',v('allow_cancel_waiting_l2',true)))
      +pnSec('pin','#7c3aed','Absen Mandiri (web / ponsel)','Absen dari browser atau ponsel pegawai',pnTg('clock.enabled','Aktifkan absen mandiri','Pegawai dapat absen sendiri lewat aplikasi',v('clock.enabled',true))+pnTg('clock.geofence','Batasi pada radius lembaga (geofence)','Absen hanya sah bila berada dalam radius lokasi lembaga',v('clock.geofence',false))
        +'<div class="pn-geof">'+pnG(fld('Latitude lembaga',inp('school.lat',v('school.lat',''),{type:'number',step:'any'}))+fld('Longitude lembaga',inp('school.lng',v('school.lng',''),{type:'number',step:'any'}))+fld('Radius (meter)',inp('school.radius_m',v('school.radius_m',150),{type:'number',min:10}))+'<div class="pn-loc"><button type="button" class="sk-b sec" data-act="useMyLoc">'+ic('pin','sm')+'Pakai lokasi saya sekarang</button></div>',3)+'</div><div class="sk-hint">Geofence memakai lokasi dari browser sehingga dapat dimanipulasi; anggap sebagai pembatas kewajaran, bukan bukti mutlak. Pegawai dengan dinas luar yang disetujui dikecualikan.</div>')
      +pnSec('swap','#0ea5e9','Jadwal Shift','Batas penjadwalan dan kebutuhan petugas',pnG(fld('Jeda minimal antar shift (jam)',inp('shift_min_gap_hours',v('shift_min_gap_hours',8),{type:'number',min:0}))+fld('Maks. hari kerja berturut-turut',inp('shift_max_consecutive_days',v('shift_max_consecutive_days',6),{type:'number',min:1})))+'<div class="pn-sub">Jumlah petugas minimum per shift per hari <span>(kosong = tanpa batas)</span></div>'+(shs.length?pnG(shs.map(s=>fld(s.name,inp('minstaff_'+s.id,mins[s.id]||'',{type:'number',min:0}))).join(''),3):'<div class="sk-sub">Belum ada shift.</div>'))
      +(ro?'':'<div class="pn-save"><span>'+ic('info','sm')+'Perubahan berlaku setelah disimpan</span><button class="sk-b" type="submit">'+ic('check','sm')+'Simpan Pengaturan</button></div>')+'</fieldset></form>';
    return wrap(form);
  }
});
ACT.pnTab=async t=>{S.pn.tab=t.dataset.k;await render()};
ACT.useMyLoc=async()=>{if(!navigator.geolocation){toast('Browser tidak mendukung lokasi','bad');return}
  await new Promise(res=>navigator.geolocation.getCurrentPosition(p=>{$('[name="school.lat"]').value=p.coords.latitude.toFixed(6);$('[name="school.lng"]').value=p.coords.longitude.toFixed(6);toast('Lokasi terisi, jangan lupa Simpan','ok');res()},e=>{toast('Lokasi tidak dapat dibaca: '+e.message,'bad');res()},{enableHighAccuracy:true,timeout:15000}))};
FORM.saveSettings=async(f,d)=>{
  const num=['session_hours','min_in_out_gap_minutes','correction_days_limit','overtime_max_hours_per_day','overtime_round_minutes','school.radius_m','shift_min_gap_hours','shift_max_consecutive_days'],bool=['approval.single_day_final_l1','allow_cancel_waiting_l2','clock.enabled','clock.geofence'];
  const items={};['school.name','timezone','leave_period_mode','late_calc_from'].forEach(k=>items[k]=d[k]);num.forEach(k=>items[k]=d[k]===''?null:+d[k]);bool.forEach(k=>items[k]=!!d[k]);
  items['school.lat']=d['school.lat']===''?null:+d['school.lat'];items['school.lng']=d['school.lng']===''?null:+d['school.lng'];
  const ms={};(S.ctx.shifts||[]).forEach(s=>{const n=+d['minstaff_'+s.id];if(n>0)ms[s.id]=n});items.shift_min_staff=ms;
  if(!String(items['school.name']||'').trim()){toast('Nama lembaga wajib diisi','warn');return}
  if(items['clock.geofence']&&(items['school.lat']==null||items['school.lng']==null)){toast('Isi latitude dan longitude lembaga untuk mengaktifkan geofence','warn');return}
  const up=[],del=[];for(const k in items){if(items[k]==null||items[k]===''||Number.isNaN(items[k])){del.push(k)}else up.push({key:k,value:items[k],description:SETTING_DESC[k]||null})}
  await DB.upsert('settings',up,'key');for(const k of del)await DB.remove('settings',{key:k});
  await audit('settings.save','settings','umum',null,items);toast('Pengaturan tersimpan','ok');Ctx.reset();await render(true);
};
ACT.sysReproc=async t=>{await busy(t,async()=>{const n=await reprocess(D.add(TODAY(),-31),TODAY());toast('Rekap dihitung ulang: '+n+' baris','ok');Ctx.reset()})};
ACT.rtEdit=t=>{
  const r=t.dataset.id?S.ctx.request_types.find(x=>x.id===+t.dataset.id):null,g=(k,d)=>r?r[k]:d;
  modal({title:r?'Ubah Jenis Pengajuan':'Tambah Jenis Pengajuan',size:'wide',body:'<form id="rtf"><div class="sk-grid sk-g2">'+fld('Nama',inp('name',g('name',''),{req:1}))+fld('Kode (unik, tanpa spasi)',inp('code',g('code',''),{req:1,dis:!!r}))+fld('Kategori',sel('category',Object.entries(CAT),g('category','leave'),{dis:!!r}),r?'Kategori tidak dapat diubah setelah dibuat':'')+fld('Cara hitung hari',sel('count_mode',[['working_days','Hari kerja'],['calendar_days','Hari kalender']],g('count_mode','working_days')))
    +fld('Jatah per periode (hari)',inp('annual_quota_days',g('annual_quota_days',0),{type:'number',min:0,step:'0.5'}))+fld('Masa kerja minimal (bulan)',inp('min_service_months',g('min_service_months',0),{type:'number',min:0}))+fld('Maks. hari berturut-turut (kosong = bebas)',inp('max_consecutive_days',g('max_consecutive_days','')==null?'':g('max_consecutive_days',''),{type:'number',min:1}))+fld('Pengajuan minimal H-sekian',inp('min_notice_days',g('min_notice_days',0),{type:'number',min:0}))
    +fld('Lampiran wajib bila lebih dari (hari, kosong = tidak wajib)',inp('requires_attachment_after_days',g('requires_attachment_after_days','')==null?'':g('requires_attachment_after_days',''),{type:'number',min:0,step:'0.5'}))+fld('Carry-over',sel('carry_over_policy',[['none','Tanpa carry-over'],['capped','Dengan batas maksimal']],g('carry_over_policy','none')))+fld('Carry-over maks. (hari)',inp('carry_over_max_days',g('carry_over_max_days',0),{type:'number',min:0,step:'0.5'}))+fld('Carry-over berlaku (bulan, 0 = tak terbatas)',inp('carry_over_valid_months',g('carry_over_valid_months',0),{type:'number',min:0}))+'</div>'
    +'<div class="sk-row">'+chk('deducts_balance','Memotong saldo cuti',g('deducts_balance',false))+chk('allow_half_day','Boleh setengah hari',g('allow_half_day',false))+chk('counts_as_present','Dihitung hadir',g('counts_as_present',false))+chk('one_time_per_employment','Hanya sekali selama masa kerja',g('one_time_per_employment',false))+chk('is_active','Aktif',g('is_active',true))+'</div></form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const d=formData($('#rtf',m.el));if(!d.name.trim()||!String(d.code).trim()){toast('Nama dan kode wajib diisi','warn');return false}
      const nn=x=>x===''||x==null?null:+x;const row={name:d.name.trim(),count_mode:d.count_mode,annual_quota_days:+d.annual_quota_days||0,min_service_months:+d.min_service_months||0,max_consecutive_days:nn(d.max_consecutive_days),min_notice_days:+d.min_notice_days||0,requires_attachment_after_days:nn(d.requires_attachment_after_days),carry_over_policy:d.carry_over_policy,carry_over_max_days:+d.carry_over_max_days||0,carry_over_valid_months:+d.carry_over_valid_months||0,deducts_balance:!!d.deducts_balance,allow_half_day:!!d.allow_half_day,counts_as_present:!!d.counts_as_present,one_time_per_employment:!!d.one_time_per_employment,is_active:!!d.is_active};
      try{if(r)await DB.update('request_types',row,{id:r.id});else{row.code=d.code.trim().toLowerCase().replace(/[^a-z0-9_]/g,'_');row.category=d.category;await DB.insert('request_types',[row])}}catch(e){if(/duplicate|unique/i.test(e.message)){toast('Kode sudah dipakai','bad');return false}throw e}
      await audit('request_type.save','request_types',r?r.id:row.code,r,row);toast('Jenis pengajuan tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.ayEdit=t=>{
  const a=t.dataset.id?S.ctx.academic_years.find(x=>x.id===+t.dataset.id):null,y=D.parse(TODAY()).y;
  modal({title:a?'Ubah Tahun Ajaran':'Tambah Tahun Ajaran',body:'<form id="ayf">'+fld('Nama (mis. 2026/2027)',inp('name',a?a.name:y+'/'+(y+1),{req:1}))+'<div class="sk-grid sk-g2">'+fld('Mulai',inp('start_date',a?String(a.start_date).slice(0,10):y+'-07-01',{type:'date',req:1}))+fld('Selesai',inp('end_date',a?String(a.end_date).slice(0,10):(y+1)+'-06-30',{type:'date',req:1}))+'</div>'+chk('is_current','Tahun ajaran berjalan',a?a.is_current:false)+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Simpan',i:'check',fn:async m=>{const d=formData($('#ayf',m.el));if(d.end_date<=d.start_date){toast('Tanggal selesai harus setelah mulai','warn');return false}
      if(d.is_current)for(const o of S.ctx.academic_years.filter(x=>x.is_current&&(!a||x.id!==a.id)))await DB.update('academic_years',{is_current:false},{id:o.id});
      const row={name:d.name.trim(),start_date:d.start_date,end_date:d.end_date,is_current:!!d.is_current};try{if(a)await DB.update('academic_years',row,{id:a.id});else await DB.insert('academic_years',[row])}catch(e){if(/duplicate|unique/i.test(e.message)){toast('Nama tahun ajaran sudah ada','bad');return false}throw e}
      toast('Tahun ajaran tersimpan','ok');Ctx.reset();render(true)}}]});
};
ACT.ayDel=async t=>{if(!await confirmBox('Hapus tahun ajaran ini?',{danger:true,ok:'Hapus'}))return;await DB.remove('academic_years',{id:+t.dataset.id});toast('Dihapus','ok');Ctx.reset();await render(true)};

/* =====================================================================
   26. ADMIN: LOG AUDIT
   ===================================================================== */
S.au={q:''};
reg('audit',{
  range:()=>[TODAY(),TODAY()],
  async render(){
    const all=await DB.select('audit_logs',{order:['id',false],limit:400}),q=S.au.q.toLowerCase();S._aud=all;
    const list=all.filter(a=>!q||(a.actor||'').toLowerCase().includes(q)||(a.action||'').toLowerCase().includes(q)||(a.entity_type||'').toLowerCase().includes(q));
    return card('Log Audit (400 catatan terakhir)','<input class="sk-in" style="max-width:300px;margin-bottom:12px" placeholder="Cari pelaku / aksi / objek" value="'+esc(S.au.q)+'" data-inp="auQ">'+tbl(['Waktu','Pelaku','Aksi','Objek',''],list.map(a=>[esc(fmtDT(a.occurred_at)),esc(a.actor||'-'),'<code>'+esc(a.action)+'</code>',esc((a.entity_type||'')+(a.entity_id?' #'+a.entity_id:'')),(a.old_value||a.new_value)?'<button class="sk-b sec sm" data-act="auView" data-id="'+a.id+'">Detail</button>':'']),{empty:'Belum ada catatan',maxh:'68vh'}));
  }
});
INP.auQ=debounce(async t=>{S.au.q=t.value;const p=t.selectionStart;await render();const i=$('[data-inp=auQ]');if(i){i.focus();try{i.setSelectionRange(p,p)}catch(e){}}},400);
ACT.auView=t=>{const a=S._aud.find(x=>x.id===+t.dataset.id);modal({title:'Detail audit',size:'wide',body:'<div class="sk-sub">'+esc(fmtDT(a.occurred_at))+' &middot; '+esc(a.actor||'')+' &middot; <code>'+esc(a.action)+'</code></div><h4 style="margin:12px 0 4px">Sebelum</h4><pre style="white-space:pre-wrap;word-break:break-word;background:#f8fafc;padding:10px;border-radius:10px;font-size:calc(12px*var(--fs,1))">'+esc(JSON.stringify(a.old_value,null,2)||'-')+'</pre><h4 style="margin:12px 0 4px">Sesudah</h4><pre style="white-space:pre-wrap;word-break:break-word;background:#f8fafc;padding:10px;border-radius:10px;font-size:calc(12px*var(--fs,1))">'+esc(JSON.stringify(a.new_value,null,2)||'-')+'</pre>',actions:[{t:'Tutup',c:'sec',fn:()=>{}}]})};

/* =====================================================================
   27. autoSetupDatabase: menyiapkan data master bila tabel masih kosong
   ===================================================================== */
async function autoSetupDatabase(){
  if(IS_PREVIEW)return false; // Preview memakai data contoh di memori; tidak menyentuh Supabase
  const cnt=async t=>(await DB.select(t,{limit:1})).length,did=[];
  const Y=D.parse(TODAY()).y;
  const stRows=await DB.select('settings'),have=new Set(stRows.map(s=>s.key));
  const miss=Object.keys(DEFAULT_SETTINGS).filter(k=>!have.has(k)).map(k=>({key:k,value:DEFAULT_SETTINGS[k],description:SETTING_DESC[k]||null}));
  if(miss.length){await DB.upsert('settings',miss,'key',true);did.push('pengaturan')}
  if(!await cnt('request_types')){await DB.insert('request_types',SEED_TYPES.map(t=>({...t})));did.push('jenis izin & cuti')}
  if(!await cnt('departments')){await DB.insert('departments',SEED_DEPTS.map(n=>({name:n,is_active:true})));did.push('departemen')}
  if(!await cnt('work_schedules')){
    const ws=await DB.insert('work_schedules',SEED_SCHEDULES.map(s=>({name:s.name,tolerance_minutes:s.tolerance_minutes,early_clock_in_minutes:120,late_clock_out_minutes:240,is_active:true})));
    const days=[];ws.forEach((w,i)=>{const s=SEED_SCHEDULES[i];for(let n=1;n<=7;n++){const d=s.days[n];days.push({work_schedule_id:w.id,weekday:n,is_workday:!!d,start_time:d?d[0]:null,end_time:d?d[1]:null,break_minutes:d?d[2]:0})}});
    await DB.insert('work_schedule_days',days);did.push('jadwal Guru & Tata Usaha')}
  if(!await cnt('shifts')){await DB.insert('shifts',SEED_SHIFTS.map(s=>({...s})));did.push('master shift');
    if(!await cnt('shift_patterns'))await DB.insert('shift_patterns',[{name:'4 Pagi - 4 Siang - 4 Malam - Libur',cycle:['Pagi','Pagi','Pagi','Pagi',null,'Siang','Siang','Siang','Siang',null,'Malam','Malam','Malam','Malam',null]}])}
  if(!await cnt('academic_years')){const a=TODAY()>=Y+'-07-01';await DB.insert('academic_years',[{name:a?Y+'/'+(Y+1):(Y-1)+'/'+Y,start_date:(a?Y:Y-1)+'-07-01',end_date:(a?Y+1:Y)+'-06-30',is_current:true}]);did.push('tahun ajaran')}
  if(!await cnt('holidays')){const rows=[];[Y-1,Y,Y+1].forEach(y=>FIXED_HOLIDAYS.forEach(([md,n])=>rows.push({holiday_date:y+'-'+md,name:n,type:'national',applies_to:'all'})));await DB.insert('holidays',rows);did.push('libur tetap')}
  if(did.length){await audit('system.autosetup','system','autoSetupDatabase',null,{seeded:did});toast('Data awal disiapkan: '+did.join(', '),'ok');Ctx.reset();await Ctx.month(ymStart(),true)}
  return did.length>0;
}


/* =====================================================================
   28. PENGGUNA: pembantu tampilan
   ===================================================================== */
const STC={present:['#dcfce7','#166534'],late:['#fef3c7','#92400e'],absent:['#fee2e2','#991b1b'],leave:['#dbeafe','#1e40af'],sick:['#ede9fe','#5b21b6'],permit:['#cffafe','#155e75'],official_duty:['#e0e7ff','#3730a3'],holiday:['#e2e8f0','#475569'],incomplete:['#ffedd5','#9a3412'],day_off:['#f1f5f9','#64748b'],pending:['#f8fafc','#94a3b8'],no_schedule:['#f8fafc','#94a3b8']};
const tag=(txt,st)=>{const c=STC[st]||STC.pending;return'<span class="tag" style="background:'+c[0]+';color:'+c[1]+'">'+esc(txt)+'</span>'};
const CAL={r:null};
function calGrid(ym,cell){
  const rg=CAL.r;CAL.r=null;const first=rg?rg[0]:D.mStart(ym),last=rg?rg[1]:D.mEnd(ym),lead=D.iso(first)-1,days=D.range(first,last),t=TODAY();
  let h='<div class="sk-cal">'+DAYS3.map(d=>'<div class="dh">'+d+'</div>').join('');
  for(let i=0;i<lead;i++)h+='<div class="d out"></div>';
  days.forEach(ds=>{h+='<div class="d'+(ds===t?' today':'')+'"><b>'+D.parse(ds).d+(rg&&(D.parse(ds).d===1||ds===first)?' '+MONS[D.parse(ds).m-1].slice(0,3):'')+'</b>'+cell(ds)+'</div>'});
  return h+'</div>';
}
function schedText(emp,ds){
  const s=Engine.sched(S.ctx,emp,ds);let x;
  if(s.kind==='work')x=s.start+'-'+s.end+(s.src==='shift'?' ('+s.name+')':'');else if(s.kind==='holiday')x='Libur: '+s.name;else if(s.kind==='day_off')x='Libur';else x='-';
  const lv=(S.ctx.requests||[]).find(r=>r.employee_id===emp&&r.status==='approved'&&['leave','sick','permit','official_duty'].includes(rtOf(r.request_type_id).category)&&String(r.start_date).slice(0,10)<=ds&&String(r.end_date).slice(0,10)>=ds);
  return x+(lv?' - '+rtOf(lv.request_type_id).name:'');
}
const greet=()=>{const h=+D.local(NOW()).hm.slice(0,2);return h<11?'Selamat pagi':h<15?'Selamat siang':h<18?'Selamat sore':'Selamat malam'};
const getPos=()=>new Promise(res=>{if(!navigator.geolocation)return res(null);navigator.geolocation.getCurrentPosition(p=>res({lat:p.coords.latitude,lng:p.coords.longitude}),()=>res(null),{enableHighAccuracy:true,timeout:12000,maximumAge:0})});

/* =====================================================================
   29. BERANDA & ABSEN
   ===================================================================== */
/* v34: Beranda minimalis - fokus pada tombol Check In / Check Out */
const todayLogs=()=>{const c=S.ctx,me=S.me,t=TODAY(),a=D.epoch(t,'00:00');return(c.raw_logs||[]).filter(l=>l.employee_id===me.id&&Date.parse(l.logged_at)>=a-6*36e5&&Date.parse(l.logged_at)<a+30*36e5).map(l=>Date.parse(l.logged_at)).filter(ms=>D.local(ms).date===t).sort((x,y)=>x-y)};
const bmD=ds=>{const p=D.parse(ds);return dayName(ds)+', '+pad(p.d)+'/'+pad(p.m)};
const bmSt=s=>{const c=STC[s]||STC.pending,x=ST[s]||[s];return'<span class="bm-st" style="background:'+c[0]+';color:'+c[1]+'">'+esc(x[0])+'</span>'};
const bmT=r=>r.check_in_at?tm(r.check_in_at)+' &ndash; '+(r.check_out_at?tm(r.check_out_at):'...'):'&ndash;';
const bmStat=(v,l,cl)=>'<div class="bm-s'+(cl?' '+cl:'')+'"><b>'+v+'</b><span>'+l+'</span></div>';
reg('beranda',{
  range:()=>{const r=payRange(payLabel(TODAY()));return[r[0],D.add(r[1],7)]},
  render(){
    const c=S.ctx,me=S.me,t=TODAY(),from=payRange(payLabel(t))[0],sm=summarize(me.id,from,t);
    const bals=Engine.balances(c,me.id,t),pend=(c.requests||[]).filter(r=>r.employee_id===me.id&&['waiting_l1','waiting_l2'].includes(r.status));
    const dep=(c.departments||[]).find(d=>d.id===me.department_id),sub=[me.position,dep&&dep.name].filter(Boolean).map(esc).join(' &middot; ');
    const ini=String(me.full_name||'?').trim().split(/\s+/).slice(0,2).map(w=>w.charAt(0)).join('').toUpperCase();
    const recent=D.range(from,t).reverse().filter(ds=>{const r=recFor(me.id,ds);return r.check_in_at||!['no_schedule','pending','day_off'].includes(r.status)}).slice(0,5);
    const headWait=isHead()?(c.requests||[]).filter(r=>r.status==='waiting_l1'&&r.employee_id!==me.id&&(me.head_dept_ids||[]).includes((empIx()[r.employee_id]||{}).department_id)).length:0;
    const qa=(k,i,l,n)=>'<a data-act="nav" data-k="'+k+'"><span class="qi">'+ic(i)+'</span><span class="ql">'+l+'</span>'+dot(n||0)+'</a>';
    const rows=recent.map(ds=>{const r=recFor(me.id,ds);return'<div class="bm-r'+(ds===t?' today':'')+'"><span class="d">'+esc(bmD(ds))+'</span><span class="s">'+bmSt(r.status)+(r.late_minutes?'<small class="lt">+'+r.late_minutes+' mnt</small>':'')+'</span><span class="t">'+bmT(r)+'</span></div>'}).join('');
    const balHtml=bals.map(b=>'<div class="bm-bal"><div class="bm-br"><b>'+esc(b.type.name)+'</b><span>sisa <b>'+nf(b.left)+'</b> / '+nf(b.total)+' hari</span></div><div class="bm-pg"><i style="width:'+(b.total?Math.min(100,Math.round(b.left*100/b.total)):0)+'%"></i></div></div>').join('');
    const pendHtml=pend.map(r=>'<div class="bm-bal"><div class="bm-br"><span><b>'+esc(rtOf(r.request_type_id).name)+'</b> <small>'+esc(reqPeriod(r))+'</small></span>'+rstBadge(r.status)+'</div></div>').join('');
    return'<div class="bm-wrap">'
      +'<div class="bm-hd"><div class="bm-av">'+esc(ini)+'</div><div><small>'+greet()+',</small><h2>'+esc(me.full_name)+'</h2>'+(sub?'<span class="sub">'+sub+'</span>':'')+'</div></div>'
      +(S.me.must_change_password?'<div class="sk-warn">Anda masih memakai kata sandi awal. Ganti di menu Profil.</div>':'')
      +'<section class="bm-abs" id="abs-box">'+absHtml()+'</section>'
      +'<div class="bm-cap"><span>Ringkasan periode ini</span><small>'+sm.sched+' hari kerja</small></div>'
      +'<div class="bm-stats">'+bmStat(sm.hadir,'Hadir','ok')+bmStat(sm.late,'Terlambat','warn')+bmStat(sm.absent,'Alpha','bad')+bmStat(sm.leave+sm.sick+sm.permit,'Cuti/Izin')+'</div>'
      +'<div class="bm-qa">'+qa('ajukan','plus','Ajukan')+qa('riwayat','list','Riwayat')+qa('jadwal-saya','calendar','Jadwal')+qa('pengajuan','file','Pengajuan',pend.length)+qa('tukar','swap','Tukar shift')+(isHead()?qa('persetujuan-l1','inbox','Persetujuan',headWait):'')+'</div>'
      +'<section class="bm-card"><div class="bm-ch"><h3>Riwayat terbaru</h3><button class="bm-lk" data-act="nav" data-k="riwayat">Lihat semua</button></div>'+(rows||'<div class="bm-empty">Belum ada riwayat absensi pada periode ini</div>')+'</section>'
      +(bals.length||pend.length?'<section class="bm-card"><div class="bm-ch"><h3>Cuti &amp; pengajuan</h3><button class="bm-lk" data-act="nav" data-k="pengajuan">Semua</button></div>'+balHtml+pendHtml+'</section>':'')
      +'</div>';
  },
  after(){const e=$('#bk-clock');if(!e)return;S.timers.push(setInterval(()=>{const x=$('#bk-clock');if(x)x.textContent=D.local(NOW()).hms},1000));if(geoOn()){refreshGeo(false);S.timers.push(setInterval(()=>{if(S.route==='beranda')refreshGeo(false)},60000))}}
});
/* ---------- absen mandiri: lokasi (geofence) dan dinas luar ---------- */
const geoOn=()=>{const x=S.settings;return x['clock.enabled']!==false&&!!x['clock.geofence']&&x['school.lat']!=null&&x['school.lat']!==''&&x['school.lng']!=null&&x['school.lng']!==''};
const geoRad=()=>+S.settings['school.radius_m']||150;
const geoDist=(a,b,c,d)=>{const R=6371000,r=x=>x*Math.PI/180;return 2*R*Math.asin(Math.sqrt(Math.sin(r(c-a)/2)**2+Math.cos(r(a))*Math.cos(r(c))*Math.sin(r(d-b)/2)**2))};
const getPosA=()=>new Promise(res=>{if(!navigator.geolocation)return res({err:'unsupported'});navigator.geolocation.getCurrentPosition(p=>res({lat:p.coords.latitude,lng:p.coords.longitude,acc:p.coords.accuracy}),e=>res({err:e.code===1?'denied':e.code===3?'timeout':'unavailable'}),{enableHighAccuracy:true,timeout:15000,maximumAge:0})});
function dutyToday(){const t=TODAY();return((S.ctx&&S.ctx.requests)||[]).filter(r=>r.employee_id===S.me.id&&rtOf(r.request_type_id).category==='official_duty'&&String(r.start_date).slice(0,10)<=t&&String(r.end_date).slice(0,10)>=t)}
const dutyApproved=()=>dutyToday().some(r=>r.status==='approved');
function geoRead(p){const x=S.settings;return{s:'ok',lat:p.lat,lng:p.lng,acc:p.acc,dist:geoDist(p.lat,p.lng,+x['school.lat'],+x['school.lng'])}}
const GEOERR={denied:'Akses lokasi ditolak. Izinkan lokasi untuk situs ini di pengaturan browser, lalu tekan Perbarui lokasi.',timeout:'Lokasi belum terbaca (waktu habis). Pindah ke area terbuka lalu tekan Perbarui lokasi.',unavailable:'Lokasi tidak tersedia. Pastikan GPS/Lokasi perangkat aktif, lalu tekan Perbarui lokasi.',unsupported:'Perangkat atau browser ini tidak mendukung pembacaan lokasi.'};
function absHtml(){
  const c=S.ctx,me=S.me,t=TODAY(),rec=recFor(me.id,t),sch=Engine.sched(c,me.id,t),x=S.settings,canClock=x['clock.enabled']!==false,gf=geoOn(),rad=geoRad();
  const duty=dutyToday(),dOk=duty.some(r=>r.status==='approved'),dWait=duty.some(r=>['waiting_l1','waiting_l2'].includes(r.status));
  const g=S.geo||{s:'idle'},inT=rec.check_in_at,outT=rec.check_out_at,mode=!inT?'in':!outT?'out':'done';
  const inside=g.s==='ok'&&g.dist<=rad,can=canClock&&(!gf||dOk||inside);
  const sTxt=sch.kind==='work'?'Jadwal <b>'+sch.start+' &ndash; '+sch.end+'</b>'+(sch.src==='shift'?' &middot; Shift '+esc(sch.name):''):sch.kind==='holiday'?'Libur: <b>'+esc(sch.name)+'</b>':sch.kind==='day_off'?'Libur sesuai jadwal':'Belum ada jadwal kerja hari ini';
  let hint='',geo='';
  const rf='<button class="bm-rf" data-act="geoRefresh" title="Perbarui lokasi" aria-label="Perbarui lokasi">'+ic('refresh','sm')+'</button>';
  const map='<a href="https://www.google.com/maps?q='+(+x['school.lat'])+','+(+x['school.lng'])+'" target="_blank" rel="noopener noreferrer">Lihat peta</a>';
  const ln=(cl,tt,sub,ac)=>'<div class="bm-geo '+cl+'"><i></i><div><b>'+tt+'</b>'+(sub?'<span>'+sub+'</span>':'')+'</div>'+(ac||'')+'</div>';
  if(!canClock)hint='Absen mandiri tidak diaktifkan lembaga. Gunakan mesin fingerprint.';
  else if(gf){
    if(dOk)geo=ln('ok','Dinas luar disetujui','Anda dapat absen dari lokasi mana pun.');
    else if(g.s==='ok'){
      const ac=g.acc?' &middot; GPS &plusmn;'+Math.round(g.acc)+' m':'';
      if(inside)geo=ln('ok','Di area lembaga',Math.round(g.dist)+' m dari titik absen'+ac+' &middot; '+map,rf);
      else{hint='Tombol aktif setelah Anda berada di area lembaga.';
        geo=ln('bad','Di luar area lembaga',Math.round(g.dist)+' m (batas '+rad+' m)'+ac+'. '+(g.acc&&g.dist-g.acc<=rad?'Akurasi GPS rendah, coba di area terbuka. ':'')+map,rf)
          +(mode==='done'?'':(dWait?'<div class="bm-note">Pengajuan dinas luar hari ini menunggu persetujuan. Setelah disetujui, Anda dapat absen dari mana pun.</div>':'<button class="bm-duty" data-act="dutyClock" data-mode="'+mode+'">'+ic('pin','sm')+'Sedang dinas luar? Ajukan absen dinas luar</button>'));}
    }else if(g.s==='err'){hint='Lokasi diperlukan untuk absen.';geo=ln('err','Lokasi belum terbaca',GEOERR[g.e]||GEOERR.unavailable,rf)}
    else{hint='Tombol aktif setelah lokasi terbaca.';geo=ln('','Memeriksa lokasi...','Izinkan akses lokasi bila diminta.')}
  }
  /* tombol Check In / Check Out: tiap tombol sekaligus menampilkan jam tercatat */
  const inDone=!!inT,outDone=!!outT;
  const bIn='<button class="bm-act in '+(inDone?'done':can?'on':'lock')+'" data-act="clockNow" data-mode="in"'+(inDone||!can?' disabled':'')+'><span class="ico">'+ic(inDone?'check':can?'finger':'clock')+'</span><b>Check In</b><em>'+(inDone?tm(inT):can?'Ketuk untuk absen masuk':'--:--')+'</em></button>';
  const outOn=inDone&&!outDone&&can;
  const bOut='<button class="bm-act out '+(outDone?'done':outOn?'on':'lock')+'" data-act="clockNow" data-mode="out"'+(!inDone||!can?' disabled':'')+'><span class="ico">'+ic(outDone?'check':outOn?'finger':'logout')+'</span><b>Check Out</b><em>'+(outDone?tm(outT)+'<small>Ketuk untuk catat ulang</small>':outOn?'Ketuk untuk absen pulang':'--:--')+'</em></button>';
  const logs=todayLogs(),extra=(rec.note?'<div class="bm-note">'+esc(rec.note)+'</div>':'')+(logs.length>2?'<div class="bm-note">Catatan absen hari ini: <b>'+logs.map(fmtT).join(' &middot; ')+'</b></div>':'');
  return'<div class="bm-top"><div><div class="bm-clock" id="bk-clock">'+D.local(NOW()).hms+'</div><div class="bm-date">'+esc(fmtDL(t))+' &middot; '+esc(x.timezone||'Asia/Jakarta')+'</div></div>'
    +'<div class="bm-tr">'+(rec.late_minutes?bd('Terlambat '+rec.late_minutes+' mnt','a'):stBadge(rec.status))+'</div></div>'
    +'<div class="bm-sch">'+sTxt+'</div>'
    +'<div class="bm-btns">'+bIn+bOut+'</div>'
    +(hint?'<div class="bm-hint">'+hint+'</div>':'')+geo+extra;
}
const paintAbs=()=>{const b=$('#abs-box');if(b)b.innerHTML=absHtml()};
async function refreshGeo(manual){
  if(!geoOn())return;if(manual||!S.geo){S.geo={s:'load'};paintAbs()}
  const p=await getPosA();if(S.route!=='beranda')return;
  S.geo=p.err?{s:'err',e:p.err}:geoRead(p);paintAbs();
}
ACT.geoRefresh=()=>refreshGeo(true);
ACT.clockNow=async t=>{
  const mode=t.dataset.mode||'in';
  await busy(t,async()=>{
    let lat=null,lng=null;
    if(S.settings['clock.geofence']){
      toast('Membaca lokasi...');const p=await getPosA();
      if(!p.err){lat=p.lat;lng=p.lng;
        if(geoOn()){S.geo=geoRead(p);paintAbs();if(!dutyApproved()&&S.geo.dist>geoRad()){toast('Anda berada di luar area lembaga ('+Math.round(S.geo.dist)+' m, batas '+geoRad()+' m). Absen tidak dikirim.','bad');return}}
      }else if(geoOn()&&!dutyApproved()){S.geo={s:'err',e:p.err};paintAbs();toast(GEOERR[p.err]||'Lokasi tidak terbaca','bad');return}
    }
    const r=await rpc('u_clock',{p_token:S.token,p_lat:lat,p_lng:lng});if(!r.ok){toast(r.message,'bad');return}
    toast((mode==='out'?'Check Out':'Check In')+' tercatat pukul '+fmtT(Date.parse(r.logged_at)),'ok');Ctx.reset();await render(true)});
};
ACT.dutyClock=async t=>{
  const types=(S.ctx.request_types||[]).filter(x=>x.category==='official_duty'&&x.is_active!==false);
  if(!types.length){toast('Jenis pengajuan Dinas Luar belum tersedia. Hubungi HRD.','bad');return}
  const mode=t.dataset.mode==='out'?'out':'in',rt=types[0],day=TODAY();
  toast('Membaca lokasi...');const p=await getPosA(),dist=!p.err&&geoOn()?geoDist(p.lat,p.lng,+S.settings['school.lat'],+S.settings['school.lng']):null;
  const locTxt=p.err?'Lokasi tidak terbaca ('+(GEOERR[p.err]||'').split('.')[0]+'). Pengajuan tetap dapat dikirim, namun atasan akan menilai dari bukti yang Anda lampirkan.':'Lokasi terbaca: '+p.lat.toFixed(5)+', '+p.lng.toFixed(5)+' (akurasi &plusmn;'+Math.round(p.acc)+' m)'+(dist!=null?', '+Math.round(dist)+' m dari lembaga':'')+'. Koordinat dan jam dikirim bersama pengajuan.';
  modal({title:'Absen Dinas Luar',size:'',body:'<form id="dcf"><div class="sk-info">Absen di luar area lembaga memerlukan <b>bukti penugasan</b> dan persetujuan atasan (Kepala Departemen, lalu HRD). Setelah disetujui, hari ini tercatat sebagai Dinas Luar.</div>'
    +'<div class="sk-sub" style="margin-bottom:12px">'+locTxt+'</div>'
    +fld('Jenis absen',sel('mode',[['in','Check In'],['out','Check Out']],mode))
    +fld('Tujuan / keterangan dinas',' <textarea class="sk-ta" name="reason" rows="3" required placeholder="mis. Rapat koordinasi di Dinas Pendidikan, pukul 09.00"></textarea>')
    +fld('Tautan bukti (wajib)',inp('url','',{type:'url',req:1,ph:'https://drive.google.com/...'}),'Surat tugas, undangan, atau foto kegiatan. Unggah berkas ke Google Drive/penyimpanan lembaga, atur agar dapat dilihat atasan, lalu tempel tautannya.')+'</form>',
    actions:[{t:'Batal',c:'sec',fn:()=>{}},{t:'Kirim untuk Persetujuan',i:'check',fn:async m=>{
      const d=formData($('#dcf',m.el)),reason=String(d.reason||'').trim(),url=String(d.url||'').trim();
      if(!reason){toast('Isi tujuan atau keterangan dinas','warn');return false}
      if(!/^https?:\/\//i.test(url)){toast('Tautan bukti wajib diisi dengan alamat yang diawali https://','warn');return false}
      const pre='[Absen dinas luar - '+(d.mode==='out'?'Check Out':'Check In')+' pukul '+fmtT(NOW())+(p.err?', lokasi tidak terbaca':', lokasi '+p.lat.toFixed(6)+','+p.lng.toFixed(6)+' (±'+Math.round(p.acc)+' m)'+(dist!=null?', '+Math.round(dist)+' m dari lembaga':''))+'] ';
      const r=await rpc('u_submit_request',{p_token:S.token,p_payload:{type_id:rt.id,start_date:day,end_date:day,day_part:'full',start_time:null,end_time:null,reason:pre+reason,attachment_url:url}});
      if(!r.ok){toast((r.errors||['Pengajuan ditolak sistem']).join(' | '),'bad');return false}
      toast('Absen dinas luar terkirim dan menunggu persetujuan','ok');Ctx.reset();await render(true)}}]});
};

/* =====================================================================
   30. RIWAYAT KEHADIRAN & JADWAL SAYA
   ===================================================================== */
S.rw={tab:'list',off:false};
reg('riwayat',{
  range:()=>payRange(pmLabel()),
  render(){
    const me=S.me,[y,to]=payRange(pmLabel()),t=TODAY(),lim=to>t?t:to,sm=summarize(me.id,y,lim<y?y:lim);
    const days=D.range(y,to),OFF=['holiday','day_off','no_schedule','pending'],list=S.rw.tab!=='cal';
    let body;
    if(!list)body='<section class="bm-card cal">'+(CAL.r=[y,to],calGrid)(y,ds=>{const r=recFor(me.id,ds);if(['no_schedule','pending'].includes(r.status)&&ds>t)return'';return tag((ST[r.status]||[r.status])[0],r.status)+(r.check_in_at?'<small>'+tm(r.check_in_at)+(r.check_out_at?'-'+tm(r.check_out_at):'')+'</small>':'')})+'</section>';
    else{
      const rows=days.filter(ds=>ds<=t||Engine.sched(S.ctx,me.id,ds).kind==='holiday').reverse().map(ds=>{
        const r=recFor(me.id,ds),off=OFF.includes(r.status)&&!r.check_in_at;
        if(off){if(S.rw.off)return'';const s=Engine.sched(S.ctx,me.id,ds);return'<div class="bm-r off"><span class="d">'+esc(bmD(ds))+'</span><span class="s">'+bmSt(r.status)+(s.kind==='holiday'?'<small class="nm">'+esc(s.name)+'</small>':'')+'</span></div>'}
        return'<div class="bm-r'+(ds===t?' today':'')+'"><span class="d">'+esc(bmD(ds))+'</span><span class="s">'+bmSt(r.status)+(r.late_minutes?'<small class="lt">+'+r.late_minutes+' mnt</small>':'')+'</span><span class="t">'+bmT(r)+(r.worked_minutes!=null?'<small>'+fmtMin(r.worked_minutes)+'</small>':'')+'</span>'+(r.note?'<span class="n">'+esc(r.note)+'</span>':'')+'</div>'}).filter(Boolean);
      body='<section class="bm-card">'+(rows.length?rows.join(''):'<div class="bm-empty">Belum ada data pada periode ini</div>')+'</section>';
    }
    return'<div class="bm-wrap w"><div class="bm-bar">'+payBar()+'<div class="bm-seg"><button data-act="rwTab" data-k="list" class="'+(list?'on':'')+'">Daftar</button><button data-act="rwTab" data-k="cal" class="'+(list?'':'on')+'">Kalender</button></div></div>'
      +'<div class="bm-stats">'+bmStat(sm.hadir,'Hadir','ok')+bmStat(sm.late,'Terlambat','warn')+bmStat(sm.absent,'Alpha','bad')+bmStat(sm.leave+sm.sick+sm.permit+sm.official_duty,'Cuti/Izin')+'</div>'
      +'<div class="bm-cap"><span>'+(sm.late?'Total terlambat '+sm.late_min+' mnt':'Tidak ada keterlambatan')+'</span>'+(list?'<button class="bm-lk" data-act="rwOff">'+(S.rw.off?'Tampilkan':'Sembunyikan')+' hari libur</button>':'')+'</div>'
      +body+'</div>';
  }
});
ACT.rwOff=async()=>{S.rw.off=!S.rw.off;await render()};
ACT.rwTab=async t=>{S.rw.tab=t.dataset.k;await render()};
reg('jadwal-saya',{
  render(){
    const me=S.me,y=ymStart(),t=TODAY();
    return'<div class="sk-row sp" style="margin-bottom:12px">'+monthBar()+'</div>'+card('Jadwal '+fmtM(y),calGrid(y,ds=>{
      const s=Engine.sched(S.ctx,me.id,ds),lv=(S.ctx.requests||[]).find(r=>r.employee_id===me.id&&r.status==='approved'&&['leave','sick','permit','official_duty'].includes(rtOf(r.request_type_id).category)&&String(r.start_date).slice(0,10)<=ds&&String(r.end_date).slice(0,10)>=ds);
      let h=s.kind==='work'?'<span class="tag" style="background:'+(s.src==='shift'&&s.color?s.color+'33':'#dbeafe')+';color:#0f172a">'+s.start+'-'+s.end+'</span>'+(s.src==='shift'?'<small>'+esc(s.name)+'</small>':''):s.kind==='holiday'?tag(s.name,'holiday'):s.kind==='day_off'?tag('Libur','day_off'):'';
      if(lv)h+=tag(rtOf(lv.request_type_id).name,rtOf(lv.request_type_id).category==='sick'?'sick':rtOf(lv.request_type_id).category==='leave'?'leave':'permit');return h}))
      +'<div class="sk-info">Jadwal shift hanya tampil setelah diterbitkan HRD. Untuk bertukar shift dengan rekan, gunakan menu Tukar Shift.</div>';
  }
});

/* =====================================================================
   31. AJUKAN (izin / cuti / koreksi / lembur)
   ===================================================================== */
S.aj={v:{}};
const LEAVEISH=['leave','permit','sick','official_duty'];
function ajPayload(v){
  const rt=(S.ctx.request_types||[]).find(t=>t.id===+v.type_id);if(!rt)return null;
  const ranged=LEAVEISH.includes(rt.category);
  return{type_id:+v.type_id,start_date:v.start_date||'',end_date:ranged?(v.end_date||v.start_date||''):(v.start_date||''),day_part:v.day_part||'full',start_time:v.start_time||null,end_time:v.end_time||null,reason:v.reason||'',attachment_url:v.attachment_url||'',correction:rt.category==='correction'?{in:v.c_in||null,out:v.c_out||null}:undefined,overtime:rt.category==='overtime'?{}:undefined};
}
reg('ajukan',{
  range:()=>{const y=D.parse(TODAY()).y;return[y+'-01-01',y+'-12-31']},
  render(){
    const c=S.ctx,types=c.request_types||[],v=S.aj.v;if(!v.type_id||!types.some(t=>t.id===+v.type_id))v.type_id=(types[0]||{}).id;
    const rt=types.find(t=>t.id===+v.type_id)||{},cat=rt.category,t=TODAY();
    const groups={};types.forEach(x=>(groups[CAT[x.category]]=groups[CAT[x.category]]||[]).push(x));
    const typeSel='<select class="sk-sel" name="type_id" data-chg="ajType">'+Object.keys(groups).map(g=>'<optgroup label="'+esc(g)+'">'+groups[g].map(x=>'<option value="'+x.id+'"'+(x.id===+v.type_id?' selected':'')+'>'+esc(x.name)+'</option>').join('')+'</optgroup>').join('')+'</select>';
    let f='';
    if(LEAVEISH.includes(cat)){
      f='<div class="sk-grid sk-g2">'+fld('Tanggal mulai',inp('start_date',v.start_date||'',{type:'date',req:1,attr:'data-chg="ajF"'}))+fld('Tanggal selesai',inp('end_date',v.end_date||'',{type:'date',attr:'data-chg="ajF"'}),'Kosongkan bila hanya satu hari')+'</div>'
        +(rt.allow_half_day?fld('Bagian hari (bila satu hari)',sel('day_part',[['full','Sehari penuh'],['morning','Setengah hari pagi'],['afternoon','Setengah hari siang']],v.day_part||'full',{attr:'data-chg="ajF"'})):'');
    }else if(cat==='time_permit'){
      f='<div class="sk-grid sk-g2">'+fld('Tanggal',inp('start_date',v.start_date||t,{type:'date',req:1,attr:'data-chg="ajF"'}))+fld(/pulang/i.test(rt.code)?'Jam pulang yang diminta':'Perkiraan jam tiba',inp('start_time',v.start_time||'',{type:'time',req:1,attr:'data-chg="ajF"'}))+'</div>';
    }else if(cat==='correction'){
      f='<div class="sk-grid sk-g3">'+fld('Tanggal yang dikoreksi',inp('start_date',v.start_date||t,{type:'date',req:1,attr:'data-chg="ajF"',max:t}))+fld('Jam masuk yang benar',inp('c_in',v.c_in||'',{type:'time',attr:'data-chg="ajF"'}))+fld('Jam pulang yang benar',inp('c_out',v.c_out||'',{type:'time',attr:'data-chg="ajF"'}))+'</div>';
    }else if(cat==='overtime'){
      f='<div class="sk-grid sk-g3">'+fld('Tanggal lembur',inp('start_date',v.start_date||t,{type:'date',req:1,attr:'data-chg="ajF"'}))+fld('Jam mulai lembur',inp('start_time',v.start_time||'',{type:'time',req:1,attr:'data-chg="ajF"'}))+fld('Jam selesai lembur',inp('end_time',v.end_time||'',{type:'time',req:1,attr:'data-chg="ajF"'}))+'</div>';
    }
    const bl=rt.deducts_balance?Engine.balances(c,S.me.id,v.start_date||t).find(b=>b.type.id===rt.id):null;
    return'<div class="sk-cols"><section class="sk-card"><div class="hd"><h3>Form Pengajuan</h3></div><form data-form="ajSubmit" id="ajf">'+fld('Jenis pengajuan',typeSel)+f
      +fld('Alasan',' <textarea class="sk-ta" name="reason" rows="3" required data-inp="ajReason" placeholder="Jelaskan alasan pengajuan">'+esc(v.reason||'')+'</textarea>')
      +fld('Tautan lampiran (opsional)',inp('attachment_url',v.attachment_url||'',{type:'url',ph:'https://drive.google.com/...',attr:'data-chg="ajF"'}),rt.requires_attachment_after_days!=null?'Wajib bila lebih dari '+rt.requires_attachment_after_days+' hari (mis. surat dokter). Unggah berkas ke Google Drive/penyimpanan lembaga lalu tempel tautannya.':'Tempel tautan berkas (Google Drive dsb.) bila ada surat pendukung')
      +'<div id="aj-prev"></div><div class="sk-row end"><button class="sk-b" type="submit">'+ic('check','sm')+'Kirim Pengajuan</button></div></form></section>'
      +'<div>'+(bl?card('Saldo '+esc(bl.type.name),'<div class="sk-row sp"><span>Sisa tersedia</span><b style="font-size:calc(22px*var(--fs,1))">'+nf(bl.left)+' hari</b></div><div class="sk-sub">Jatah '+nf(bl.total)+' &middot; terpakai '+nf(bl.used)+' &middot; menunggu '+nf(bl.reserved)+' (periode '+esc(bl.period)+')</div>'):'')
      +card('Aturan jenis ini','<ul style="padding-left:18px;line-height:1.7;font-size:calc(13px*var(--fs,1))">'+(rt.deducts_balance?'<li>Memotong saldo cuti</li>':'<li>Tidak memotong saldo cuti</li>')+(rt.min_notice_days?'<li>Diajukan minimal '+rt.min_notice_days+' hari sebelumnya</li>':'')+(rt.max_consecutive_days?'<li>Maksimal '+rt.max_consecutive_days+' hari berturut-turut</li>':'')+(rt.min_service_months?'<li>Masa kerja minimal '+rt.min_service_months+' bulan</li>':'')+(cat==='correction'?'<li>Maksimal '+nf(S.settings.correction_days_limit||3)+' hari kerja ke belakang</li>':'')+(cat==='overtime'?'<li>Maksimal '+nf(S.settings.overtime_max_hours_per_day||4)+' jam per hari</li>':'')+'<li>Alur persetujuan: Kepala Departemen, lalu HRD Manager</li></ul>')+'</div></div>';
  },
  after(){if(S.aj.v.start_date||['time_permit','correction','overtime'].includes((rtOf(+S.aj.v.type_id)||{}).category))ajPreview()}
});
CHG.ajType=async t=>{S.aj.v={...S.aj.v,type_id:+t.value,start_date:'',end_date:'',start_time:'',end_time:'',c_in:'',c_out:'',day_part:'full'};await render()};
CHG.ajF=t=>{const f=$('#ajf');S.aj.v={...S.aj.v,...formData(f)};ajPreview()};
INP.ajReason=t=>{S.aj.v.reason=t.value};
const ajPreview=debounce(async()=>{
  const el=$('#aj-prev');if(!el||!S.token)return;const v={...S.aj.v,...formData($('#ajf'))};const p=ajPayload(v);if(!p||!p.start_date){el.innerHTML='';return}
  try{const r=await rpc('u_preview_request',{p_token:S.token,p_payload:{...p,reason:p.reason||'-'}});const e2=$('#aj-prev');if(!e2)return;
    if(r.ok)e2.innerHTML='<div class="sk-ok">Pengajuan memenuhi syarat.'+(r.total_days?' Jumlah hari: <b>'+nf(r.total_days)+'</b>.':'')+(r.deducts?' Sisa saldo setelah disetujui: <b>'+nf(r.after)+'</b> hari.':'')+'</div>';
    else e2.innerHTML='<div class="sk-err"><b>Belum dapat dikirim:</b><br>'+(r.errors||[]).filter(x=>!/Alasan wajib/.test(x)||v.reason).map(esc).join('<br>')+'</div>'}catch(e){console.warn(e)}
},350);
FORM.ajSubmit=async(f,d)=>{
  S.aj.v={...S.aj.v,...d};const p=ajPayload(S.aj.v);if(!p)return;
  const r=await rpc('u_submit_request',{p_token:S.token,p_payload:p});
  if(!r.ok){$('#aj-prev').innerHTML='<div class="sk-err"><b>Pengajuan ditolak sistem:</b><br>'+(r.errors||[]).map(esc).join('<br>')+'</div>';return}
  S.aj.v={};toast('Pengajuan terkirim'+(r.status==='waiting_l1'?' dan menunggu Kepala Departemen':' dan menunggu HRD'),'ok');Ctx.reset();go('pengajuan');
};

/* =====================================================================
   32. PENGAJUAN SAYA
   ===================================================================== */
S.pj={tab:'all'};
reg('pengajuan',{
  range:()=>{const y=D.parse(TODAY()).y;return[y+'-01-01',y+'-12-31']},
  render(){
    const c=S.ctx,me=S.me,mine=(c.requests||[]).filter(r=>r.employee_id===me.id),tab=S.pj.tab,list=mine.filter(r=>tab==='all'||(tab==='wait'&&['waiting_l1','waiting_l2'].includes(r.status))||(tab==='ok'&&r.status==='approved')||(tab==='no'&&['rejected','cancelled'].includes(r.status)));
    const bals=Engine.balances(c,me.id,TODAY());
    const rows=list.map(r=>{const canCancel=['waiting_l1','waiting_l2'].includes(r.status)||(r.status==='approved'&&!r.cancel_requested&&String(r.end_date).slice(0,10)>=TODAY());
      return[esc(rtOf(r.request_type_id).name),esc(reqPeriod(r)+reqExtra(r)),nf(r.total_days),'<div style="max-width:220px">'+esc(r.reason||'')+'</div>',rstBadge(r.status)+(r.cancel_requested?' '+bd('Minta batal','o'):''),esc(fmtD(r.submitted_at)),'<div class="sk-row" style="gap:5px"><button class="sk-b sec sm" data-act="pjView" data-id="'+r.id+'">Detail</button>'+(canCancel?'<button class="sk-b bad sm" data-act="pjCancel" data-id="'+r.id+'">'+(r.status==='approved'?'Minta batal':'Batalkan')+'</button>':'')+'</div>']});
    return(bals.length?'<div class="sk-kpis">'+bals.map(b=>kpi('brief','#2563eb',nf(b.left)+' hari','Sisa '+b.type.name+' ('+b.period+')')).join('')+'</div>':'')
      +stabs([['all','Semua'],['wait','Menunggu'],['ok','Disetujui'],['no','Ditolak / Batal']],tab,'pjTab')
      +card('Pengajuan saya ('+list.length+')',tbl(['Jenis','Tanggal',{t:'Hari',c:'num'},'Alasan','Status','Diajukan',''],rows,{empty:'Belum ada pengajuan'}),'<button class="sk-b" data-act="nav" data-k="ajukan">'+ic('plus','sm')+'Ajukan Baru</button>');
  }
});
ACT.pjTab=async t=>{S.pj.tab=t.dataset.k;await render()};
ACT.pjView=t=>{const r=S.ctx.requests.find(x=>x.id===+t.dataset.id);modal({title:'Detail Pengajuan',size:'wide',body:requestDetail(r,false),actions:[{t:'Tutup',c:'sec',fn:()=>{}}]})};
ACT.pjCancel=async t=>{
  const r=S.ctx.requests.find(x=>x.id===+t.dataset.id),appr=r.status==='approved';
  const reason=await promptBox(appr?'Minta pembatalan':'Batalkan pengajuan',appr?'Alasan pembatalan (wajib)':'Alasan (opsional)',{required:appr,ok:appr?'Kirim permintaan':'Batalkan',danger:!appr,hint:appr?'Pengajuan yang sudah disetujui hanya dapat dibatalkan setelah disetujui HRD.':''});if(reason===null)return;
  await busy(t,async()=>{const x=await rpc('u_cancel_request',{p_token:S.token,p_id:r.id,p_reason:reason});toast(x.status==='cancel_requested'?'Permintaan pembatalan dikirim ke HRD':'Pengajuan dibatalkan','ok');Ctx.reset();await render(true)});
};

/* =====================================================================
   33. TUKAR SHIFT (pegawai)
   ===================================================================== */
reg('tukar',{
  async render(){
    const me=S.me,o=await rpc('u_swap_options',{p_token:S.token});S._swo=o;S.people={};(o.people||[]).forEach(p=>S.people[p.id]=p);(o.partners||[]).forEach(p=>S.people[p.id]=p);
    const shn=id=>(shiftById(id)||{name:'?'}).name,swaps=(S.ctx.swaps||[]).filter(s=>s.requester_employee_id===me.id||s.partner_employee_id===me.id);
    const inc=swaps.filter(s=>s.partner_employee_id===me.id&&s.status==='waiting_partner');
    const form=o.mine.length&&o.partners.length?'<form data-form="swCreate" id="swf"><div class="sk-grid sk-g2">'+fld('Shift saya yang ingin ditukar',sel('my',o.mine.map(m=>[m.work_date,fmtDL(m.work_date)+' - '+shn(m.shift_id)]),''))+fld('Tukar dengan',sel('partner',o.partners.map(p=>[p.id,p.full_name]),'',{attr:'data-chg="swPartner"'}))+'</div>'+fld('Shift milik rekan yang diambil',sel('pdate',[],'',{attr:'id="sw-pdate"'}))+fld('Alasan',inp('reason','',{ph:'mis. ada keperluan keluarga'}))+'<button class="sk-b" type="submit">'+ic('swap','sm')+'Ajukan Tukar Shift</button></form>':'<div class="sk-empty">'+(o.mine.length?'Belum ada rekan sedepartemen dengan shift terbit.':'Anda tidak memiliki shift terbit yang dapat ditukar dalam 60 hari ke depan.')+'</div>';
    return(inc.length?card('Permintaan dari rekan untuk Anda',tbl(['Dari','Tanggal','Alasan',''],inc.map(s=>[esc(nameOf(s.requester_employee_id)),esc('Anda ambil '+fmtD(s.requester_date)+', rekan ambil '+fmtD(s.partner_date)),esc(s.reason||'-'),'<div class="sk-row" style="gap:5px"><button class="sk-b ok sm" data-act="swResp" data-id="'+s.id+'" data-a="1">Setuju</button><button class="sk-b bad sm" data-act="swResp" data-id="'+s.id+'" data-a="0">Tolak</button></div>'])))
      :'')+card('Ajukan Tukar Shift',form)
      +card('Riwayat tukar shift saya',tbl(['Pengaju','Pengganti','Alasan','Status',''],swaps.map(s=>swapRow(s,'')),{empty:'Belum ada permintaan'}))
      +'<div class="sk-info">Alur: Anda mengajukan → rekan menyetujui → HRD menyetujui → jadwal kedua pegawai otomatis bertukar. Tanggal boleh sama (tukar posisi shift) atau berbeda (tukar hari kerja).</div>';
  },
  after(){const s=$('[name=partner]');if(s)CHG.swPartner(s)}
});
CHG.swPartner=t=>{const p=(S._swo.partners||[]).find(x=>x.id===+t.value),el=$('#sw-pdate');if(!el)return;el.innerHTML=(p?p.shifts:[]).map(x=>'<option value="'+x.work_date+'">'+esc(fmtDL(x.work_date)+' - '+((shiftById(x.shift_id)||{name:'?'}).name))+'</option>').join('')};
FORM.swCreate=async(f,d)=>{const r=await rpc('u_swap_create',{p_token:S.token,p_partner:+d.partner,p_my_date:d.my,p_partner_date:d.pdate,p_reason:d.reason});toast('Permintaan tukar shift dikirim','ok');Ctx.reset();await render(true)};
ACT.swResp=async t=>{await busy(t,async()=>{await rpc('u_swap_respond',{p_token:S.token,p_id:+t.dataset.id,p_accept:t.dataset.a==='1'});toast(t.dataset.a==='1'?'Anda menyetujui. Menunggu persetujuan HRD.':'Permintaan ditolak','ok');Ctx.reset();await render(true)})};

/* =====================================================================
   34. KEPALA DEPARTEMEN: kotak persetujuan & kehadiran departemen
   ===================================================================== */
S.l1={tab:'wait'};
reg('persetujuan-l1',{
  range:()=>{const y=D.parse(TODAY()).y;return[y+'-01-01',y+'-12-31']},
  render(){
    const c=S.ctx,me=S.me,hd=me.head_dept_ids||[],inD=r=>hd.includes((empIx()[r.employee_id]||{}).department_id)&&r.employee_id!==me.id;
    const all=(c.requests||[]).filter(inD),wait=all.filter(r=>r.status==='waiting_l1'),hist=all.filter(r=>r.status!=='waiting_l1');
    const list=S.l1.tab==='wait'?wait:hist;
    const hp=c.employees.filter(e=>hd.includes(e.department_id)&&e.id!==me.id);efFix('l1',hp);
    const rows=list.filter(r=>efPassId('l1',r.employee_id)).map(r=>{const rt=rtOf(r.request_type_id);
      const clash=['leave','permit','sick'].includes(rt.category)?all.filter(o=>o.id!==r.id&&['waiting_l1','waiting_l2','approved'].includes(o.status)&&['leave','permit','sick'].includes(rtOf(o.request_type_id).category)&&o.start_date<=r.end_date&&o.end_date>=r.start_date).length:0;
      return[esc(nameOf(r.employee_id)),esc(rt.name),esc(reqPeriod(r)+reqExtra(r)),nf(r.total_days),'<div style="max-width:240px">'+esc(r.reason||'')+'</div>'+(clash?'<div class="sk-sub" style="color:#b45309">'+clash+' rekan lain juga tidak masuk pada tanggal ini</div>':''),rstBadge(r.status),'<div class="sk-row" style="gap:5px"><button class="sk-b sec sm" data-act="l1View" data-id="'+r.id+'">Detail</button>'+(r.status==='waiting_l1'?'<button class="sk-b ok sm" data-act="l1Decide" data-id="'+r.id+'" data-d="approved">Setujui</button><button class="sk-b bad sm" data-act="l1Decide" data-id="'+r.id+'" data-d="rejected">Tolak</button>':'')+'</div>']});
    return stabs([['wait','Menunggu saya ('+wait.length+')'],['hist','Riwayat']],S.l1.tab,'l1Tab')+card('Pengajuan pegawai departemen '+esc(hd.map(deptName).join(', ')),tbl(['Pegawai','Jenis','Tanggal',{t:'Hari',c:'num'},'Alasan','Status',''],rows,{empty:'Tidak ada pengajuan',ef:'l1'}),'<div class="sk-tool">'+efBar('l1',hp,{inl:1})+'</div>')
      +'<div class="sk-info">Setelah Anda menyetujui, pengajuan diteruskan ke HRD Manager untuk persetujuan akhir. Pengajuan Anda sendiri langsung ke HRD.</div>';
  }
});
ACT.l1Tab=async t=>{S.l1.tab=t.dataset.k;await render()};
ACT.l1View=t=>{const r=S.ctx.requests.find(x=>x.id===+t.dataset.id);modal({title:'Detail Pengajuan',size:'wide',body:requestDetail(r,true),actions:[{t:'Tutup',c:'sec',fn:()=>{}}]})};
ACT.l1Decide=async t=>{
  const id=+t.dataset.id,dec=t.dataset.d,r=S.ctx.requests.find(x=>x.id===id);
  const note=await promptBox(dec==='approved'?'Setujui pengajuan':'Tolak pengajuan',dec==='approved'?'Catatan (opsional)':'Alasan penolakan (wajib)',{required:dec!=='approved',ok:dec==='approved'?'Setujui':'Tolak',danger:dec!=='approved',hint:nameOf(r.employee_id)+' - '+rtOf(r.request_type_id).name+' '+reqPeriod(r)});if(note===null)return;
  await busy(t,async()=>{await rpc('u_decide_l1',{p_token:S.token,p_id:id,p_decision:dec,p_note:note});toast(dec==='approved'?'Disetujui dan diteruskan ke HRD':'Pengajuan ditolak','ok');Ctx.reset();await render(true)});
};
S.dp={tab:'ring'};
reg('dept',{
  range:()=>payRange(pmLabel()),
  render(){
    const c=S.ctx,[y,to]=payRange(pmLabel()),t=TODAY(),lim=to>t?t:to,emps0=c.employees.filter(e=>e.is_active&&(S.me.head_dept_ids||[]).includes(e.department_id)),emps=(efFix('dp',emps0),emps0.filter(e=>efPass('dp',e)));
    const bar=payBar();let body;
    if(S.dp.tab==='today'){const rows=emps.map(e=>{const r=recFor(e.id,t);return[esc(e.full_name),stBadge(r.status),esc(schedText(e.id,t)),tm(r.check_in_at),tm(r.check_out_at),r.late_minutes?r.late_minutes+' mnt':'-']});S._dpRows={h:['Pegawai','Status','Jadwal','Masuk','Pulang','Terlambat'],r:rows.map(r=>r.map(x=>String(x).replace(/<[^>]+>/g,'')))};body=card('Kehadiran hari ini - '+fmtDL(t),tbl(['Pegawai','Status','Jadwal','Masuk','Pulang','Terlambat'],rows,{empty:'Tidak ada pegawai',ef:'dp'}))}
    else{const rows=emps.map(e=>{const s=summarize(e.id,y,lim<y?y:lim);return[e.full_name,s.sched,s.hadir,s.late,s.late_min,s.absent,s.leave,s.sick,s.permit,s.official_duty,s.pct+'%']});S._dpRows={h:['Pegawai','Hari Kerja','Hadir','Terlambat','Menit Terlambat','Alpha','Cuti','Sakit','Izin','Dinas','% Hadir'],r:rows};body=card('Ringkasan '+payTitle(),tbl(['Pegawai',{t:'Hari Kerja',c:'num'},{t:'Hadir',c:'num'},{t:'Terlambat',c:'num'},{t:'Menit Terlambat',c:'num'},{t:'Alpha',c:'num'},{t:'Cuti',c:'num'},{t:'Sakit',c:'num'},{t:'Izin',c:'num'},{t:'Dinas',c:'num'},{t:'% Hadir',c:'num'}],rows.map(r=>[esc(r[0]),...r.slice(1)]),{empty:'Tidak ada pegawai',cls:'sk-fit',ef:'dp'}),'<div class="sk-row"><button class="sk-b sec sm" data-act="dpExport">'+ic('download','sm')+'Excel</button></div>')}
    return'<div class="sk-row sp" style="margin-bottom:12px">'+bar+'</div>'+stabs([['ring','Ringkasan Bulanan'],['today','Hari Ini']],S.dp.tab,'dpTab')+'<div class="ef-top">'+efBar('dp',emps0)+'</div>'+body;
  }
});
ACT.dpTab=async t=>{S.dp.tab=t.dataset.k;await render()};
ACT.dpExport=()=>exportRows('kehadiran-departemen-'+pmLabel().slice(0,7),'Departemen',S._dpRows.h,S._dpRows.r);

/* =====================================================================
   35. PROFIL
   ===================================================================== */
reg('profil',{
  needs:false,
  render(){
    const me=S.me,dep=(S.ctx&&S.ctx.departments||[]).find(d=>d.id===me.department_id);
    return card('Profil Saya','<div class="sk-grid sk-g2"><div><div class="sk-sub">Nama</div><b>'+esc(me.full_name)+'</b></div><div><div class="sk-sub">NIP</div><b>'+esc(me.employee_no)+'</b></div><div><div class="sk-sub">Departemen</div><b>'+esc(dep?dep.name:'-')+'</b></div><div><div class="sk-sub">Jabatan</div><b>'+esc(me.position||'-')+'</b></div><div><div class="sk-sub">Tipe / status</div><b>'+esc(TYPE[me.type]||me.type)+' / '+esc(EMPST[me.employment_status]||'')+'</b></div><div><div class="sk-sub">Bergabung</div><b>'+esc(fmtD(me.join_date))+'</b></div><div style="grid-column:1/-1"><div class="sk-sub">Peran</div>'+me.roles.map(r=>bd(ROLE[r]||r,'b')).join(' ')+((me.head_dept_ids||[]).length?' '+bd('Manager / Kepala Departemen','v'):'')+'</div></div>')
      +card('Kontak','<form data-form="profSave">'+'<div class="sk-grid sk-g2">'+fld('No. HP',inp('phone',me.phone||''))+fld('Email',inp('email',me.email||'',{type:'email'}))+'</div><button class="sk-b" type="submit">'+ic('check','sm')+'Simpan Kontak</button></form>')
      +card('Keamanan & Aplikasi','<div class="sk-row"><button class="sk-b sec" data-act="chPw">'+ic('key','sm')+'Ganti Kata Sandi</button><button class="sk-b sec'+(pwaStandalone()?' sk-hide':'')+'" id="sk-install" data-act="install">'+ic('download','sm')+'Pasang Aplikasi</button><button class="sk-b bad" data-act="logout">'+ic('logout','sm')+'Keluar</button></div>');
  }
});
FORM.profSave=async(f,d)=>{await rpc('u_update_profile',{p_token:S.token,p_phone:d.phone,p_email:d.email});S.me.phone=d.phone;S.me.email=d.email;toast('Kontak tersimpan','ok')};
ACT.chPw=()=>openChangePassword(false);

/* ---------- mulai ---------- */
boot();
})();
