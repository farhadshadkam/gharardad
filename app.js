/* ================= سامانه صدور قراردادهای آموزشی ================= */
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwXjADq0Y-i0QCAIe8v2b-XQgdXO2NDoQ2JNz9Zvp3O4lPIM4mXG1e36CClhNfSPaXTRw/exec';

/* مقادیر پیش‌فرض ۱۴۰۵ (در صورت اتصال، از تب Settings گوگل‌شیت بازنویسی می‌شود) */
const DEFAULT_SETTINGS = {
 'refRate.industry':577500,'refRate.agriculture':577500,'refRate.services':480000,'refRate.culture':375000,
 'hourly.industry':144375,'hourly.agriculture':144375,'hourly.services':120000,'hourly.culture':93750,
 'fee.register':300000,'fee.counseling':300000,'fee.exam':1300000,'contract.year':'1405',
 'acc1.title':'درآمد حاصل از خدمات آموزشی','acc1.owner':'خزانه‌داری کل کشور','acc1.sheba':'IR420100004001018703042310','acc1.shetab':'4201-0000-4001-0187-0304-2310',
 'acc2.title':'درآمد حاصل از حق ثبت‌نام','acc2.owner':'حساب موقت مرکز آموزش فنی و حرفه‌ای','acc2.sheba':'IR330100004001018703042278','acc2.shetab':'3301-0000-4001-0187-0304-2278',
 'acc3.title':'درآمد حاصل از حق مشاوره','acc3.owner':'حساب موقت مرکز آموزش فنی و حرفه‌ای','acc3.sheba':'IR110100004001018703042286','acc3.shetab':'1101-0000-4001-0187-0304-2286',
 'acc4.title':'درآمد حاصل از حق آزمون','acc4.owner':'حساب موقت مرکز آموزش فنی و حرفه‌ای','acc4.sheba':'IR330100004001018703042278','acc4.shetab':'3301-0000-4001-0187-0304-2278',
 'acc5.title':'درآمد حاصل از ارجاع ارائه خدمات آموزشی','acc5.owner':'خزانه‌داری کل کشور','acc5.sheba':'IR2501000400101871033261','acc5.shetab':'2501-0004-0010-1871-0332-61',
};

const KHOSHE = {
 industry:{name:'صنعت',ref:'refRate.industry',hr:'hourly.industry'},
 agriculture:{name:'کشاورزی',ref:'refRate.agriculture',hr:'hourly.agriculture'},
 services:{name:'خدمات',ref:'refRate.services',hr:'hourly.services'},
 culture:{name:'فرهنگ و هنر',ref:'refRate.culture',hr:'hourly.culture'},
};

/* نهاده‌های آموزشی — نام‌ها و وزن‌های پیش‌فرض مطابق فایل اکسل (قابل ویرایش در فرم) */
const ITEMS = [
 {name:'مربی',w:20,owner:2},
 {name:'استاندارد (ایمنی/ماهیت معنوی)',w:10,owner:2},
 {name:'کارشناس متخصص',w:5,owner:2},
 {name:'نیروی پشتیبانی',w:9,owner:1},
 {name:'فضا',w:9,owner:1},
 {name:'تجهیزات',w:9,owner:1},
 {name:'ابزارالات',w:9,owner:1},
 {name:'مواد مصرفی',w:9,owner:1},
 {name:'ارزیابی (سوخت)',w:5,owner:1},
 {name:'ارزشیابی بین دوره',w:5,owner:2},
 {name:'ایمنی و بهداشت',w:5,owner:1},
 {name:'مستندسازی الکترونیکی فرایند آموزش',w:5,owner:2},
];

const state = {mode:'new',step:1,settings:{...DEFAULT_SETTINGS},data:null,contractNo:null,trackingCode:null};

/* ---------- ابزارهای عمومی ---------- */
const $ = id => document.getElementById(id);
const faD = s => String(s).replace(/[0-9]/g, d => '۰۱۲۳۴۵۶۷۸۹'[+d]);
const normD = s => String(s??'').replace(/[۰-۹]/g,c=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(c))
  .replace(/[٠-٩]/g,c=>'٠١٢٣٤٥٦٧٨٩'.indexOf(c)).replace(/[٫,]/g,'.');
const num = v => { const n = parseFloat(normD(v)); return isNaN(n) ? 0 : n; };
const fmt = n => faD(Math.round(n).toLocaleString('en-US')).replace(/,/g,'٬');
const esc = s => String(s??'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');

function jalaliToday(){
  try{ const f=new Intl.DateTimeFormat('fa-IR-u-nu-latn',{calendar:'persian',year:'numeric',month:'2-digit',day:'2-digit'});
    return f.format(new Date()); }catch(e){ return ''; }
}

/* ---------- موتور محاسبات (قفل‌شده مطابق شیوه‌نامه ۱۴۰۵) ---------- */
function effectiveHours(h){
  h = num(h);
  if(h<=100) return h;
  if(h<=200) return 100 + (h-100)*0.5;
  if(h<=300) return 150 + (h-200)*0.25;
  return 175 + (h-300)*0.15;
}
function calc(d){
  const k = KHOSHE[d.khoshe] || KHOSHE.services;
  const ref = num(state.settings[k.ref]), hr = num(state.settings[k.hr]);
  const total = (d.nLow||0) + (d.nHigh||0);
  const pd = (d.stdHours>0 && d.sessionHours>0) ? Math.ceil(d.stdHours/d.sessionHours) : 0;
  const eff = effectiveHours(d.stdHours||0);
  const W1 = d.items.filter(i=>i.owner===1).reduce((a,i)=>a+i.w,0);
  const netRate = Math.round(ref*(100-W1)/100);
  const initial = Math.round(netRate*pd*total);
  const tariffPer = Math.round(hr*eff);
  const tariff = Math.round(tariffPer*(d.nHigh||0));
  const reg  = num(state.settings['fee.register'])*total;
  const coun = num(state.settings['fee.counseling'])*total;
  const exam = num(state.settings['fee.exam'])*total;
  return {ref,hr,total,pd,eff,W1,netRate,initial,tariffPer,tariff,reg,coun,exam,
          grand: initial+tariff+reg+coun+exam};
}

/* ---------- داده پیش‌فرض ---------- */
function defaultAccounts(s){
  return [1,2,3,4,5].map(i=>({title:s['acc'+i+'.title']||'',owner:s['acc'+i+'.owner']||'',
    sheba:s['acc'+i+'.sheba']||'',shetab:s['acc'+i+'.shetab']||''}));
}
function blankData(){
  return {centerName:'',headName:'',gDir:'',gDirName:'',partyTitle:'',partyRep:'',partyPost:'',
    partyReg:'',partyTel:'',partyAddr:'',letterNo:'',letterDate:'',
    courseTitle:'',courseCode:'',khoshe:'services',stdHours:0,sessionHours:4,nLow:0,nHigh:0,
    workshop:'',instructor:'',days:'',startTime:'',startDate:'',endDate:'',
    items:ITEMS.map(i=>({...i})), accounts:defaultAccounts(state.settings)};
}

/* ---------- ناوبری ---------- */
function show(id){ ['scrHome','scrForm','scrResult','scrRetrieve'].forEach(s=>$(s).classList.toggle('hidden',s!==id)); window.scrollTo(0,0); }
function showHome(){ show('scrHome'); }
function showRetrieve(){ show('scrRetrieve'); }
function newContract(){
  state.mode='new'; state.contractNo=null; state.trackingCode=null;
  state.data = blankData(); fillForm();
  $('editBanner').classList.add('hidden');
  $('btnSubmit').textContent='صدور قرارداد';
  gotoStep(1); show('scrForm');
}
function gotoStep(n){
  state.step=n;
  [1,2,3,4].forEach(i=>$('step'+i).classList.toggle('hidden',i!==n));
  document.querySelectorAll('#stepper li').forEach(li=>{
    const s=+li.dataset.s; li.classList.toggle('active',s===n); li.classList.toggle('done',s<n); });
  $('btnPrev').style.visibility = n===1?'hidden':'visible';
  $('btnNext').classList.toggle('hidden',n===4);
  $('btnSubmit').classList.toggle('hidden',n!==4);
  updateSummary(); window.scrollTo(0,0);
}
function nextStep(){ const e=validate(state.step); if(e.length){ alert('لطفاً خطاهای زیر را برطرف کنید:\n\n• '+e.join('\n• ')); return; } if(state.step<4) gotoStep(state.step+1); }
function prevStep(){ if(state.step>1) gotoStep(state.step-1); }

/* ---------- اعتبارسنجی ---------- */
function validate(n){
  const d=state.data, e=[];
  if(n===1){
    [['centerName','نام مرکز'],['headName','نام رئیس مرکز'],['gDir','عنوان اداره کل'],
     ['gDirName','نام مدیرکل'],['partyTitle','عنوان دستگاه طرف قرارداد'],
     ['partyRep','نام نماینده'],['partyPost','عنوان پست نماینده']]
     .forEach(([k,l])=>{ if(!d[k].trim()) e.push(l+' الزامی است'); });
  }
  if(n===2){
    if(!d.courseTitle.trim()) e.push('عنوان دوره الزامی است');
    if(!(d.stdHours>0)) e.push('ساعت استاندارد باید عددی بزرگ‌تر از صفر باشد');
    if(d.nLow+d.nHigh<1) e.push('تعداد کارآموزان نمی‌تواند صفر باشد');
  }
  if(n===3){
    d.items.forEach(it=>{ if(!Number.isInteger(it.w)||it.w<1||it.w>25) e.push('وزن «'+it.name+'» باید عدد صحیح بین ۱ تا ۲۵ باشد'); });
    const s=d.items.reduce((a,i)=>a+i.w,0);
    if(s!==100) e.push('جمع وزن‌ها باید دقیقاً ۱۰۰ باشد (الان '+faD(s)+' است)');
  }
  return e;
}

/* ---------- خواندن/نوشتن فرم ---------- */
function readForm(){
  const d=state.data, g=id=>$(id)?$(id).value:'';
  d.centerName=g('fCenter').trim(); d.headName=g('fHead').trim();
  d.gDir=g('fGDir').trim(); d.gDirName=g('fGDirName').trim();
  d.partyTitle=g('fParty').trim(); d.partyRep=g('fRep').trim(); d.partyPost=g('fPost').trim();
  d.partyReg=g('fReg').trim(); d.partyTel=g('fTel').trim(); d.partyAddr=g('fAddr').trim();
  d.letterNo=g('fLetterNo').trim(); d.letterDate=g('fLetterDate').trim();
  d.courseTitle=g('fCourse').trim(); d.courseCode=g('fCode').trim();
  d.khoshe=g('fKhoshe')||'services';
  d.stdHours=Math.round(num(g('fHours'))); d.sessionHours=num(g('fSession'))||4;
  d.nLow=Math.max(0,Math.round(num(g('fLow')))); d.nHigh=Math.max(0,Math.round(num(g('fHigh'))));
  d.workshop=g('fWorkshop').trim(); d.instructor=g('fInstructor').trim();
  d.days=g('fDays').trim(); d.startTime=g('fStart').trim();
  d.startDate=g('fStartDate').trim(); d.endDate=g('fEndDate').trim();
}
function fillForm(){
  const d=state.data, set=(id,v)=>{ if($(id)) $(id).value = v??''; };
  set('fCenter',d.centerName); set('fHead',d.headName); set('fGDir',d.gDir); set('fGDirName',d.gDirName);
  set('fParty',d.partyTitle); set('fRep',d.partyRep); set('fPost',d.partyPost);
  set('fReg',d.partyReg); set('fTel',d.partyTel); set('fAddr',d.partyAddr);
  set('fLetterNo',d.letterNo); set('fLetterDate',d.letterDate);
  set('fCourse',d.courseTitle); set('fCode',d.courseCode); set('fKhoshe',d.khoshe);
  set('fHours',d.stdHours||''); set('fSession',String(d.sessionHours));
  set('fLow',d.nLow); set('fHigh',d.nHigh);
  set('fWorkshop',d.workshop); set('fInstructor',d.instructor); set('fDays',d.days);
  set('fStart',d.startTime); set('fStartDate',d.startDate); set('fEndDate',d.endDate);
}

/* ---------- ویرایشگر نهاده‌ها ---------- */
function renderItems(){
  const tb=document.querySelector('#itemsTable tbody'); tb.innerHTML='';
  state.data.items.forEach((it,ix)=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${faD(ix+1)}</td><td class="rt">${it.name}</td>
      <td><input type="number" min="1" max="25" step="1" value="${it.w}"/></td>
      <td><div class="seg"><button class="${it.owner===1?'on':''}" data-o="1">نهاد ۱</button>
      <button class="${it.owner===2?'on':''}" data-o="2">نهاد ۲</button></div></td>`;
    tr.querySelector('input').addEventListener('input',ev=>{ it.w=Math.max(0,Math.round(num(ev.target.value))); updateW(); updateSummary(); });
    tr.querySelectorAll('.seg button').forEach(b=>b.addEventListener('click',()=>{ it.owner=+b.dataset.o; renderItems(); updateSummary(); }));
    tb.appendChild(tr);
  });
  updateW();
}
function updateW(){
  const s=state.data.items.reduce((a,i)=>a+i.w,0);
  $('wSum').textContent=faD(s);
  const st=$('wState');
  st.textContent = s===100?'✓':(s>100?'— بیشتر از حد مجاز':'— کمتر از ۱۰۰');
  st.className = s===100?'ok':'';
}

/* ---------- ویرایشگر حساب‌ها ---------- */
function renderAccounts(){
  const tb=document.querySelector('#accTable tbody'); tb.innerHTML='';
  const c=calc(state.data), amounts=[c.tariff,c.reg,c.coun,c.exam,c.initial];
  state.data.accounts.forEach((a,ix)=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${faD(ix+1)}</td><td><input value="${esc(a.title)}"/></td>
      <td><input value="${esc(a.owner)}"/></td><td class="mono">${fmt(amounts[ix])}</td>
      <td><input value="${esc(a.shetab)}"/></td><td><input value="${esc(a.sheba)}"/></td>`;
    const ins=tr.querySelectorAll('input');
    ins[0].addEventListener('input',e=>a.title=e.target.value);
    ins[1].addEventListener('input',e=>a.owner=e.target.value);
    ins[2].addEventListener('input',e=>a.shetab=e.target.value);
    ins[3].addEventListener('input',e=>a.sheba=e.target.value);
    tb.appendChild(tr);
  });
}

/* ---------- پنل محاسبات زنده + بازبینی ---------- */
function updateSummary(){
  readForm();
  const d=state.data, c=calc(d);
  if($('fPersonDays')) $('fPersonDays').value = c.pd?faD(c.pd):'—';
  const rows=[
    ['خوشه',KHOSHE[d.khoshe].name],
    ['ساعت استاندارد',d.stdHours?faD(d.stdHours):'—'],
    ['ساعت هر جلسه',faD(String(d.sessionHours).replace('.','٫'))],
    ['نفر-روز (گرد به بالا)',c.pd||'—'],
    ['ساعت مؤثر',faD(c.eff)],
    ['مبلغ مرجع هر نفر-روز',fmt(c.ref)],
    ['جمع وزن نهاد ۱',faD(c.W1)+'٪'],
    ['نرخ هر نفر-روز پس از کسر',fmt(c.netRate)],
    ['تعداد کل کارآموزان',c.total||'—'],
    ['مبلغ اولیه قرارداد',fmt(c.initial)],
    ['تعرفه هر کارآموز دهک ۶+',fmt(c.tariffPer)],
    ['جمع تعرفه آموزشی',fmt(c.tariff)],
    ['جمع حق ثبت‌نام',fmt(c.reg)],
    ['جمع حق مشاوره',fmt(c.coun)],
    ['جمع حق آزمون',fmt(c.exam)],
  ];
  $('summaryBody').innerHTML =
    rows.map(([k,v])=>`<div class="srow"><span>${k}</span><b>${v}</b></div>`).join('')+
    `<div class="srow total"><span>مبلغ کل قرارداد</span><b>${fmt(c.grand)} ریال</b></div>`;
  if(state.step===4){ renderAccounts(); renderReview(c); }
}
function renderReview(c){
  const d=state.data, li=(k,v)=>`<div><b>${k}:</b> ${v}</div>`;
  $('review').innerHTML =
    li('طرف اول',`${d.headName} — رئیس ${d.centerName}`)+
    li('طرف دوم (نهاد ۱)',`${d.partyRep} (${d.partyPost}) — ${d.partyTitle}`)+
    li('دوره',`${d.courseTitle} — خوشه ${KHOSHE[d.khoshe].name} — ${faD(d.stdHours)} ساعت — ${faD(c.pd)} نفر-روز`)+
    li('کارآموزان',`دهک ۵−: ${faD(d.nLow)} نفر — دهک ۶+: ${faD(d.nHigh)} نفر — جمع: ${faD(c.total)} نفر`)+
    li('نرخ هر نفر-روز',`${fmt(c.netRate)} ریال (مرجع ${fmt(c.ref)} منهای ${faD(c.W1)}٪ نهاد ۱)`)+
    li('مبلغ اولیه قرارداد',fmt(c.initial)+' ریال')+
    li('جمع تعرفه آموزشی دهک ۶+',fmt(c.tariff)+' ریال')+
    li('ثبت‌نام / مشاوره / آزمون',`${fmt(c.reg)} / ${fmt(c.coun)} / ${fmt(c.exam)} ریال`)+
    `<div style="margin-top:6px;border-top:1px dashed var(--line);padding-top:8px"><b>مبلغ کل قرارداد: ${fmt(c.grand)} ریال</b> <span style="color:var(--muted)">(${num2fa(c.grand)} ریال)</span></div>`;
}

/* ---------- ارتباط با سرور ---------- */
async function api(action,payload={}){
  const r=await fetch(APPS_SCRIPT_URL,{method:'POST',body:JSON.stringify({action,...payload})});
  return r.json();
}
 $('btnSubmit').addEventListener('click',submit);
async function submit(){
  readForm();
  const errs=[...validate(1),...validate(2),...validate(3)];
  $('finalErr').textContent='';
  if(errs.length){ $('finalErr').innerHTML='• '+errs.join('<br>• '); return; }
  const btn=$('btnSubmit'); btn.disabled=true; btn.textContent='در حال ثبت…';
  try{
    const res=await api(state.mode==='new'?'save':'update',
      {data:{...state.data,grandTotal:calc(state.data).grand}});
    if(!res.ok) throw new Error(res.error||'خطای ناشناخته');
    state.contractNo=res.contractNo; state.trackingCode=res.trackingCode;
    $('resNo').textContent=faD(state.contractNo);
    $('resCode').textContent=faD(state.trackingCode);
    show('scrResult');
  }catch(ex){
    $('finalErr').textContent='خطا در ذخیره‌سازی: '+ex.message+' — اتصال اینترنت و آدرس Apps Script را بررسی کنید.';
  }
  btn.disabled=false; btn.textContent = state.mode==='new'?'صدور قرارداد':'ذخیره تغییرات و صدور مجدد خروجی‌ها';
}
async function retrieve(){
  const no=normD($('rNo').value).trim(), tc=normD($('rCode').value).trim();
  $('rErr').textContent='';
  if(!no||!tc){ $('rErr').textContent='هر دو فیلد الزامی است.'; return; }
  try{
    const res=await api('get',{contractNo:no,trackingCode:tc});
    if(!res.ok){ $('rErr').textContent='قراردادی با این شماره و کد رهگیری یافت نشد.'; return; }
    state.mode='edit'; state.contractNo=res.contractNo; state.trackingCode=res.trackingCode;
    state.data={...blankData(),...res.data};
    if(!state.data.items||!state.data.items.length) state.data.items=ITEMS.map(i=>({...i}));
    if(!state.data.accounts||!state.data.accounts.length) state.data.accounts=defaultAccounts(state.settings);
    fillForm();
    $('editBanner').textContent=`✎ ویرایش قرارداد شماره ${faD(no)} — کد رهگیری ${faD(tc)}`;
    $('editBanner').classList.remove('hidden');
    $('btnSubmit').textContent='ذخیره تغییرات و صدور مجدد خروجی‌ها';
    gotoStep(1); show('scrForm');
  }catch(ex){ $('rErr').textContent='خطا در ارتباط با سرور: '+ex.message; }
}
function copyVal(id,btn){
  const raw=String(id==='resNo'?state.contractNo:state.trackingCode);
  (navigator.clipboard?navigator.clipboard.writeText(raw):Promise.reject())
    .then(()=>{ btn.textContent='کپی شد ✓'; setTimeout(()=>btn.textContent='کپی',1200); })
    .catch(()=>{ btn.textContent=raw; });
}
function banner(msg){ const b=$('connBanner'); b.textContent=msg; b.classList.remove('hidden'); }

/* ---------- راه‌اندازی ---------- */
(async function init(){
  state.data=blankData();
  document.querySelectorAll('#step1 input,#step2 input,#step2 select')
    .forEach(el=>el.addEventListener('input',updateSummary));
  renderItems(); fillForm(); gotoStep(1); show('scrHome');
  if(APPS_SCRIPT_URL.includes('PASTE')){
    banner('آدرس Apps Script هنوز تنظیم نشده است؛ محاسبات و خروجی‌ها کار می‌کنند اما ذخیره‌سازی و شماره‌گذاری غیرفعال است.');
    return;
  }
  try{
    const r=await api('getSettings');
    if(r.ok){
      state.settings={...DEFAULT_SETTINGS,...r.settings};
      state.data=blankData(); fillForm(); updateSummary();
    }
  }catch(e){ banner('اتصال به بانک اطلاعاتی برقرار نشد؛ از نرخ‌های پیش‌فرض ۱۴۰۵ استفاده می‌شود.'); }
})();
