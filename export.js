/* ================= خروجی Word و Excel ================= */
function download(blob,name){
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),5000);
}

/* عنوان کامل اداره کل — اگر کاربر فقط نام استان وارد کرده باشد، به‌صورت کامل ساخته می‌شود */
function gdirFull(d){
  const t=(d.gDir||'').trim();
  return t ? (t.includes('اداره کل') ? t : 'اداره کل آموزش فنی و حرفه‌ای استان ' + t)
           : 'اداره کل آموزش فنی و حرفه‌ای استان .........';
}

/* تبدیل عدد به حروف فارسی */
function num2fa(n){
  n=Math.round(n); if(n===0) return 'صفر';
  const y=['','یک','دو','سه','چهار','پنج','شش','هفت','هشت','نه'];
  const t=['ده','یازده','دوازده','سیزده','چهارده','پانزده','شانزده','هفده','هجده','نوزده'];
  const d=['','','بیست','سی','چهل','پنجاه','شصت','هفتاد','هشتاد','نود'];
  const s=['','صد','دویست','سیصد','چهارصد','پانصد','ششصد','هفتصد','هشتصد','نهصد'];
  const sc=['','هزار','میلیون','میلیارد','هزار میلیارد'];
  const three=x=>{ const p=[],c=Math.floor(x/100),r=x%100;
    if(c)p.push(s[c]);
    if(r>=10&&r<20)p.push(t[r-10]);
    else{ const dd=Math.floor(r/10),yy=r%10; if(dd)p.push(d[dd]); if(yy)p.push(y[yy]); }
    return p.join(' و '); };
  const g=[]; let x=n; while(x>0){ g.push(x%1000); x=Math.floor(x/1000); }
  const out=[];
  for(let i=g.length-1;i>=0;i--){ if(!g[i])continue; out.push(three(g[i])+(sc[i]?' '+sc[i]:'')); }
  return out.join(' و ');
}

/* ================= ساخت فایل Word ================= */
async function exportWord(){
  const d=state.data, c=calc(d), s=state.settings;
  const no=state.contractNo||'پیش‌نویس', today=faD(jalaliToday());
  const kh=KHOSHE[d.khoshe].name;
  const L=[];

  /* --- ابزارهای ساخت سند (راست‌به‌چپ، B Nazanin) --- */
  const T=(t,o={})=>new docx.TextRun({text:String(t),bold:!!o.b,size:o.s||24,sizeComplexScript:o.s||24,rightToLeft:true,font:'B Nazanin'});
  const P=(t,o={})=>new docx.Paragraph({bidirectional:true,
    alignment:o.a??docx.AlignmentType.JUSTIFIED,spacing:{after:o.af??120,line:o.ln??320},
    children:Array.isArray(t)?t:[T(t,o)]});
  const EMPTY=()=>new docx.Paragraph({text:'',bidirectional:true});
  const C=(t,o={})=>new docx.TableCell({verticalAlign:docx.VerticalAlign.CENTER,
    width:o.w?{size:o.w,type:docx.WidthType.PERCENTAGE}:undefined,
    shading:o.f?{fill:o.f}:undefined,
    children:[new docx.Paragraph({bidirectional:true,alignment:docx.AlignmentType.CENTER,
      spacing:{after:0},children:[T(t,{b:o.b,s:o.s||20})]})]});
  const R=(cs,o={})=>new docx.TableRow({children:cs,tableHeader:!!o.h});
  const TB=rs=>new docx.Table({visuallyRightToLeft:true,width:{size:100,type:docx.WidthType.PERCENTAGE},rows:rs});

  /* --- سربرگ --- */
  L.push(P('قرارداد همکاری آموزشی',{a:docx.AlignmentType.CENTER,b:1,s:30,af:160}));
  L.push(P(`شماره: ${faD(no)}          تاریخ: ${today}`,{a:docx.AlignmentType.CENTER,s:22}));
  L.push(EMPTY());

  /* --- مقدمه --- */
  L.push(P('مقدمه',{b:1,a:docx.AlignmentType.RIGHT,af:120}));
  L.push(P(`به استناد شیوه‌نامه اجرایی خدمات آموزشی و فرهنگی سازمان آموزش فنی و حرفه‌ای کشور و در راستای ارتقای سطح شایستگی‌های حرفه‌ای متقاضیان، این قرارداد آموزشی بین ${d.partyTitle} به نمایندگی ${d.partyRep} به‌عنوان ${d.partyPost}، به آدرس: ${d.partyAddr||'............'} به شماره تلفن ${d.partyTel||'............'}، که از این پس طرف اول نامیده می‌شود، و ${gdirFull(d)} به نمایندگی ${d.gDirName} به‌عنوان مدیرکل به نشانی: ${s['gDir.addr']||'............'} به شماره تلفن ${s['gDir.phone']||'............'} که از این پس طرف دوم نامیده می‌شود، به شرح ذیل توافق و به امضاء می‌رسد.`));
  L.push(EMPTY());

  /* --- ماده ۱ --- */
  L.push(P('ماده ۱- موضوع و هدف:',{b:1}));
  L.push(P(`موضوع و هدف قرارداد حاضر، آموزش‌رسانی مهارت به هنرجویان فنی و حرفه‌ای آزاد در دوره «${d.courseTitle}» متعلق به خوشه «${kh}» به مدت ${faD(d.stdHours)} ساعت معادل ${faD(c.pd)} نفر-روز و برای تعداد ${faD(c.total)} نفر طبق استاندارد شماره ${faD(d.courseCode||'............')} می‌باشد.`));
  L.push(P('الف) ارائه مدرک در جهت اتمام مسیر دوره‌های آموزشی مهارتی؛'));
  L.push(P('ب) اطلاع‌رسانی و ارزشیابی مهارت‌های فراگرفته‌شده توسط فراگیران؛'));
  L.push(P('ج) تسهیل‌گری در جهت اجرای مسیر دوره‌های آموزشی مهارتی.'));
  L.push(EMPTY());

  /* --- ماده ۲ --- */
  const fr=num(s['fee.register']), fc=num(s['fee.counseling']), fe=num(s['fee.exam']);
  L.push(P('ماده ۲- هزینه اجرای دوره:',{b:1}));
  L.push(P(`نفر-روز معادل حاصل تقسیم ساعت استاندارد دوره بر ساعت آموزشی هر جلسه (که فقط می‌تواند ۲، ۲٫۵، ۳، ۳٫۵ یا ۴ ساعت باشد) است. بر این اساس هزینه اجرای دوره معادل ${faD(c.pd)} نفر-روز × ${faD(c.total)} نفر × ${fmt(c.netRate)} ریال جمعاً ${fmt(c.initial)} ریال (${num2fa(c.initial)} ریال) می‌باشد که طرف اول موظف است پیش از شروع دوره مطابق جدول ماده ۴ این قرارداد نسبت به واریز آن اقدام نماید.`));
  L.push(P('تبصره ۱- تعرفه آموزشی کارآموزان دارای دهک درآمدی ۵ و پایین‌تر صفر و سهم دولتی محسوب می‌گردد.'));
  L.push(P(`تبصره ۲- تعرفه آموزشی هر کارآموز دارای دهک درآمدی ۶ و بالاتر معادل ${fmt(c.hr)} ریال بابت هر ساعت مؤثر است و با احتساب ${faD(c.eff)} ساعت مؤثر این دوره، جمعاً ${fmt(c.tariff)} ریال به شرح ردیف ۱ جدول ماده ۴ به حساب خزانه‌داری کل واریز می‌گردد.`));
  L.push(P(`تبصره ۳- حق ثبت‌نام هر کارآموز ${fmt(fr)} ریال، جمعاً ${fmt(c.reg)} ریال، به شرح ردیف ۲ جدول ماده ۴ به حساب موقت مرکز واریز می‌گردد.`));
  L.push(P(`تبصره ۴- حق مشاوره هر کارآموز ${fmt(fc)} ریال، جمعاً ${fmt(c.coun)} ریال، به شرح ردیف ۳ جدول ماده ۴ به حساب موقت مرکز واریز می‌گردد.`));
  L.push(P(`تبصره ۵- حق آزمون هر کارآموز ${fmt(fe)} ریال، جمعاً ${fmt(c.exam)} ریال، به شرح ردیف ۴ جدول ماده ۴ به حساب موقت مرکز واریز می‌گردد.`));
  L.push(P(`تبصره ۶- مبلغ کل قرارداد معادل جمع مبالغ پنج‌گانه جدول ماده ۴ بالغ بر ${fmt(c.grand)} ریال (${num2fa(c.grand)} ریال) می‌باشد و دریافت هرگونه وجه از کارآموزان ممنوع است.`));
  L.push(EMPTY());

  /* --- ماده ۳ + جدول اجرا --- */
  L.push(P('ماده ۳- زمان و مکان برگزاری دوره آموزشی:',{b:1}));
  L.push(P(`روزها و ساعات برگزاری دوره می‌تواند به مدت ${faD(c.pd)} نفر-روز از تاریخ ${d.startDate?faD(d.startDate):'............'} لغایت ${d.endDate?faD(d.endDate):'............'} در کارگاه ${d.workshop||'............'} به شرح جدول ذیل برگزار گردد:`));
  L.push(TB([
    R([C('ردیف',{b:1,f:'E8EDF3',w:7}),C('نام کارگاه',{b:1,f:'E8EDF3',w:15}),C('استاد',{b:1,f:'E8EDF3',w:13}),
       C('ایام هفته',{b:1,f:'E8EDF3',w:12}),C('ساعت برگزاری',{b:1,f:'E8EDF3',w:11}),C('مدت اجرا',{b:1,f:'E8EDF3',w:14}),
       C('تعداد شرکت‌کننده',{b:1,f:'E8EDF3',w:9}),C('نفر-روز',{b:1,f:'E8EDF3',w:7}),
       C('هزینه هر نفر-روز (ریال)',{b:1,f:'E8EDF3',w:9}),C('مبلغ کل (ریال)',{b:1,f:'E8EDF3',w:12})],{h:1}),
    R([C('۱'),C(d.workshop||'—'),C(d.instructor||'—'),C(d.days||'—'),C(d.startTime||'—'),
       C(`از ${d.startDate||'—'} تا ${d.endDate||'—'}`),C(faD(c.total)),C(faD(c.pd)),C(fmt(c.netRate)),C(fmt(c.initial))]),
    R([new docx.TableCell({columnSpan:7,verticalAlign:docx.VerticalAlign.CENTER,
        children:[P('جمع',{a:docx.AlignmentType.CENTER,b:1,s:20,af:0})]}),
       C(faD(c.pd),{b:1}),C(fmt(c.netRate),{b:1}),C(fmt(c.initial),{b:1})])
  ]));
  L.push(EMPTY());

  /* --- جدول تجزیه و تحلیل خدمات (آنالیز نهاده‌ها) --- */
  L.push(P('جدول تجزیه و تحلیل خدمات آموزشی (آنالیز نهاده‌ها):',{b:1,af:80}));
  const AR=[R([C('ردیف',{b:1,f:'E8EDF3',w:6}),C('نام نهاد آموزشی',{b:1,f:'E8EDF3',w:28}),
    C('نهاد ۱ (طرف قرارداد)',{b:1,f:'E8EDF3',w:11}),C('نهاد ۲ (فنی و حرفه‌ای)',{b:1,f:'E8EDF3',w:11}),
    C('درصد وزنی',{b:1,f:'E8EDF3',w:9}),C('سهم مصوب خوشه در یک نفر-روز (ریال)',{b:1,f:'E8EDF3',w:13}),
    C('سهم اداره‌کل پس از کسر نهاد ۱ (ریال)',{b:1,f:'E8EDF3',w:13})],{h:1})];
  d.items.forEach((it,i)=>AR.push(R([C(faD(i+1)),C(it.name),C(it.owner===1?'✓':'−'),
    C(it.owner===2?'✓':'−'),C(faD(it.w)+'٪'),C(fmt(c.ref*it.w/100)),C(fmt(Math.round(c.netRate*it.w/100)))])));
  AR.push(R([new docx.TableCell({columnSpan:4,verticalAlign:docx.VerticalAlign.CENTER,shading:{fill:'F5F5F5'},
      children:[P('جمع',{a:docx.AlignmentType.CENTER,b:1,s:20,af:0})]}),
    C('۱۰۰٪',{b:1,f:'F5F5F5'}),C(fmt(c.ref),{b:1,f:'F5F5F5'}),C(fmt(c.netRate),{b:1,f:'F5F5F5'})]));
  AR.push(R([new docx.TableCell({columnSpan:6,verticalAlign:docx.VerticalAlign.CENTER,shading:{fill:'FFF3CD'},
      children:[P('جمع هزینه‌ی یک نفر-روز خدمات آموزشی (ریال)',{a:docx.AlignmentType.CENTER,b:1,s:20,af:0})]}),
    C(fmt(c.netRate),{b:1,f:'FFF3CD'})]));
  L.push(TB(AR)); L.push(EMPTY());

  /* --- ماده ۴ + جدول حساب --- */
  L.push(P('ماده ۴- نحوه پرداخت:',{b:1}));
  L.push(P('مبلغ قرارداد به شرح جدول ذیل توسط طرف اول به حساب‌های اعلامی واریز می‌گردد و واریز مبالغ پیش از شروع دوره الزامی است:'));
  const amounts=[c.tariff,c.reg,c.coun,c.exam,c.initial];
  const AC=[R([C('ردیف',{b:1,f:'E8EDF3',w:6}),C('عنوان حساب',{b:1,f:'E8EDF3',w:22}),
    C('صاحب حساب',{b:1,f:'E8EDF3',w:14}),C('مبلغ (ریال)',{b:1,f:'E8EDF3',w:12}),
    C('کد شتاب ۳۰ رقمی',{b:1,f:'E8EDF3',w:14}),C('شماره شبا',{b:1,f:'E8EDF3',w:16})],{h:1})];
  d.accounts.forEach((a,i)=>AC.push(R([C(faD(i+1)),C(a.title),C(a.owner),C(fmt(amounts[i])),C(a.shetab),C(a.sheba)])));
  AC.push(R([new docx.TableCell({columnSpan:3,verticalAlign:docx.VerticalAlign.CENTER,shading:{fill:'E2EFDA'},
      children:[P('جمع کل قرارداد',{a:docx.AlignmentType.CENTER,b:1,s:20,af:0})]}),
    C(fmt(c.grand),{b:1,f:'E2EFDA'}),
    new docx.TableCell({columnSpan:2,children:[P('',{af:0})]})]));
  L.push(TB(AC));
  L.push(P('تبصره- ردیف‌های ۱ و ۵ جدول فوق به حساب خزانه‌داری کل واریز می‌گردد و ردیف‌های ۲ تا ۴ (حق ثبت‌نام، مشاوره و آزمون) به حساب موقت مرکز واریز و توسط کارشناس مرکز در پورتال سازمان تودیع می‌گردد. دریافت هرگونه وجه از کارآموزان ممنوع است.',{s:22}));
  L.push(EMPTY());

  /* --- مواد ۵ تا ۱۱ --- */
  const pCenter=d.items.filter(i=>i.owner===2).map(i=>`${i.name} (${faD(i.w)}٪)`);
  const pParty=d.items.filter(i=>i.owner===1).map(i=>`${i.name} (${faD(i.w)}٪)`);
  L.push(P('ماده ۵- تعهدات طرف دوم (اداره کل و مرکز آموزش فنی و حرفه‌ای):',{b:1}));
  L.push(P('۱- اجرای دوره آموزشی موضوع این قرارداد مطابق استاندارد مصوب سازمان و در چارچوب ساعت‌های مندرج در ماده ۱؛'));
  L.push(P('۲- تأمین و به‌کارگیری مربی/استاد مجرب و رابط آموزشی کارآمد؛'));
  L.push(P('۳- مشاوره و راهنمایی آموزشی به فراگیران؛'));
  L.push(P('۴- برگزاری آزمون، تصحیح، ثبت نمرات و صدور گواهی‌نامه/تأییدیه برای افراد شرکت‌کننده در دوره؛'));
  L.push(P('۵- نگارش و ارائه پرونده آموزشی دوره و مستندسازی الکترونیکی فرایند آموزش؛'));
  L.push(P('۶- برگزاری مراسم اعطای مدارک در پایان دوره برای افراد شرکت‌کننده در دوره‌های آموزشی.'));
  if(pCenter.length) L.push(P(`۷- تأمین اقلام ذیل مطابق جدول تجزیه و تحلیل خدمات آموزشی: ${pCenter.join('، ')}.`));
  L.push(EMPTY());
  L.push(P('ماده ۶- تعهدات طرف اول (نهاد ۱):',{b:1}));
  L.push(P('۱- پرداخت هزینه‌های مندرج در ماده ۲ و ماده ۴ این قرارداد در موعد مقرر به حساب‌های اعلامی؛'));
  L.push(P('۲- معرفی کارشناس/نماینده به‌صورت مکتوب جهت هماهنگی و نظارت بر اجرای دوره؛'));
  if(pParty.length) L.push(P(`۳- تأمین و برعهده‌گیری اقلام ذیل مطابق جدول تجزیه و تحلیل خدمات آموزشی: ${pParty.join('، ')}.`));
  L.push(P('۴- اعلام به‌موقع فهرست نام کارآموزان و مشخصات دهک درآمدی آنان به مرکز.'));
  L.push(EMPTY());
  L.push(P('ماده ۷- تعهدات مشترک:',{b:1}));
  L.push(P('۱- همکاری در برگزاری سنجش و انجام مطالعات آینده‌پژوهی متقابل؛'));
  L.push(P('۲- همکاری در برگزاری نمایشگاه، کنفرانس و مسابقات مهارت‌های تخصصی؛'));
  L.push(P('۳- تدوین و بازنگری استاندارد آموزشی و نمایشگاه شغل؛'));
  L.push(P('۴- همکاری در نظارت بر اجرای تعهدات و ارائه گزارش.'));
  L.push(EMPTY());
  L.push(P('ماده ۸- پیش‌بینی‌های لازم:',{b:1}));
  L.push(P(`۱- کلیه هماهنگی‌های اجرایی این قرارداد از طریق ${d.headName} (رئیس مرکز آموزش فنی و حرفه‌ای ${d.centerName}) به‌عنوان مسئول قرارداد انجام می‌پذیرد؛`));
  L.push(P(`۲- ${d.partyRep} (${d.partyPost}) به‌عنوان نماینده طرف اول در کلیه مراحل اجرا، تحویل و تحول و مفاصاحساب قرارداد شناخته می‌شود؛`));
  L.push(P('۳- تحویل فضاهای آموزشی، تجهیزات و اقلام برعهده‌گرفته‌شده طرف اول، پیش از آغاز دوره الزامی است.'));
  L.push(EMPTY());
  L.push(P('ماده ۹- نامشخص بودن زمان و مکان برگزاری:',{b:1}));
  L.push(P('هرگاه زمان و یا مکان برگزاری دوره‌های آموزشی بر اثر عوامل پیش‌بینی‌نشده دچار تغییر و تحول شود، مراتب با توافق طرفین تعیین و به امضای ایشان می‌رسد.'));
  L.push(EMPTY());
  L.push(P('ماده ۱۰- کمیسیون اختلافات:',{b:1}));
  L.push(P('در اجرای ماده ۱۱ قانون الحاق یک ماده به قانون برنامه و بودجه سال ۱۳۷۸، هرگاه بین طرفین اختلافی بروز نماید، موضوع در کمیسیونی مرکب از نمایندگان طرفین (هر طرف یک نفر) مطرح می‌گردد؛ رئیس بر حسب مورد، طرح اختلاف و راه‌حل مناسب را بررسی و اعلام می‌نماید و نظر کمیسیون در چارچوب مفاد این قرارداد برای طرفین لازم‌الاجرا است.'));
  L.push(EMPTY());
  L.push(P('ماده ۱۱- کلیات قرارداد:',{b:1}));
  L.push(P('این قرارداد در یازده ماده و هفت تبصره به همراه پیوست‌های آن (جدول زمان اجرا، جدول تجزیه و تحلیل خدمات آموزشی و جدول نحوه پرداخت) تنظیم و به امضای طرفین رسیده و از تاریخ امضا برای طرفین حکم اجرا دارد و مفاد همه نسخه‌ها یکسان و واحد است. یک نسخه جهت اداره مالی و یک نسخه جهت معاونت آموزشی ارسال می‌گردد.'));
  L.push(EMPTY());

  /* --- امضاها --- */
  const NB={style:docx.BorderStyle.NONE,size:0,color:'FFFFFF'};
  const SP=(t,s)=>new docx.Paragraph({bidirectional:true,alignment:docx.AlignmentType.CENTER,
    spacing:{after:60},children:[T(t,{b:!!s,s:24})]});
  L.push(EMPTY());
  L.push(new docx.Table({visuallyRightToLeft:true,width:{size:100,type:docx.WidthType.PERCENTAGE},
    borders:{top:NB,bottom:NB,left:NB,right:NB,insideHorizontal:NB,insideVertical:NB},
    rows:[R([
      new docx.TableCell({verticalAlign:docx.VerticalAlign.CENTER,width:{size:50,type:docx.WidthType.PERCENTAGE},
        children:[SP('طرف اول',1),SP(d.partyRep),SP(d.partyPost+' — '+d.partyTitle),SP('مهر و امضا')]}),
      new docx.TableCell({verticalAlign:docx.VerticalAlign.CENTER,width:{size:50,type:docx.WidthType.PERCENTAGE},
        children:[SP('طرف دوم',1),SP(d.gDirName),SP('مدیرکل '+gdirFull(d)),SP('مهر و امضا')]})
    ])]}));

  /* --- بسته‌بندی --- */
  const doc=new docx.Document({
    creator:'سامانه قراردادهای آموزشی', title:'قرارداد همکاری آموزشی '+no,
    styles:{default:{document:{run:{font:'B Nazanin',size:24}}}},
    sections:[{properties:{page:{size:{width:11906,height:16838},
      margin:{top:1134,bottom:1134,right:1134,left:1134}}},children:L}]});
  const blob=await docx.Packer.toBlob(doc);
  download(blob,'Gharardad-'+no+'.docx');
}

/* ================= ساخت فایل Excel ================= */
function exportExcel(){
  const d=state.data, c=calc(d), no=state.contractNo||'jadid';
  const wb=XLSX.utils.book_new();
  const mk=(aoa,cols)=>{ const ws=XLSX.utils.aoa_to_sheet(aoa); ws['!cols']=cols; ws['!views']=[{rightToLeft:true}]; return ws; };

  const g1=[['سامانه صدور قراردادهای آموزشی — محاسبه‌گر'],[],
    ['شماره قرارداد',state.contractNo||'-'],['کد رهگیری',state.trackingCode||'-'],[],
    ['خوشه',KHOSHE[d.khoshe].name],['ساعت استاندارد',d.stdHours],['ساعت هر جلسه',d.sessionHours],
    ['نفر-روز (گرد به بالا)',c.pd],['ساعت مؤثر',c.eff],
    ['مبلغ مرجع هر نفر-روز (ریال)',c.ref],['جمع وزن نهاد ۱ (٪)',c.W1],
    ['نرخ هر نفر-روز پس از کسر (ریال)',c.netRate],['تعداد کل کارآموزان',c.total],
    ['مبلغ اولیه قرارداد (ریال)',c.initial],['تعرفه ساعتی خوشه (ریال)',c.hr],
    ['تعرفه هر کارآموز دهک ۶+ (ریال)',c.tariffPer],['تعداد کارآموزان دهک ۶+',d.nHigh],
    ['جمع تعرفه آموزشی (ریال)',c.tariff],['جمع حق ثبت‌نام (ریال)',c.reg],
    ['جمع حق مشاوره (ریال)',c.coun],['جمع حق آزمون (ریال)',c.exam],
    ['مبلغ کل قرارداد (ریال)',c.grand]];
  XLSX.utils.book_append_sheet(wb,mk(g1,[{wch:38},{wch:18}]),'محاسبه‌گر');

  const g2=[['ردیف','نام نهاد آموزشی','نهاد ۱ (طرف قرارداد)','نهاد ۲ (فنی و حرفه‌ای)','درصد وزنی','سهم مصوب خوشه در یک نفر-روز (ریال)','سهم اداره‌کل پس از کسر نهاد ۱ (ریال)']];
  d.items.forEach((it,i)=>g2.push([i+1,it.name,it.owner===1?1:0,it.owner===2?1:0,it.w,Math.round(c.ref*it.w/100),Math.round(c.netRate*it.w/100)]));
  g2.push(['','جمع','','',100,c.ref,c.netRate]);
  g2.push(['','جمع هزینه‌ی یک نفر-روز خدمات آموزشی (ریال)','','','','',c.netRate]);
  XLSX.utils.book_append_sheet(wb,mk(g2,[{wch:6},{wch:34},{wch:12},{wch:12},{wch:9},{wch:20},{wch:20}]),'تجزیه و تحلیل');

  const amounts=[c.tariff,c.reg,c.coun,c.exam,c.initial];
  const g3=[['ردیف','عنوان حساب','صاحب حساب','مبلغ (ریال)','کد شتاب ۳۰ رقمی','شماره شبا']];
  d.accounts.forEach((a,i)=>g3.push([i+1,a.title,a.owner,amounts[i],a.shetab,a.sheba]));
  g3.push(['','جمع کل قرارداد','',c.grand,'','']);
  XLSX.utils.book_append_sheet(wb,mk(g3,[{wch:6},{wch:32},{wch:24},{wch:16},{wch:28},{wch:28}]),'حساب قرارداد');

  const g4=[['ردیف','نام کارگاه','استاد','ایام هفته','ساعت برگزاری','مدت اجرا','تعداد شرکت‌کننده','نفر-روز','هزینه هر نفر-روز (ریال)','مبلغ کل (ریال)'],
    [1,d.workshop,d.instructor,d.days,d.startTime,`از ${d.startDate} تا ${d.endDate}`,c.total,c.pd,c.netRate,c.initial]];
  XLSX.utils.book_append_sheet(wb,mk(g4,[{wch:6},{wch:16},{wch:16},{wch:14},{wch:12},{wch:22},{wch:10},{wch:8},{wch:14},{wch:16}]),'اجرای دوره');

  XLSX.writeFile(wb,'MahasebeGharardad-'+no+'.xlsx');
}
