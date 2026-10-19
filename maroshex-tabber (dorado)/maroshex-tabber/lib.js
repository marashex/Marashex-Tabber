// Shared helpers: tiny xlsx reader, training normalizer, ICS parser, conflict check.
const unesc = s => s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
async function inflateRaw(u8){const ds=new DecompressionStream('deflate-raw');const w=ds.writable.getWriter();w.write(u8);w.close();return new Uint8Array(await new Response(ds.readable).arrayBuffer());}
async function unzip(buf){
  const u=new Uint8Array(buf),dv=new DataView(buf),td=new TextDecoder();let e=u.length-22;
  while(e>=0&&dv.getUint32(e,true)!==0x06054b50)e--;
  if(e<0)throw new Error('notzip');
  const n=dv.getUint16(e+10,true);let p=dv.getUint32(e+16,true);const out={};
  for(let i=0;i<n;i++){
    const method=dv.getUint16(p+10,true),cs=dv.getUint32(p+20,true),nl=dv.getUint16(p+28,true),el=dv.getUint16(p+30,true),cl=dv.getUint16(p+32,true),lo=dv.getUint32(p+42,true);
    out[td.decode(u.subarray(p+46,p+46+nl))]={method,cs,lo};p+=46+nl+el+cl;
  }
  return {names:Object.keys(out),get:async name=>{
    const f=out[name];if(!f)return null;
    const s=f.lo+30+dv.getUint16(f.lo+26,true)+dv.getUint16(f.lo+28,true),d=u.subarray(s,s+f.cs);
    return td.decode(f.method===0?d:await inflateRaw(d));}};
}
const colIdx=r=>{let n=0;for(const ch of r.replace(/\d/g,''))n=n*26+ch.charCodeAt(0)-64;return n-1;};
function parseSheet(xml,strs){
  const rows=[],links={};
  for(const rm of xml.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)){
    const rn=+(/r="(\d+)"/.exec(rm[1])||[])[1]-1,row=[],ri=isNaN(rn)?rows.length:rn;
    for(const cm of (rm[2]||'').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
      const ref=/r="([A-Z]+\d+)"/.exec(cm[1]);if(!ref)continue;
      const t=(/t="(\w+)"/.exec(cm[1])||[])[1],body=cm[2]||'',v=(/<v>([\s\S]*?)<\/v>/.exec(body)||[])[1];
      let val='';
      if(t==='s')val=strs[+v]||'';
      else if(t==='inlineStr')val=unesc([...body.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x=>x[1]).join(''));
      else if(v!==undefined)val=unesc(v);
      const fm=/<f[^>]*>([\s\S]*?)<\/f>/.exec(body);if(fm){const hm=/HYPERLINK\(\s*"([^"]+)"/i.exec(unesc(fm[1]));if(hm)links[ri+','+colIdx(ref[1])]=hm[1];}
      row[colIdx(ref[1])]=String(val).trim();
    }
    rows[isNaN(rn)?rows.length:rn]=Array.from(row,x=>x||'');
  }
  const res=Array.from(rows,r=>r||[]);res.links=links;return res;
}
// Reads EVERY visible tab: returns [{name, rows}]
async function parseWorkbook(buf){
  const z=await unzip(buf);
  const ss=await z.get('xl/sharedStrings.xml');
  const strs=ss?[...ss.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m=>unesc([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x=>x[1]).join(''))):[];
  const wb=(await z.get('xl/workbook.xml'))||'',rel=(await z.get('xl/_rels/workbook.xml.rels'))||'',tg={};
  for(const m of rel.matchAll(/<Relationship\b([^>]*?)\/?>/g)){const id=/\bId="([^"]+)"/.exec(m[1]),t=/Target="([^"]+)"/.exec(m[1]);if(id&&t)tg[id[1]]=t[1].replace(/^\/?(xl\/)?/,'xl/');}
  const out=[];
  for(const m of wb.matchAll(/<sheet\b([^>]*?)\/?>/g)){
    const nm=/name="([^"]*)"/.exec(m[1]),rid=/r:id="([^"]+)"/.exec(m[1]);
    if(!nm||!rid||/state="(hidden|veryHidden)"/.test(m[1]))continue;
    const xml=tg[rid[1]]&&await z.get(tg[rid[1]]);if(!xml)continue;
    const rows=parseSheet(xml,strs),rp=tg[rid[1]].replace(/([^\/]+)$/,'_rels/$1.rels'),rx=(await z.get(rp))||'',rt={};
    for(const m2 of rx.matchAll(/<Relationship\b([^>]*?)\/?>/g)){const id=/\bId="([^"]+)"/.exec(m2[1]),t=/Target="([^"]+)"/.exec(m2[1]);if(id&&t&&/External/i.test(m2[1]))rt[id[1]]=unesc(t[1]);}
    for(const hm of xml.matchAll(/<hyperlink\b([^>]*?)\/?>/g)){const ref=/ref="([A-Z]+)(\d+)/.exec(hm[1]),id=/r:id="([^"]+)"/.exec(hm[1]);if(ref&&id&&rt[id[1]])rows.links[(+ref[2]-1)+','+colIdx(ref[1])]=rt[id[1]];}
    out.push({name:unesc(nm[1]),rows});
  }
  if(!out.length){const sn=z.names.filter(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort()[0];if(sn)out.push({name:'Sheet1',rows:parseSheet(await z.get(sn),strs)});}
  return out;
}
// ---- normalize sheet rows into trainings (column names auto-detected from the header row)
const serialToDate=n=>{const d=new Date(Math.round((n-25569)*864e5));return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate(),d.getUTCHours(),d.getUTCMinutes());};
function minutes(str){ // "2pm", "14:30", "2:00 PM", or Excel fraction
  if(/^0?\.\d+$|^0$|^1$/.test(str)&&str!=='')return Math.round(parseFloat(str)*1440);
  const m=/(\d{1,2})(?::(\d{2}))?\s*(am|pm|a|p)?/i.exec(str||'');if(!m)return null;
  let h=+m[1];const mi=+(m[2]||0),ap=(m[3]||'').toLowerCase();
  if(ap.startsWith('p')&&h<12)h+=12;if(ap.startsWith('a')&&h===12)h=0;return h*60+mi;
}
const DATE_RE=/\b(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\b|\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2}\b|\b\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?\b/i;
const TIME_RE=/\b\d{1,2}(?::\d{2})?\s*(?:am|pm)\b|\b\d{1,2}:\d{2}\b/i;
const SIGN_RE=/student|attendee|participant|signed|sign.?up|^name|slot\s*#?\s*\d|seat\s*#?\s*\d|person\s*\d|registered|reserved|trainee|member/i;
const MAT_RE=/materials?/i;
const STRICT_DATE=/^(?:(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s+)?(?:(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2}|\d{1,2}[\/\-]\d{1,2})/i;
const normDate=s=>String(s||'').replace(/(\d)(?:st|nd|rd|th)\b/gi,'$1').replace(/\bSept\b/i,'Sep').replace(/\s+/g,' ').trim();
function toTrainings(rows,cfg={},sheetName=''){
  const HW=/date|time|start|end|avail|spot|seat|slot|sign|session|train|class|status|open|name|student|instructor|capacity|when|day/i;
  const hdrLike=r=>r.filter(x=>x&&x.length<=28&&!/[:,]/.test(x)&&HW.test(x)).length>=2;
  const firstDated=rows.findIndex(r=>!hdrLike(r)&&r.some(x=>STRICT_DATE.test(normDate(x))));
  let hi=-1;for(let i=0;i<rows.length&&(firstDated<0||i<firstDated);i++)if(hdrLike(rows[i]))hi=i;   // header row nearest above the first dated row
  const hdrs=hi>=0?rows[hi]:[],H=hdrs.map(h=>(h||'').toLowerCase()),width=Math.max(0,...rows.map(r=>r.length));
  const data=rows.filter((r,i)=>i!==hi),find=f=>H.findIndex(f);
  const byPattern=(re,norm)=>{let best=-1,bc=0;for(let j=0;j<width;j++){const n=data.filter(r=>re.test(norm?normDate(r[j]):(r[j]||''))).length;if(n>bc){bc=n;best=j;}}return best;};
  const c={date:find(h=>/date|day/.test(h)),time:find(h=>/time|start|when/.test(h)),end:find(h=>/^end|finish/.test(h)),
    avail:find(h=>/avail|spot|open|remain|capacity|status|left/.test(h)&&!/\d/.test(h)),
    type:find(h=>/training type|session type|^type\b|type$/.test(h))};
  const wi=find(h=>/trainer|instructor|staff|lead|run by|taught/.test(h));
  c.name=wi>=0?wi:find(h=>/class|course|title|topic|name/.test(h)&&!/student|attendee|participant|email/.test(h));
  if(c.date<0)c.date=byPattern(STRICT_DATE,true);
  if(c.time<0||c.time===c.date)c.time=byPattern(TIME_RE);
  if(hi<0&&c.name<0)c.name=[...Array(width).keys()].find(j=>j!==c.date&&j!==c.time)??-1;
  if(c.name===c.type)c.name=-1;
  for(const k of['name','date','time','avail','type'])if(cfg[k]!==undefined&&cfg[k]!=='')c[k]=+cfg[k];
  const sign=H.map((h,j)=>SIGN_RE.test(h)&&![c.name,c.date,c.time,c.end,c.avail,c.type].includes(j)?j:-1).filter(j=>j>=0);
  const meta={hi,headers:hdrs,c,sign},lc=c.date>=0?c.date:0;
  // links per row (hyperlinks + bare URLs)
  const byRow={},LK=rows.links||{};
  for(const k in LK){const [i,j]=k.split(',');(byRow[i]=byRow[i]||[]).push([+j,LK[k]]);}
  rows.forEach((r,i)=>r.forEach((x,j)=>{if(/^https?:\/\/\S+$/.test(x||''))(byRow[i]=byRow[i]||[]).push([j,x]);}));
  let tabMat=null;rows.forEach((r,i)=>{if(!tabMat&&byRow[i]&&r.some(x=>MAT_RE.test(x||'')))tabMat=byRow[i][0][1];});
  const out=[],yr=new Date().getFullYear(),rule=cfg.rule||'auto',todayMs=new Date().setHours(0,0,0,0),clean=x=>(x&&x!=='-'&&x!=='–')?x:'';
  const nm={},desc={},dmat={},notes=[];let names=false,lastTxt='',lastDay=null,sawDate=false,past=0,dated=0;
  for(let ri=0;ri<rows.length;ri++){
    if(ri===hi)continue;
    const texts=rows[ri].map(x=>(x||'').trim());if(!texts.some(Boolean))continue;
    let dateTxt=normDate(texts[lc]),timeTxt=c.time>=0?texts[c.time]||'':'',day=null,sm=null,em=null;
    if(c.time>=0&&c.time===c.date){const m=TIME_RE.exec(dateTxt);timeTxt=m?m[0]:'';dateTxt=dateTxt.replace(TIME_RE,'').replace(/\s+(at|@|-)?\s*$/i,'').trim();}
    if(/^\d{4,6}(\.\d+)?$/.test(dateTxt)&&+dateTxt>20000){const d=serialToDate(+dateTxt);day=new Date(d.getFullYear(),d.getMonth(),d.getDate());if(d.getHours()||d.getMinutes())sm=d.getHours()*60+d.getMinutes();}
    else if(dateTxt&&STRICT_DATE.test(dateTxt)){
      let t=Date.parse(/\d{4}/.test(dateTxt)?dateTxt:dateTxt+' '+yr);
      if(!isNaN(t)&&!/\d{4}/.test(dateTxt)&&t<todayMs-120*864e5)t=Date.parse(dateTxt+' '+(yr+1));
      if(!isNaN(t)){const d=new Date(t);day=new Date(d.getFullYear(),d.getMonth(),d.getDate());}
    }else if(!dateTxt&&TIME_RE.test(timeTxt)&&lastDay){day=lastDay;}
    if(!day){ // description / info row
      const lab=texts[lc]||'';
      if(/^(date|day)$/i.test(lab))continue;
      const vals=texts.map((x,j)=>[j,x]).filter(([j,x])=>x&&j!==lc&&!/^https?:\/\//.test(x));
      if(!vals.length){if(lab.length>20)notes.push({t:lab,u:''});continue;}
      if(!lab&&!names){vals.forEach(([j,x])=>nm[j]=x);names=true;continue;}
      const L0=lab.replace(/:\s*$/,''),isMat=texts.some(x=>MAT_RE.test(x));
      vals.forEach(([j,x])=>{const lk=(byRow[ri]||[]).find(p=>p[0]===j);(desc[j]=desc[j]||[]).push({t:(L0?L0+': ':'')+x,u:lk?lk[1]:''});if(lk&&isMat&&!dmat[j])dmat[j]=lk[1];});
      continue;
    }
    dateTxt=day.toLocaleDateString([],{weekday:'short',month:'short',day:'numeric'});
    sawDate=true;dated++;lastTxt=dateTxt;lastDay=day;
    const endTxt=c.end>=0?texts[c.end]||'':'';
    if(/^0?\.\d+$/.test(endTxt))em=Math.round(+endTxt*1440);else if(endTxt){const e2=minutes(endTxt);if(e2!==null)em=e2;}
    if(/^0?\.\d+$/.test(timeTxt)){sm=minutes(timeTxt);timeTxt=fmtMin(sm)+(em!==null?' – '+fmtMin(em):'');}
    else if(timeTxt){const parts=timeTxt.split(/\s*(?:-|–|—|to)\s*/i),a=minutes(parts[0]),b2=parts[1]?minutes(parts[1]):null;
      if(a!==null){sm=a;if(b2!==null){em=b2;if(/pm/i.test(parts[1])&&!/am|pm/i.test(parts[0])&&sm+720<=em)sm+=720;}}}
    else if(sm!==null)timeTxt=fmtMin(sm)+(em!==null?' – '+fmtMin(em):'');
    const start=sm!==null?day.getTime()+sm*6e4:null,end=start&&em!==null?day.getTime()+em*6e4:null;
    if(start?start<Date.now():day.getTime()<todayMs){past++;continue;}   // only today-and-later
    const cap=+cfg.cap||sign.length||1,tok=x=>x.split(/[,;\n]|\band\b|&/i).filter(s=>s.trim()).length;
    const cells=sign.length?sign.map(j=>texts[j]):texts.filter((x,j)=>x&&j!==c.name&&j!==c.date&&j!==c.time&&j!==c.end&&j!==c.type&&!/^https?:\/\//.test(x));
    const filled=cells.reduce((a,x)=>a+(x?tok(x):0),0);
    let availTxt=c.avail>=0?texts[c.avail]||'':'',open;
    if(rule==='empty_open'){open=!availTxt;availTxt=String(open?cap:0);}
    else if(rule==='filled_open'){open=!!availTxt;availTxt=String(open?cap:0);}
    else if(c.avail>=0&&availTxt){const n=parseFloat(availTxt);open=!isNaN(n)&&/^\s*-?\d/.test(availTxt)?n>0:!/full|closed|none|waitlist|taken|no\b/i.test(availTxt);}
    else{const sp=Math.max(0,cap-filled);open=sp>0;availTxt=String(sp);}
    const who=c.name>=0?clean(texts[c.name]):'',type=(c.type>=0?clean(texts[c.type]):'')||sheetName;
    out.push({key:[sheetName,type,who,dateTxt,timeTxt].join('|'),name:sheetName||type||'Training',type,who,sub:'',label:type+(who?' – '+who:''),info:'',mat:tabMat||'',date:dateTxt,time:timeTxt,start,end,avail:availTxt,open});
  }
  const cols=[...new Set([...Object.keys(nm),...Object.keys(desc)])].map(Number).sort((a,b)=>a-b);
  const secs=cols.map(j=>({type:nm[j]||'',info:(desc[j]||[]).slice(0,10),mat:dmat[j]||''})).filter(s=>s.info.length);
  if(notes.length)secs.unshift({type:'',info:notes.slice(0,6),mat:''});
  out.meta={...meta,skipped:!sawDate,past,dated,scanned:rows.length,tabMat:tabMat||'',secs};return out;
}
function fmtMin(m){const h=Math.floor(m/60)%24,mi=m%60;return `${h%12||12}:${String(mi).padStart(2,'0')} ${h<12?'AM':'PM'}`;}
// ---- ICS (Outlook published calendar or exported .ics)
function parseICS(text,horizonDays=90){
  const lines=text.replace(/\r?\n[ \t]/g,'').split(/\r?\n/),evs=[];let cur=null;
  const pd=v=>{const m=/^(\d{4})(\d\d)(\d\d)(?:T(\d\d)(\d\d)(\d\d)?)?(Z)?$/.exec(v||'');if(!m)return null;
    return m[7]?Date.UTC(+m[1],m[2]-1,+m[3],+m[4]||0,+m[5]||0,+m[6]||0):new Date(+m[1],m[2]-1,+m[3],+m[4]||0,+m[5]||0).getTime();};
  for(const l of lines){
    if(l==='BEGIN:VEVENT')cur={};else if(l==='END:VEVENT'&&cur){if(cur.s&&!cur.free)evs.push(cur);cur=null;}
    else if(cur){const i=l.indexOf(':');if(i<0)continue;const k=l.slice(0,i).split(';')[0],v=l.slice(i+1);
      if(k==='DTSTART'){cur.s=pd(v);cur.ad=!/T/.test(v);}else if(k==='DTEND')cur.e=pd(v);else if(k==='SUMMARY')cur.t=unesc(v).replace(/\\,/g,',');
      else if(k==='RRULE')cur.rr=v;else if(k==='TRANSP'&&v==='TRANSPARENT')cur.free=true;else if(k==='X-MICROSOFT-CDO-BUSYSTATUS'&&/FREE/i.test(v))cur.free=true;else if(k==='STATUS'&&v==='CANCELLED')cur.free=true;else if(k==='EXDATE')(cur.ex=cur.ex||[]).push(...v.split(',').map(x=>pd(x)));}
  }
  const now=Date.now(),lim=now+horizonDays*864e5,out=[];
  for(const e of evs){
    const dur=(e.e||e.s)-e.s,ex=new Set((e.ex||[]).filter(Boolean));
    if(!e.rr){out.push({s:e.s,e:e.s+dur,t:e.t||'busy',ad:!!e.ad});continue;}
    const R=Object.fromEntries(e.rr.split(';').map(x=>x.split('='))),iv=+(R.INTERVAL||1),until=R.UNTIL?pd(R.UNTIL):Infinity,cap=R.COUNT?+R.COUNT:Infinity;
    const wd=['SU','MO','TU','WE','TH','FR','SA'],sd=new Date(e.s),days=R.BYDAY?R.BYDAY.split(',').map(x=>x.slice(-2)):[wd[sd.getDay()]];
    let cnt=0;
    for(let d=0;d<=730&&cnt<cap;d++){
      const o=new Date(e.s);o.setDate(o.getDate()+d);const t=o.getTime();if(t>until||t>lim)break;
      let hit=false;
      if(R.FREQ==='DAILY')hit=d%iv===0;
      else if(R.FREQ==='WEEKLY')hit=Math.floor(d/7)%iv===0&&days.includes(wd[o.getDay()])&&(d>=0);
      else if(R.FREQ==='MONTHLY')hit=o.getDate()===sd.getDate()&&(((o.getFullYear()-sd.getFullYear())*12+o.getMonth()-sd.getMonth())%iv===0);
      if(!hit)continue;cnt++;
      if(!ex.has(t)&&t+dur>=now)out.push({s:t,e:t+dur,t:e.t||'busy',ad:!!e.ad});
    }
  }
  return out.filter(x=>x.e>=now-864e5);
}
function conflictsFor(t,events,defMin=60,ignoreAD=true){
  if(!t.start)return null;const end=t.end||t.start+defMin*6e4;
  return events.filter(e=>!(ignoreAD&&e.ad)&&e.s<end&&e.e>t.start);
}
